import clr
clr.AddReference("System.Windows.Forms")
from math import ceil
from os import listdir, makedirs, write
from os.path import isdir
from shutil import copy2, rmtree
import sys
sys.path.append(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\RayStation")

from connect import *
from datetime import datetime
from pydicom import dcmread
from System.Windows.Forms import *


case = None


def compute_new_ids(exam):
    # Helper function that creates a unique DICOM StudyInstanceUID and SeriesInstanceUID for an examination
    # Return a 2-tuple of (new StudyInstanceUID, new SeriesInstanceUID)
    
    dcm = exam.GetAcquisitionDataFromDicom()
    new_ids = []
    for id_type in ["Study", "Series"]:
        module_key = "{}Module".format(id_type)
        id_key = "{}InstanceUID".format(id_type)
        all_ids = [e.GetAcquisitionDataFromDicom()[module_key][id_key] for e in case.Examinations]
        new_id = dcm[module_key][id_key]
        while new_id in all_ids:
            dot_idx = new_id.rfind(".")
            new_id = "{}.{}".format(new_id[:dot_idx], int(new_id[(dot_idx + 1):]) + 1)
        new_ids.append(new_id)
    return new_ids


def copy_dicom_files(export_path, exam, dist, sup=True):
    # Helper function that copies the inferior or superior slice's DICOM file at `export_path` according to the expansion (`dist`) needed
    # `export_path`: Absolute path to the exported DICOM files for this run of the script
    # `exam`: The examination to be extended
    # `dist`: The distance to extend (cm)
    # `sup`: True if exam should be extended in superior direction, False for inferior
    # Set SOP Instance UID (also use for filename), Slice Location, z-coordinate of Image Position (Patient), and Instance Number in the copied slice DICOM files
    # Return the distance by which the exam was expanded

    # Info for computing slice UIDs for new filenames
    # In RS, slices are ordered superior to inferior, smallest ID to largest
    if sup:
        slice_id = exam.Series[0].ImageStack.ImportedDicomSliceUIDs[-1]  # E.g., "1.2.840.113704.1.111.2528.1583439123.410"
    else:  # Inf
        slice_id = list(exam.Series[0].ImageStack.ImportedDicomSliceUIDs)[0]
    
    # Split slice ID into first part (used for all new slice IDs) and second part (incremented/decremented for each new slice ID)
    dot_idx = slice_id.rfind(".")
    part_1, part_2 = slice_id[:dot_idx], int(slice_id[(dot_idx + 1):])

    # Get DICOM data for top (for sup) or bottom (for inf) slice
    dcm_filepath = r"{}\CT{}.dcm".format(export_path, slice_id)
    dcm = dcmread(dcm_filepath)

    # Slice data
    slice_thickness = dcm.SliceThickness
    slice_loc = dcm.SliceLocation
    num_copies = ceil((5 - dist) * 10 / slice_thickness)  # Convert to mm

    # `copy_instance_num` = InstanceNumber for next slice
    if sup:
        copy_instance_num = len(listdir(export_path))  # We will increment the largest slice ID (see above)
    else:
        copy_instance_num = num_copies + 1  # We will decrement the smallest slice ID (see above)

    # Make `num_copies` copies of top slice
    for _ in range(num_copies):
        if sup:
            part_2 += 1  # Last part of slice UID (filename) is 1 more than previous slice (filename)
            copy_instance_num += 1
            slice_loc += slice_thickness
        else:
            part_2 -= 1  # Last part of slice UID (filename) is 1 less than previous slice (filename)
            copy_instance_num -= 1
            slice_loc -= slice_thickness

        # Copy DICOM file to appropriate new filename
        instance_id = "{}.{}".format(part_1, part_2)
        copy_dcm_filepath = r"{}\CT{}.dcm".format(export_path, instance_id) 
        copy2(dcm_filepath, copy_dcm_filepath)

        # Change instance data in new slice
        copy_dcm = dcmread(copy_dcm_filepath)
        copy_dcm.SOPInstanceUID = instance_id
        copy_dcm.SliceLocation = slice_loc
        copy_dcm.ImagePositionPatient[2] = slice_loc
        copy_dcm.InstanceNumber = copy_instance_num

        # Overwrite copied DICOM file
        copy_dcm.save_as(copy_dcm_filepath)

    # Renumber the old instances if we added slices to the beginning
    if not sup:
        copy_instance_num = num_copies + 1
        for f in listdir(export_path)[num_copies:]:  # Ignore the copies, as they already have the correct instance number
            f = r"{}\{}".format(export_path, f)
            dcm = dcmread(f)
            dcm.InstanceNumber = copy_instance_num
            dcm.save_as(f)
            copy_instance_num += 1


def get_tx_technique(bs):
    # Helper function that returns a beam set's treatment technique
    # If technique is not recognized, return None
    # Modified from a function written by RaySearch support
    
    if bs.Modality == "Photons":
        if bs.PlanGenerationTechnique == "Imrt":
            if bs.DeliveryTechnique == "SMLC":
                return "SMLC"
            if bs.DeliveryTechnique == "DynamicArc":
                return "VMAT"
            if bs.DeliveryTechnique == "DMLC":
                return "DMLC"
        if bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                return "SMLC" # Changed from "Conformal". Failing with forward plans.
            if bs.DeliveryTechnique == "Arc":
                return "Conformal Arc"
    if bs.Modality == "Electrons":
        if bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                return "ApplicatorAndCutout"


def name_item(item, l, max_len=sys.maxsize):
    # Helper function that generates a unique name for `item` in list `l` (case insensitive)
    # Limit name to `max_len` characters
    # E.g., name_item("Isocenter Name A", ["Isocenter Name A", "Isocenter Na (1)", "Isocenter N (10)"]) -> "Isocenter Na (2)"

    l_lower = [l_item.lower() for l_item in l]
    copy_num = 0
    old_item = item
    while item.lower() in l_lower:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
    return item[:max_len]


def extend_5_cm():
    """Extend the current exam so that the target is at least 5 cm from the superior and inferior edges of the exam

    If a beam set is loaded and its Rx is to a non-empty target, use that target. If not, use the first non-empty PTV, GTV, or CTV (checked in that order) found on the exam
    Export exam, copy the top and/or bottom slice(s) so that target is far enough from the edges, and reimport
    New exam name is old exam name plus " - Expanded" (possibly with a copy number - e.g., "SBRT Lung_R - Extended (2)")
    Copy ROI and POI geometries from old exam to new exam
    Do not copy any plans to new exam
    """

    global case

    # Get current objects
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)
    try:
        exam = get_current("Examination")
    except:
        MessageBox.Show("There are no exams in the current case. Click OK to abort the script.", "No Exams")
        sys.exit(1)
    patient_db = get_current("PatientDB")
    patient = get_current("Patient")
    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Find the target
    target = None  # Assume no targets on exam
    try:
        rx_struct = get_current("BeamSet").Prescription.PrimaryDosePrescription.OnStructure
        if rx_struct.OrganData.OrganType == "Target" and struct_set.RoiGeometries[rx_struct.Name].HasContours():
            target = struct_set.RoiGeometries[rx_struct.Name]
    except:
        # Find first PTV, GTV, or CTV (in that order) contoured on the exam
        for target_type in ["PTV", "GTV", "CTV"]:
            targets = [geom for geom in struct_set.RoiGeometries if geom.OfRoi.Type.upper() == target_type and geom.HasContours()]
            if targets:
                target = targets[0]
                break
    
    # If no targets contoured on exam, alert user and exit script with an error
    if target is None:
        MessageBox.Show("There are no target geometries on the current exam. Click OK to abort the script.", "No Target Geometries")
        sys.exit(1)

    # Exam bounds
    img_bounds = exam.Series[0].ImageStack.GetBoundingBox()
    img_inf, img_sup = img_bounds[0].z, img_bounds[1].z

    # Target bounds
    target_bounds = target.GetBoundingBox()
    target_inf, target_sup = target_bounds[0].z, target_bounds[1].z

    # Distance from target to inf and sup exam edges
    inf_dist = abs(img_inf - target_inf)
    sup_dist = abs(img_sup - target_sup)

    if inf_dist < 5 or sup_dist < 5:  # Target is < 5 cm from top or bottom of exam
        # Create export directory
        base_path = r"\\vs20filesvr01\groups\CANCER\Physics\Temp\Extend2CmScript"
        export_path = r"{}\{}".format(base_path, datetime.now().strftime("%m-%d-%Y %H_%M_%S"))
        makedirs(export_path)
        
        # Export exam
        # Note that we could also export the structure set, 
        # but there is no way to access a structure set's UID from RS, 
        # and exporting every structure set so we could get the UID from the DICOM would unnecessary slow down the script.
        # Instead, after importing the new exam (later), we simply copy all ROI and POI geometries from the old exam to the new exam
        patient.Save()  # Error if you attempt to export when there are unsaved modifications
        try:
            case.ScriptableDicomExport(ExportFolderPath=export_path, Examinations=[exam.Name], IgnorePreConditionWarnings=False)
        except:
            case.ScriptableDicomExport(ExportFolderPath=export_path, Examinations=[exam.Name], IgnorePreConditionWarnings=True)

        # Compute new study and series IDs so RS doesn't think the new exam is the same as the old
        study_id, series_id = compute_new_ids(exam)

        # Add slices to top, if necessary
        if sup_dist < 5:
            copy_dicom_files(export_path, exam, sup_dist)
        
        # Add slices to top, if necessary
        if inf_dist < 5:
            copy_dicom_files(export_path, exam, inf_dist, False)

        # Change study and series UIDs in all files so RS doesn't think the new exam is the same as the old
        for f in listdir(export_path):
            f = r"{}\{}".format(export_path, f)  # Absolute path
            dcm = dcmread(f)
            dcm.StudyInstanceUID = study_id
            dcm.SeriesInstanceUID = series_id
            dcm.save_as(f)

        # Import new exam
        study = patient_db.QueryStudiesFromPath(Path=export_path, SearchCriterias={"PatientID": patient.PatientID})[0]  # There is only one study in the directory
        series = patient_db.QuerySeriesFromPath(Path=export_path, SearchCriterias=study)  # Series belonging to the study
        patient.ImportDataFromPath(Path=export_path, CaseName=case.CaseName, SeriesOrInstances=series)
        
        # Select new exam
        new_exam = [e for e in case.Examinations if e.Series[0].ImportedDicomUID == series_id][0]
        new_exam.Name = name_item("{} - Extended".format(exam.Name), [e.Name for e in case.Examinations])

        # Set new exam imaging system
        if exam.EquipmentInfo.ImagingSystemReference:
            new_exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=exam.EquipmentInfo.ImagingSystemReference.ImagingSystemName)

        # Add external geometry to new exam
        ext = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]
        if ext:
            ext = ext[0]
        else:
            ext_name = name_item("External", [roi.Name for roi in case.PatientModel.RegionsOfInterest], 16)
            ext = case.PatientModel.CreateRoi(Name=ext_name, Color="255, 255, 255", Type="External")
        ext.CreateExternalGeometry(Examination=new_exam)

        # Copy ROI geometries from old exam to new exam
        geom_names = [geom.OfRoi.Name for geom in case.PatientModel.StructureSets[exam.Name].RoiGeometries if geom.HasContours() and geom.OfRoi.Type != "External"]
        if geom_names:
            case.PatientModel.CopyRoiGeometries(SourceExamination=exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)

        # Update derived geometries (this shouldn't change any geometries since the new exam is effectively the same as the old)
        for geom in case.PatientModel.StructureSets[exam.Name].RoiGeometries:
            roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
            if geom.OfRoi.DerivedRoiExpression and geom.PrimaryShape.DerivedRoiStatus and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
                roi.UpdateDerivedGeometry(Examination=new_exam)

        # Copy POI geometries from old exam to new exam
        for i, poi in enumerate(case.PatientModel.StructureSets[exam.Name].PoiGeometries):
            if abs(poi.Point.x) < 1000:  # Empty POI geometry if infinite coordinates
                case.PatientModel.StructureSets[new_exam.Name].PoiGeometries[i].Point = poi.Point

        # Delete the temporary directory and all its contents
        rmtree(export_path)

    else:
        MessageBox.Show("The target ('{}') is at least 5 cm from the inferior and superior edges of the planning exam. No action is necessary.".format(target.OfRoi.Name), "Exam OK")
