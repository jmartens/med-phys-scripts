# For GUI (MessageBoxes display errors)
import clr
clr.AddReference("System.Windows.Forms")
import sys

from datetime import datetime
from os import listdir, makedirs
from shutil import rmtree  # Easy function for deleting nonempty directories

import pydicom  # Manipulate DICOM files/data
from connect import *  # Interact w/ RS
from System.Windows.Forms import *  # GUI


case = None


def compute_new_id(study_or_series):
    # Helper function that creates a unique DICOM StudyInstanceUID or SeriesInstanceUID for an examination
    # `study_or_series`: either "Study" or "Series"
    
    all_ids = [e.GetAcquisitionDataFromDicom()["{}Module".format(study_or_series)]["{}InstanceUID".format(study_or_series)] for e in case.Examinations]
    all_ids.sort(key=lambda id: int(id.split(".")[-1]))
    dot_idx = all_ids[-1].rfind(".")
    return "{}{}".format(all_ids[-1][:(dot_idx + 1)], int(all_ids[-1][(dot_idx + 1):]) + 1)


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


def copy_exam(exam_name=None):
    """Copy the given exam, including all structure sets on that exam

    Note that new exam "Used for" is "Evaluation" regardless of copied exam "Used for".

    Parameters
    ----------
    exam_name: str
        Name of the exam to copy
        Defaults to the current (primary) exam

    New exam name is old exam name plus " - Copy", possibly with a copy number (e.g., "Breast 1/1/21 - Copy (1)")
    New exam has a new study UID and series ID
    """

    global case

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort script.", "No Patient Loaded")
        sys.exit(1)
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)
    patient_db = get_current("PatientDB")

    # Get exam
    if exam_name is None:
        exam = get_current("Examination")
    else:
        try:
            exam = case.Examinations[exam_name]
        except:
            raise ValueError("There is no examination '{}' in the current case.")

    # Create empty directory for exported DICOM files
    base_path = r"\\vs20filesvr01\groups\CANCER\Physics\Temp\CopyExamScript"
    folder = r"{}\{}".format(base_path, datetime.now().strftime("%m-%d-%Y %H_%M_%S"))
    makedirs(folder)  # Create directory

    # Export exam and beam sets
    patient.Save()  # Error if you attempt to export when there are unsaved modifications
    export_args = {"ExportFolderPath": folder, "Examinations": [exam.Name], "IgnorePreConditionWarnings": False}
    try:
        case.ScriptableDicomExport(**export_args)
    except:
        export_args["IgnorePreConditionWarnings"] = True
        case.ScriptableDicomExport(**export_args)  # Retry the export, ignoring warnings

    # Compute new study and series IDs
    new_study_id = compute_new_id("Study")
    new_series_id = compute_new_id("Series")

    for f in listdir(folder):
        f = r"{}\{}".format(folder, f)  # Absolute path to file
        dcm = pydicom.dcmread(f)  # Read DICOM data
        dcm.StudyInstanceUID = new_study_id
        dcm.SeriesInstanceUID = new_series_id
        dcm.save_as(f, write_like_original=False)  # Overwrite original DICOM file

    # Find and import the edited DICOM files
    study = patient_db.QueryStudiesFromPath(Path=folder, SearchCriterias={"PatientID": patient.PatientID})[0]  # There is only one study in the directory
    series = patient_db.QuerySeriesFromPath(Path=folder, SearchCriterias=study)  # Series belonging to the study
    patient.ImportDataFromPath(Path=folder, SeriesOrInstances=series, CaseName=case.CaseName)  # Import into current case

    # The exam that was just imported
    new_exam = [e for e in case.Examinations if e.Series[0].ImportedDicomUID == new_series_id][0]  
    
    # Rename new exam
    new_exam.Name = name_item("{} - Copy".format(exam.Name), [e.Name for e in case.Examinations])

    # Set new exam imaging system
    if exam.EquipmentInfo.ImagingSystemReference:
        new_exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName=exam.EquipmentInfo.ImagingSystemReference.ImagingSystemName)

    # Copy ROI geometries from old exam to new exam
    geom_names = [geom.OfRoi.Name for geom in case.PatientModel.StructureSets[exam.Name].RoiGeometries if geom.HasContours()]
    if geom_names:
        case.PatientModel.CopyRoiGeometries(SourceExamination=exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)

    # Update derived geometries (this shouldn't change any geometries since the new exam is the same as the old)
    for geom in case.PatientModel.StructureSets[exam.Name].RoiGeometries:
        roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
        if geom.HasContours() and geom.OfRoi.DerivedRoiExpression and geom.PrimaryShape.DerivedRoiStatus and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
            roi.UpdateDerivedGeometry(Examination=new_exam)

    # Copy POI geometries from old exam to new exam
    for i, poi in enumerate(case.PatientModel.StructureSets[exam.Name].PoiGeometries):
        if abs(poi.Point.x) < 1000:
            case.PatientModel.StructureSets[new_exam.Name].PoiGeometries[i].Point = poi.Point

    # Delete the temporary directory and all its contents
    rmtree(folder)

    return new_exam.Name
