import clr
clr.AddReference("System.Windows.Forms")
import sys
from re import search

import pandas as pd
from connect import *

from System.Windows.Forms import MessageBox


case = exam = None


def create_roi_if_absent(roi_name, roi_type):
    # Helper function that returns the latest unapproved ROI (ROI whose name has the highest copy number) with the given name
    # If no such ROI exists, return a new ROI with the given name (made unique) and type
    #  
    # CRMC standard ROI colors
    colors = pd.read_csv(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\TG263 Nomenclature with CRMC Colors.csv", index_col="TG263-Primary Name", usecols=["TG263-Primary Name", "Color"])["Color"]

    roi = get_latest_roi(roi_name, unapproved_only=True)  # Latest unapproved ROI
    if roi is None:  # No such ROIs exist
        color = colors[roi_name].replace(";", ",")  # E.g., "255; 1; 2; 3" -> "255, 1, 2, 3"
        roi_name = name_item(roi_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)
        roi = case.PatientModel.CreateRoi(Name=roi_name, Type=roi_type, Color=color)  # Create a new ROI
    return roi


def get_latest_roi(base_roi_name, **kwargs):
    # Helper function that returns the ROI with the given "base name" and largest copy number
    # kwargs:
    # unapproved_only: If True, consider only the ROIs that are not part of any approved structure set in the case
    # non_empty_only: If True, consider only the ROIs with geometries on the exam

    unapproved_only = kwargs.get("unapproved_only", False)
    non_empty_only = kwargs.get("non_empty_only", False)

    base_roi_name = base_roi_name.lower()

    rois = case.PatientModel.RegionsOfInterest
    if unapproved_only:
        approved_roi_names = [geom.OfRoi.Name for approved_ss in case.PatientModel.StructureSets[exam.Name].ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures]
        approved_roi_names.extend([geom.OfRoi.Name for plan in case.TreatmentPlans for bs in plan.BeamSets if bs.DependentApprovedStructureSet is not None for geom in bs.DependentApprovedStructureSet.ApprovedRoiStructures])
        rois = [roi for roi in rois if roi.Name not in approved_roi_names]
    if non_empty_only:
        rois = [roi for roi in rois if case.PatientModel.StructureSets[exam.Name].RoiGeometries[roi.Name].HasContours()]
    
    latest_copy_num = copy_num = -1
    latest_roi = None
    for roi in rois:
        roi_name = roi.Name.lower()
        if roi_name == base_roi_name:
            copy_num = 0
        else:
            m = search(" \((\d+)\)".format(base_roi_name), roi_name)
            if m:  # There is a copy number
                grp = m.group()
                length = min(16 - len(grp), len(base_roi_name))
                if roi_name[:length] == base_roi_name[:length]:
                    copy_num = int(m.group(1))
        if copy_num > latest_copy_num:
            latest_copy_num = copy_num
            latest_roi = roi
    return latest_roi


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


def add_box_to_external():
    """Add a box to the external geometry on the current exam

    Overwrite the existing external geometry
    If the external geometry is approved, create a new organ ROI to hold the external plus the box
    Box is as wide as the inner couch geometry, its height extends from the top of the couch to the loc point, and its depth is the depth of the external geometry.
    Useful for including the VacLock for SBRT lung plans

    Assumptions
    -----------
    Outer and inner couch structures are support structures and have respective materials cork and PMI foam.
    """

    global case, exam

    # Get current objects
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)
    try:
        exam = get_current("Examination")
    except:
        MessageBox.Show("There are no examinations in the current case. Click OK to abort script.", "No Examinations")
        sys.exit(1) 
    ss = case.PatientModel.StructureSets[exam.Name]

    # External
    ext = [geom for geom in ss.RoiGeometries if geom.OfRoi.Type == "External"]
    if not ext:
        MessageBox.Show("There is no external geometry on the current exam. Click OK to abort script.", "No External Geometry")
        sys.exit(1)
    ext = ext[0]  # There should be only 1 external

    inner_couch = [geom for geom in ss.RoiGeometries if geom.OfRoi.Type == "Support" and geom.OfRoi.RoiMaterial.OfMaterial.Name == "PMI foam"]
    outer_couch = [geom for geom in ss.RoiGeometries if geom.OfRoi.Type == "Support" and geom.OfRoi.RoiMaterial.OfMaterial.Name == "Cork"]
    if not inner_couch or not outer_couch:
        MessageBox.Show("Current exam is missing couch geometry(ies). Click OK to abort script.", "No Couch Geometry(ies)")
        sys.exit(1)
    inner_couch, outer_couch = inner_couch[0], outer_couch[0]

    # RL width = exam width minus a px on each side (RS throws error if external is as wide as image)
    px_sz = exam.Series[0].ImageStack.PixelSize.x
    exam_bb = exam.Series[0].ImageStack.GetBoundingBox()
    exam_min_x, exam_max_x = exam_bb[0].x + px_sz, exam_bb[1].x - px_sz
    couch_bb = inner_couch.GetBoundingBox()
    couch_min_x, couch_max_x = couch_bb[0].x, couch_bb[1].x
    x = min(couch_max_x, exam_max_x) - max(couch_min_x, exam_min_x)

    # PA height
    loc_y = ss.LocalizationPoiGeometry
    if loc_y is None or abs(loc_y.Point.y) > 1000:  # No localization point, or localization geometry has infinite coordinates
        MessageBox.Show("There is no localization geometry on the current exam.", "No Localization Geometry")
        sys.exit(1)
    loc_y = loc_y.Point.y  # Loc point
    couch_y = outer_couch.GetBoundingBox()[exam.PatientPosition.endswith("P")].y  # Top of couch is larger y for prone patient, smaller y for supine
    y = abs(couch_y - loc_y)

    # IS height = height of External
    ext_min_z = ext.GetBoundingBox()[0].z
    ext_max_z = ext.GetBoundingBox()[1].z
    z = ext_max_z - ext_min_z

    # Box center (x center = 0)
    z_ctr = (ext_max_z + ext_min_z) / 2
    y_ctr = (couch_y + loc_y) / 2  

    # Create box geometry
    box = create_roi_if_absent("Box", "Control")
    box.CreateBoxGeometry(Size={ "x": x, "y": y, "z": z }, Examination=exam, Center={ "x": 0, "y": y_ctr, "z": z_ctr })
    
    # Copy external into new ROI "External^NoBox"
    ext_no_box = create_roi_if_absent("External^NoBox", "Organ")
    ext_no_box.CreateMarginGeometry(Examination=exam, SourceRoiName=ext.OfRoi.Name, MarginSettings={ "Type": "Expand", "Superior": 0, "Inferior": 0, "Anterior": 0, "Posterior": 0, "Right": 0, "Left": 0 })
    
    # Add box to external
    approved_roi_names = [geom.OfRoi.Name for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures]
    approved_roi_names.extend([geom.OfRoi.Name for plan in case.TreatmentPlans for bs in plan.BeamSets if bs.DependentApprovedStructureSet is not None for geom in bs.DependentApprovedStructureSet.ApprovedRoiStructures])
    if ext.OfRoi.Name in approved_roi_names:
        unapproved_ext = create_roi_if_absent("External", "Organ")
        unapproved_ext.CreateMarginGeometry(Examination=exam, SourceRoiName=ext.OfRoi.Name, MarginSettings={ "Type": "Expand", "Superior": 0, "Inferior": 0, "Anterior": 0, "Posterior": 0, "Right": 0, "Left": 0 })
        ext = unapproved_ext
    else:
        ext = ext.OfRoi
    ext.CreateAlgebraGeometry(Examination=exam, ExpressionA={ "Operation": "Union", "SourceRoiNames": [ext.Name], "MarginSettings": { "Type": "Expand", "Superior": 0, "Inferior": 0, "Anterior": 0, "Posterior": 0, "Right": 0, "Left": 0 } }, ExpressionB={ "Operation": "Union", "SourceRoiNames": [box.Name], "MarginSettings": { "Type": "Expand", "Superior": 0, "Inferior": 0, "Anterior": 0, "Posterior": 0, "Right": 0, "Left": 0 } }, ResultOperation="Union", ResultMarginSettings={ "Type": "Expand", "Superior": 0, "Inferior": 0, "Anterior": 0, "Posterior": 0, "Right": 0, "Left": 0 })

    # Delete unnecessary ROI
    box.DeleteRoi()