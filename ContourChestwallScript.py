import clr
clr.AddReference("System.Windows.Forms")
import sys
from random import randint
from re import search

import pandas as pd

from connect import *
from System.Windows.Forms import MessageBox


case = exam = None


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
        approved_roi_names = set(geom.OfRoi.Name for approved_ss in case.PatientModel.StructureSets[exam.Name].ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures)
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
            m = search(" \((\d+)\)$".format(), roi_name)
            if m:  # There is a copy number
                grp = m.group()
                length = 16 - len(grp)
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


def contour_chestwall():
    """Create Chestwall_L and Chestwall_R geometries on the current examination

    Based on lung expansions
    If an unapproved chestwall ROI exists, overwrite its geometry. Otherwise, create a new chestwall ROI.

    Assumptions
    -----------
    Left and right lung ROI names are TG-263 compliant.
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

    # Determine external ROI name
    try:
        case.PatientModel.RegionsOfInterest["External^NoBox"]
        ext_name = "External^NoBox"
    except:
        ext_name = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]
        if not ext_name:
            MessageBox.Show("There is no external geometry on the current exam. Click OK to abort the script.", "No External Geometry")
            sys.exit(1)
        ext_name = ext_name[0]

    # Get TG-263 colors
    colors = pd.read_csv(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\TG263 Nomenclature with CRMC Colors.csv", index_col="TG263-Primary Name", usecols=["TG263-Primary Name", "Color"])["Color"]

    chestwall_names = []
    # Create chestwall contours
    for side in ["L", "R"]:
        # Select lung
        lung_name = "Lung_{}".format(side)
        lung = get_latest_roi(lung_name, non_empty_only=True)
        if lung is None:
            continue

        # Chestwall name and color
        chestwall_name = "Chestwall_{}".format(side)  # "Chestwall_L" or "Chestwall_R"
        chestwall_color = colors[chestwall_name].replace(";", ",")  # E.g., "255; 1; 2; 3" -> "255, 1, 2, 3"
        
        # Do any chestwall ROIs already exist?
        chestwall = get_latest_roi(chestwall_name, unapproved_only=True)
        if chestwall is None:
            chestwall_name = name_item(chestwall_name, [roi.Name for roi in case.PatientModel.RegionsOfInterest], 16)  # Unique name for new chestwall ROI
            chestwall = case.PatientModel.CreateRoi(Name=chestwall_name, Type="Organ", Color=chestwall_color)  # Create new chestwall ROI  

        # Create chestwall geometry based on lung
        margin_a = { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 2, 'Posterior': 2 }  # Expand lung posteriorly, anteriorly, and in the `side` direction
        margin_b = { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0 }  # Expand lung oppsite the `side` direction
        if side == "L":
            margin_a["Left"] = margin_b["Right"] = 2
            margin_a["Right"] = margin_b["Left"] = 0
        else:
            margin_a["Right"] = margin_b["Left"] = 2
            margin_a["Left"] = margin_b["Right"] = 0
        chestwall.CreateAlgebraGeometry(Examination=exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [lung.Name], 'MarginSettings': margin_a }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [lung.Name], 'MarginSettings': margin_b }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
 
        # Remove parts of chestwall that extend outside the external
        chestwall.CreateAlgebraGeometry(Examination=exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [chestwall.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [ext_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Intersection", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
        chestwall_names.append(chestwall.Name)

    if not chestwall_names:
        MessageBox.Show("There are no right or left lung geometries on the current exam. Click OK to abort the script.", "No Lung Geometries")
        sys.exit(1)
    
    if len(chestwall_names) == 2:  # 2 chestwalls were added
        case.PatientModel.StructureSets[exam.Name].SimplifyContours(RoiNames=chestwall_names, RemoveHoles3D=True, RemoveSmallContours=True, AreaThreshold=0.1, ResolveOverlappingContours=True)
        # For some reason, the above does not remove chestwall overlap, so subtract one chestwall from the other
        minuend_idx = randint(0, 1)
        minuend = case.PatientModel.RegionsOfInterest[chestwall_names[minuend_idx]]
        subtrahend_name = chestwall_names[not minuend_idx]
        minuend.CreateAlgebraGeometry(Examination=exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [minuend.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [subtrahend_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
