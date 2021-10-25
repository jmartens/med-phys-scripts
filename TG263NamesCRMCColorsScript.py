import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
from random import randint
from re import search, split, IGNORECASE

import numpy as np
import pandas as pd
from connect import *
from System.Drawing import Color
from System.Windows.Forms import *


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


def tg263_names_crmc_colors():
    """Rename and recolor ROIs in the current case, according to "TG263 Nomenclature with CRMC Colors.csv" spreadsheet
    
    Recolor target ROIs according to target type
    Ensure that ROI names are unique. E.g., if two ROIs need to be renamed to "Lungs", the first is "Lungs" and the second is "Lungs (1)". Duplicate colors are allowed (but aren't ideal), except for targets.
    If an ROI name is not already in the spreadsheet list of TG-263-compliant names, try to match it to a name in the "Possible Alternate Names" column
    If no match is found, write ROI name to "TG263NamesCRMCColors.txt" output file.
    For best results, an ROI should be in at most one "Possible Alternate Names" lists. If it is in multiple, choose the first from the top.
    """

    # Get current case
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)

    # Open file
    tg263 = pd.read_csv(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\TG263 Nomenclature with CRMC Colors.csv")
    
    approved_roi_names = set(geom.OfRoi.Name for ss in case.PatientModel.StructureSets for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures)

    # Targets
    target_types = ["CTV", "GTV", "ITV", "PTV"]
    targets = {target_type: [] for target_type in target_types}
    
    with open(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\TG263NamesCRMCColors\TG263NamesCRMCColors.txt", "a") as no_match:  # Output file
        for roi in case.PatientModel.RegionsOfInterest:
            if roi.Name in approved_roi_names:
                continue

            roi_name = split(" \(\d+\)$", roi.Name)[0]  # Remove copy number if it is present
            new_name = None
            
            # Add ROI to target types list if necessary
            for target_type in target_types:
                if roi.Type.upper() == target_type or search("(?<!\-){}".format(target_type), roi_name, IGNORECASE):  # ROI name contains target type, but not after a minus sign (e.g., "PTV 70 Gy" but not "Bladder-PTV")
                    targets[target_type].append(roi)  # Add to list of those targets

            # Find the matching TG-263 name
            # Check alternate names first due to special case w/ "Bowel" not getting renamed to "Bag_Bowel" b/c "Bowel" is also in the "TG263-Primary Name" column
            for _, row in tg263.iterrows():
                alt_names = [alt_name.lower() for alt_name in str(row["Possible Alternate Names"]).split("; ")]
                if roi_name.lower() in alt_names:
                    new_name = row["TG263-Primary Name"]

                    # Add ROI to target types list if necessary
                    if row["Target Type"] == "Target":
                        targets[row["Major Category"]].append(roi)
                    break
            
            # If it's not an alternate name, is it already TG-263 compliant?
            if row["Target Type"] != "Target" and new_name is None and roi_name in tg263["TG263-Primary Name"].values:
                new_name = roi_name
           
            # Account for left/right structure
            temp_name = new_name if new_name is not None else roi_name
            if not search("_[LR]$", temp_name):
                new_name_l = "{}_L".format(temp_name)
                new_name_r = "{}_R".format(temp_name)
                if new_name_l in tg263["TG263-Primary Name"].values and new_name_r in tg263["TG263-Primary Name"].values:  # Right and left names are valid
                    geom_on_left = [ss.RoiGeometries[roi.Name].GetCenterOfRoi().x > 0 for ss in case.PatientModel.StructureSets if ss.RoiGeometries[roi.Name].HasContours()]  # True if the geometry is on the patient's left, False otherwise
                    if geom_on_left:  # There are geometries
                        if all(geom_on_left):  # Geometry is on the left on all exams
                            new_name = new_name_l
                        elif not any(geom_on_left):  # Geometry is on the right on all exams
                            new_name = new_name_r

            # No match
            if new_name is None:
                no_match.write("{}\n".format(roi.Name))
                continue

            # Rename, and recolor non-target
            roi.Name = name_item(new_name, [r.Name for r in case.PatientModel.RegionsOfInterest if r.Name != roi.Name], 16)
            
            all_target_names = [target.Name for target_list in targets.values() for target in target_list]
            if roi.Name not in all_target_names:
                new_color = tg263.loc[tg263["TG263-Primary Name"] == new_name, "Color"].values[0]
                a, r, g, b = [int(component) for component in new_color.split("; ")]
                new_color = Color.FromArgb(a, r, g, b)
                roi.Color = new_color

        # Recolor targets and change type if necessary
        for target_type, rois in targets.items():
            target_color = tg263.loc[tg263["TG263-Primary Name"] == target_type.upper(), "Color"].values[0]
            a, r, g, b = [int(component) for component in target_color.split("; ")]
            # Evenly spaced shades from half to full opacity
            for i, roi in enumerate(rois):
                if len(rois) > 1:
                    a = 255 - int(128 / (len(rois) - 1)) * i
                roi.Type = target_type
                roi.Color = Color.FromArgb(a, r, g, b)

