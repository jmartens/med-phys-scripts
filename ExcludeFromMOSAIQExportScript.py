# For GUI (MessageBoxes display errors)
import clr
clr.AddReference("System.Windows.Forms")

import os
import sys

from connect import *  # Interact w/ RS
from System.Windows.Forms import *  # For GUI


def exclude_from_mosaiq_export():
    """Include/exclude the current plan's ROIs from export
    Make visible all included ROIs and makes invisible all excluded ROIs

    Include and make visible:
        - SpinalCord
        - Targets
        - Support ROIs
        - External
        - Any otherwise excluded ROIs with max dose >= 10 Gy

    Exclude and make invisible:
        - ROI named "box" (case insensitive)
        - Control ROIs
        - ROIs with empty geometries on planning exam
        - Any otherwise included ROIs with max dose < 10 Gy

    Assume all ROI names are TG-263 compliant
    
    Upon script completion, the user should review the included/excluded ROIs in ROI/POI Details and make any necessary changes
    """

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
    try:
        plan = get_current("Plan")
    except:
        MessageBox.Show("There is no plan loaded. Click OK to abort script.", "No Plan Loaded")
        sys.exit(1)

    dose_dist = plan.TreatmentCourse.TotalDose

    # Lists of OARs to include, other structures to include, and all structures to exclude
    # Separate OARs and other structures so we can remove OAR dependencies 
    include_organs, include_other, exclude = [], [], []  
    
    ss = plan.GetStructureSet()
    for geom in ss.RoiGeometries:
        roi_name = geom.OfRoi.Name
        if roi_name.lower() == "box":
            exclude.append(roi_name)
        if geom.OfRoi.Type == "Control" or not geom.HasContours():  # ROI is a control or geometry has no contours
            exclude.append(roi_name)
        elif roi_name == "SpinalCord":
            include_organs.append(roi_name)
        elif geom.OfRoi.OrganData.OrganType == "Target" or geom.OfRoi.Type in ["Support", "External"]:
            include_other.append(roi_name)
        else:  # include/exclude determined by max dose @ appr vol
            rel_vol = 0.035 / geom.GetRoiVolume()
            max_dose = dose_dist.GetDoseAtRelativeVolumes(RoiName=roi_name, RelativeVolumes=[rel_vol])[0]
            if max_dose >= 1000:
                include_organs.append(roi_name)
            else:
                exclude.append(roi_name)

    # Remove dependent ROIs  
    dependent_rois = [ss.RoiGeometries[geom_name].GetDependentRois() for geom_name in include_organs]
    include_organs = [geom for geom in include_organs if geom not in dependent_rois]
    include = include_organs + include_other
 
    # Include/exclude from export and set visibility
    with CompositeAction("Apply ROI changes"):
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=False, RegionOfInterests=include)
        for roi in include:
            patient.SetRoiVisibility(RoiName=roi, IsVisible=True)
        case.PatientModel.ToggleExcludeFromExport(ExcludeFromExport=True, RegionOfInterests=exclude)
        for roi in exclude:
            patient.SetRoiVisibility(RoiName=roi, IsVisible=False)
