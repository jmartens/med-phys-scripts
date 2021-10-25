import clr
clr.AddReference("System.Windows.Forms")
import sys

from connect import *  # Interact w/ RS
from System.Windows.Forms import MessageBox


def delete_empty_rois():
    """Delete all ROIs in the current case that are empty on all exams
    
    Alert the user of any empty ROIs that could be deleted because they are approved
    """
    
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error

    roi_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest]
    approved_roi_names = [geom.OfRoi.Name for ss in case.PatientModel.StructureSets for approved_ss in ss.ApprovedStructureSets for geom in approved_ss.ApprovedRoiStructures]
    approved_roi_names.extend([geom.OfRoi.Name for plan in case.TreatmentPlans for bs in plan.BeamSets if bs.DependentApprovedStructureSet is not None for geom in bs.DependentApprovedStructureSet.ApprovedRoiStructures])
    empty_approved_roi_names = []  # Names of ROIs that could not be deleted because they are approved
    with CompositeAction("Delete empty ROIs"):
        for roi_name in roi_names:
            if not any(ss.RoiGeometries[roi_name].HasContours() for ss in case.PatientModel.StructureSets):
                if roi_name in approved_roi_names:
                    empty_approved_roi_names.append(roi_name)
                else:
                    case.PatientModel.RegionsOfInterest[roi_name].DeleteRoi()
    
    # Alert user if any empty ROIs were not deleted
    if empty_approved_roi_names:
        msg = "The following ROIs could not be deleted because they are part of approved structure set(s): {}.".format(", ".join(empty_approved_roi_names))
        MessageBox.Show(msg, "Delete Empty ROIs")
