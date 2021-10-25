# For GUI (MessageBox displays errors)
import clr
clr.AddReference("System.Windows.Forms")

from connect import *  # Interact w/ RS
from System.Windows.Forms import MessageBox  # I use to display errors


def show_nonempty_geometries():
    """Show (make visible) all ROIs and POIs with geometries on the current exam
    Hide (make invisible) all ROIs and POIs empty on the current exam
    """
    
    # Get current objects
    try:
        case = get_current("Case")
        try:
            exam = get_current("Examination")
        except:
            MessageBox.Show("There are no exams in the current case.", "Click OK to abort the script.", "No Examinations")
            sys.exit(1)
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)
    patient = get_current("Patient")
    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Show or hide each ROI
    with CompositeAction("Show/Hide ROIs"):
        for roi in case.PatientModel.RegionsOfInterest:
            visible = struct_set.RoiGeometries[roi.Name].HasContours()
            patient.SetRoiVisibility(RoiName=roi.Name, IsVisible=visible)

    # Show or hide each POI
    with CompositeAction("Show/Hide POIs"):
        for poi in case.PatientModel.PointsOfInterest:
            visible = abs(struct_set.PoiGeometries[poi.Name].Point.x) < 1000  # Infinite coordinates indicate empty geometry
            patient.SetPoiVisibility(PoiName=poi.Name, IsVisible=visible)
