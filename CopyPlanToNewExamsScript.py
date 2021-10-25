t_path = r"\\vs20filesvr01\groups\CANCER\Physics"


import clr
clr.AddReference("System.Windows.Forms")

import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))

from connect import *
from System.Windows.Forms import MessageBox

from CopyBeamSetScript import copy_beam_set


def get_tx_technique(beam_set):
    if beam_set.Modality == "Photons":
        if beam_set.PlanGenerationTechnique == "Imrt":
            if beam_set.DeliveryTechnique == "SMLC":
                return "SMLC"
            if beam_set.DeliveryTechnique == "DynamicArc":
                return "VMAT"
            if beam_set.DeliveryTechnique == "DMLC":
                return "DMLC"
        if beam_set.PlanGenerationTechnique == "Conformal":
            if beam_set.DeliveryTechnique == "SMLC":
                return "SMLC" # Changed from "Conformal". Failing with forward plans.
            if beam_set.DeliveryTechnique == "Arc":
                return "Conformal Arc"
    if beam_set.Modality == "Electrons":
        if beam_set.PlanGenerationTechnique == "Conformal":
            if beam_set.DeliveryTechnique == "SMLC":
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


def copy_plan_to_new_exams():
    """Copy current plan to all other exams in the same frame of reference as the planning exam

    Copy ROI and POI geometries, dose grid, beam sets (see CopyBeamSetScript.py), and Clinical Goals
    This script is designed for copying a plan to an exam that is a copy of the planning exam, so use for other purposes is at your own risk
    """

    try:
        plan = get_current("Plan")
    except:
        MessageBox.Show("There is no plan loaded. Click OK to abort the script.", "No Plan Loaded")
        sys.exit()
    case = get_current("Case")
    ext = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"][0]

    old_struct_set = plan.GetStructureSet()
    planning_exam = old_struct_set.OnExamination
    planning_exam_for = planning_exam.EquipmentInfo.FrameOfReference
    for new_exam in case.Examinations:
        if new_exam.Name == planning_exam.Name or new_exam.EquipmentInfo.FrameOfReference != planning_exam_for:
            continue

        new_struct_set = case.PatientModel.StructureSets[new_exam.Name]

        ## Copy ROI and POI geometries, if geometries don't already exist on new exam
        geom_names = [geom.OfRoi.Name for geom in old_struct_set.RoiGeometries if geom.HasContours() and not new_struct_set.RoiGeometries[geom.OfRoi.Name].HasContours()]
        if geom_names:
            case.PatientModel.CopyRoiGeometries(SourceExamination=planning_exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)

        # Update derived geometries
        for geom in new_struct_set.RoiGeometries:
            roi = case.PatientModel.RegionsOfInterest[geom.OfRoi.Name]
            if geom.HasContours() and roi.DerivedRoiExpression is not None and geom.PrimaryShape.DerivedRoiStatus is not None and not geom.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
                roi.UpdateDerivedGeometry(Examination=new_exam)

        # Copy POI geometries from old exam to new exam
        for i, poi in enumerate(old_struct_set.PoiGeometries):
            if abs(poi.Point.x) < 1000:
                new_struct_set.PoiGeometries[i].Point = poi.Point

        ## Create new plan on new exam
        new_plan_name = name_item(plan.Name, [p.Name for p in case.TreatmentPlans], 16)
        new_plan_comments = "'{}' copied to '{}'".format(plan.Name, new_exam.Name)
        new_plan = case.AddNewPlan(PlanName=new_plan_name, PlannedBy=plan.PlannedBy, Comment=new_plan_comments, ExaminationName=new_exam.Name, AllowDuplicateNames=False)
        
        # Copy dose grid
        dg = plan.GetDoseGrid()
        if not new_struct_set.RoiGeometries[ext.Name].HasContours():
            ext.CreateExternalGeometry(Examination=new_exam)
        new_plan.SetDefaultDoseGrid(VoxelSize=dg.VoxelSize)

        # Copy beam sets
        for beam_set in plan.BeamSets:
            copy_beam_set(old_beam_set_id=beam_set.BeamSetIdentifier(), new_plan_name=new_plan.Name)

    # Copy clinical goals
    for func in plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions:
        try:
            new_plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=func.ForRegionOfInterest.Name, GoalCriteria=func.PlanningGoal.GoalCriteria, GoalType=func.PlanningGoal.Type, AcceptanceLevel=func.PlanningGoal.AcceptanceLevel, ParameterValue=func.PlanningGoal.ParameterValue, IsComparativeGoal=func.PlanningGoal.IsComparativeGoal, Priority=func.PlanningGoal.Priority)
        except:  # Clinical goal already exists (e.g., previous run of this script was stopped prematurely)
            continue
        