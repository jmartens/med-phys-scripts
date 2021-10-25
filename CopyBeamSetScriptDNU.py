t_path = r"\\vs20filesvr01\groups\CANCER\Physics"  # T: drive


# For GUI (MessageBox displays errors)
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))

from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *

from CopyPlanWithoutChangesScript import copy_plan_without_changes


case = exam = None


class ChooseNewPlansAndExamsForm(Form):
    def __init__(self, old_plan_name):
        print("init")
        self.Text = "Plan(s)/Exam(s) to Copy Beam Set to"  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        y = 15  # Vertical coordinate of next control

        # Add plans checkboxes
        lbl = Label()
        lbl.AutoSize = True
        lbl.Location = Point(15, y)
        lbl.Text = "Choose the plan(s) to which to copy the beam set.\nPlans with an asterisk are approved, so the beam set will be added to a copy of the plan instead of the original."
        self.Controls.Add(lbl)
        y += lbl.Height + 5

        self.plans_gb = GroupBox()
        self.plans_gb.AutoSize = True
        self.plans_gb.Location = Point(15, y)
        self.plans_gb.Text = "Plan(s) to copy to:"
        cb_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = "Select all"
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        plan_names = sorted(p.Name if p.Review is None or p.Review.ApprovalStatus != "Approved" else "*{}".format(p.Name) for p in case.TreatmentPlans)
        if len(plan_names) == 1:
            cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        else:
            cb.CheckState = cb.Tag = CheckState.Indeterminate
        cb.Click += self.select_all_clicked
        self.plans_gb.Controls.Add(cb)
        cb_y += cb.Height
        for plan_name in plan_names:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)
            cb.Checked = plan_name == old_plan_name
            cb.Click += self.checkbox_clicked
            cb.Text = plan_name
            self.plans_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.plans_gb.Height = len(plan_names) * 20
        self.Controls.Add(self.plans_gb)
        y += self.plans_gb.Height + 15

        # Add exams checkboxes
        lbl = Label()
        lbl.AutoSize = True
        lbl.Location = Point(15, y)
        lbl.Text = "Choose the exam(s) to which to copy the beam set.\nA new plan to hold the copied beam set will be created on each exam."
        self.Controls.Add(lbl)
        y += lbl.Height + 5

        self.exams_gb = GroupBox()
        self.exams_gb.AutoSize = True
        self.exams_gb.Location = Point(15, y)
        self.exams_gb.Text = "Plan(s) to copy to:"
        cb_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = "Select all"
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        exam_names = sorted(e.Name for e in case.Examinations)
        cb.CheckState = cb.Tag = CheckState.Unchecked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        cb.Click += self.select_all_clicked
        self.exams_gb.Controls.Add(cb)
        cb_y += cb.Height
        for exam_name in exam_names:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)
            cb.Checked = False
            cb.Click += self.checkbox_clicked
            cb.Text = exam_name
            self.exams_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.exams_gb.Height = len(exam_names) * 20
        self.Controls.Add(self.exams_gb)
        y += self.exams_gb.Height + 15

        # Add "OK" button
        self.ok_btn = Button()
        self.ok_btn.AutoSize = True
        self.ok_btn.Click += self.ok_clicked
        self.ok_btn.Location = Point(15, y)
        self.ok_btn.Text = "Copy Beam Set"
        self.AcceptButton = self.ok_btn
        self.Controls.Add(self.ok_btn)

        self.ShowDialog()  # Launch window

    def select_all_clicked(self, sender, event):
        # Helper method that sets the check state of the select all checkbox and all other checkboxes

        if sender.Tag == CheckState.Checked:  # If checked, uncheck
            sender.CheckState = sender.Tag = CheckState.Unchecked
            for cb in list(sender.Parent.Controls)[1:]:  # The first checkbox is select all
                cb.Checked = False
        else:  #If unchecked or indeterminate, check
            all_enabled = True
            for cb in list(sender.Parent.Controls)[1:]:
                if cb.Enabled:
                    cb.Checked = True
                else:
                    all_enabled = False
            if all_enabled:
                sender.CheckState = sender.Tag = CheckState.Checked   
            else:
                sender.CheckState = sender.Tag = CheckState.Indeterminate
        self.set_ok_enabled()

    def checkbox_clicked(self, sender, event=None):
        # Helper method that sets the check state of the select all checkbox when another checkbox is clicked

        select_all = list(sender.Parent.Controls)[0]
        cbs = list(sender.Parent.Controls)[1:]
        cbs_cked = [cb.Checked for cb in cbs]
        if all(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Checked
        elif any(cbs_cked):
            select_all.CheckState = select_all.Tag = CheckState.Indeterminate
        else:
            select_all.CheckState = select_all.Tag = CheckState.Unchecked
        select_all.Enabled = any(cb.Enabled for cb in cbs)
        self.set_ok_enabled()

    def set_ok_enabled(self):
        # Helper method that enables or disables "ok Report" button
        # Enable button if a PTV is selected, at least one stat is selected, and at least one plan or eval dose is selected

        plan_cked = any(cb.Checked for cb in self.plans_gb.Controls)
        exam_cked = any(cb.Checked for cb in self.exams_gb.Controls)
        self.ok_btn.Enabled = plan_cked or exam_cked

    def ok_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


def get_tx_technique(bs):
    # Helper function that returns the treatment technique for the given beam set.
    # SMLC, VMAT, DMLC, 3D-CRT, conformal arc, or applicator and cutout
    # Return "?" if treatment technique cannot be determined
    # Code modified from RS support

    if bs.Modality == "Photons":
        if bs.PlanGenerationTechnique == "Imrt":
            if bs.DeliveryTechnique == "SMLC":
                return "SMLC"
            if bs.DeliveryTechnique == "DynamicArc":
                return "VMAT"
            if bs.DeliveryTechnique == "DMLC":
                return "DMLC"
        elif bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                # return "SMLC" # Changed from "Conformal". Failing with forward plans.
                return "Conformal"
                # return "3D-CRT"
            if bs.DeliveryTechnique == "Arc":
                return "Conformal Arc"
    elif bs.Modality == "Electrons":
        if bs.PlanGenerationTechnique == "Conformal":
            if bs.DeliveryTechnique == "SMLC":
                return "ApplicatorAndCutout"
    return "?"


def name_item(item, l, max_len=sys.maxsize):
    # Helper function that oks a unique name for `item` in list `l` (case insensitive)
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



def get_opt_func_args(opt_func):
    # Helper function that returns a dictionary of arguments to pass into AddOptimizationFunction function
    # Used for copying objectives and contraints
    # Code is from the RS 8B Scripting Guideline

    dfp = opt_func.DoseFunctionParameters
    args = {}
    args["RoiName"] = opt_func.ForRegionOfInterest.Name
    args["IsRobust"] = opt_func.UseRobustness
    args["Weight"] = dfp.Weight
    if hasattr(dfp, "FunctionType"):
        if dfp.FunctionType == "UniformEud":
            args["FunctionType"] = "TargetEud"
        else:
            args["FunctionType"] = dfp.FunctionType
        args["DoseLevel"] = dfp.DoseLevel
        if "Eud" in dfp.FunctionType:
            args["EudParameterA"] = dfp.EudParameterA
        elif "Dvh" in dfp.FunctionType:
            args["PercentVolume"] = dfp.PercentVolume
    elif hasattr(dfp, "HighDoseLevel"):
        # Dose falloff function does not have FunctionType attribute
        args["FunctionType"] = "DoseFallOff"
        args["HighDoseLevel"] = dfp.HighDoseLevel
        args["LowDoseLevel"] = dfp.LowDoseLevel
        args["LowDoseDistance"] = dfp.LowDoseDistance
    elif hasattr(dfp, "PercentStdDeviation"):
        # Uniformity constraint does not have FunctionType attribute
        args["FunctionType"] = "UniformityConstraint"
        args["PercentStdDeviation"] = dfp.PercentStdDeviation
    return args


def set_opt_func_args(opt_func, args):
    # Helper function that sets the DoseFunctionParameters attributes of an optimization function (objective or constraint)
    # Used for copying objectives and contraints
    # Code is from the RS 8B Scripting Guideline

    dfp = opt_func.DoseFunctionParameters
    dfp.Weight = args["Weight"]
    if args["FunctionType"] == "DoseFallOff":
        dfp.HighDoseLevel = args["HighDoseLevel"]
        dfp.LowDoseLevel = args["LowDoseLevel"]
        dfp.LowDoseDistance = args["LowDoseDistance"]
    elif args["FunctionType"] == "UniformityConstraint":
        dfp.PercentStdDeviation = args["PercentStdDeviation"]
    else:
        dfp.DoseLevel = args["DoseLevel"]
        if "Eud" in dfp.FunctionType:
            dfp.EudParameterA = args["EudParameterA"]
        elif "Dvh" in dfp.FunctionType:
            dfp.PercentVolume = args["PercentVolume"]


def copy_plan_opt(old_plan_opt, new_plan_opt):
    # Helper function that copies some optimization parameters, inclusing objectives and constraints, from one plan optimization to another
    # Only copy the optimization parameters that CRMC ever uses

    ## Optimization parameters

    with CompositeAction("Copy Select Optimization Parameters"):
        new_plan_opt.AutoScaleToPrescription = old_plan_opt.AutoScaleToPrescription

        new_plan_opt.OptimizationParameters.Algorithm.MaxNumberOfIterations = old_plan_opt.OptimizationParameters.Algorithm.MaxNumberOfIterations
        new_plan_opt.OptimizationParameters.Algorithm.OptimalityTolerance = old_plan_opt.OptimizationParameters.Algorithm.OptimalityTolerance
        
        new_plan_opt.OptimizationParameters.DoseCalculation.ComputeFinalDose = old_plan_opt.OptimizationParameters.DoseCalculation.ComputeFinalDose
        new_plan_opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose = old_plan_opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose
        new_plan_opt.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase = old_plan_opt.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase

        if get_tx_technique(old_plan_opt.OptimizedBeamSets[0]) == "VMAT":
            for i, old_tss in enumerate(old_plan_opt.OptimizationParameters.TreatmentSetupSettings):
                new_tss = new_plan_opt.OptimizationParameters.TreatmentSetupSettings[i]
                
                new_tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree = old_tss.SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree
                new_tss.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree = old_tss.SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree

                for j, old_bs in enumerate(old_tss.BeamSettings):
                    old_bs = old_bs.ArcConversionPropertiesPerBeam
                    new_bs = new_tss.BeamSettings[j].ArcConversionPropertiesPerBeam
                    if old_bs.NumberOfArcs != new_bs.NumberOfArcs or old_bs.FinalArcGantrySpacing != new_bs.FinalArcGantrySpacing or old_bs.MaxArcDeliveryTime != new_bs.MaxArcDeliveryTime:
                        new_bs.EditArcBasedBeamOptimizationSettings(CreateDualArcs=old_bs.NumberOfArcs == 2, FinalGantrySpacing=old_bs.FinalArcGantrySpacing, MaxArcDeliveryTime=old_bs.MaxArcDeliveryTime) 

    ## Objectives and constraints

    args_list = []

    # Constraints
    for opt_func in old_plan_opt.Constraints:
        args = get_opt_func_args(opt_func)
        args["IsConstraint"] = False
        args_list.append(args)
        
    # Objectives
    if old_plan_opt.Objective is not None:
        for opt_func in old_plan_opt.Objective.ConstituentFunctions:
            args = get_opt_func_args(opt_func)
            args["IsConstraint"] = False
            args_list.append(args)
    
    # Create each constraint/objective in new plan opt
    with CompositeAction("Copy Optimization Functions"):
        for args in args_list:
            opt_func = new_plan_opt.AddOptimizationFunction(FunctionType=args["FunctionType"], RoiName=args["RoiName"], IsConstraint=args["IsConstraint"], IsRobust=args["IsRobust"])
            set_opt_func_args(opt_func, args)


def copy_beam_set(**kwargs):
    """Copy the given beam set to a new beam set in the given plan

    Keyword Arguments
    -----------------
    old_beam_set_id: str
        BeamSetIdentifier of the beam set to copy
        If None, copy current beam set
    new_plan_names: Union[str, List[str]]
        Plan name or list of plan name(s) to which to copy the beam set
        If None, use the current plan
    new_exam_names: Union[str, List[str]]
        Name of the exams, or list of exam name(s), to which to copy the beam set
        A new plan is created on each exam!
        If None, use the current plan
    
    Copy electron or photon (including VMAT) beam sets:
        - Treatment and setup beams
            * Unique beam numbers across all cases in current patient
            * Beam names are same as numbers
        - AutoScaleToPrescription
        - For VMAT beam sets, other select optimization settings:
            * Maximum number of iterations
            * Optimality tolerance
            * Calculate intermediate and final doses
            * Iterations in preparation phase
            * Max leaf travel distance per degree (and enabled/disabled)
            * Dual arcs
            * Max gantry spacing
            * Max delivery time
            * Objectives and constraints
        - Prescriptions

    Unfortunately, dose cannot accurately be copied for VMAT, so we do the next-best things and provide the dose on additional set if the beam set is copied to a plan with a different planning exam.

    Do not optimize or compute dose
    """

    global case, exam

    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)

    gui = kwargs.get("gui", False)
    print(gui)

    if not gui:
        old_beam_set_id = kwargs.get("old_beam_set_id")
        new_plan_names = kwargs.get("new_plan_names")
        new_exam_names = kwargs.get("new_exam_names")

    if gui or old_beam_set_id is None:
        try:
            plan = get_current("Plan")
        except:
            MessageBox.Show("There is no plan loaded. Click OK to abort the script.", "No Plan Loaded")
            sys.exit(1)
        try:
            old_beam_set = get_current("BeamSet")
        except:
            MessageBox.Show("There is no beam set loaded. Click OK to abort the script.", "No Beam Set Loaded")
            sys.exit(1)
        case = get_current("Case")
        old_beam_set_id = old_beam_set.BeamSetIdentifier()
        old_plan_name, old_beam_set_name = old_beam_set_id.split(":")
    elif not gui:
        try:
            case = get_current("Case")
        except:
            raise IOError("There is no case loaded.")
        old_plan_name, old_beam_set_name = old_beam_set_id.split(":")
        try:
            plan = case.TreatmentPlans[old_plan_name]
            try:
                old_beam_set = plan.BeamSets[old_beam_set_name]
            except:
                raise ValueError("Invalid argument for `old_beam_set_id`: there is no beam set '{}' in plan '{}'.".format(old_beam_set_name, old_plan_name))
        except:
            raise ValueError("Invalid argument for `old_beam_set_id`: there is no plan '{}' in the current case.".format(old_plan_name))
    exam = old_beam_set.GetPlanningExamination()
    
    if gui:
        form = ChooseNewPlansAndExamsForm(old_plan_name)
        if form.DialogResult != DialogResult.OK:
            sys.exit()
        new_plan_names = [cb.Text for cb in form.plans_gb.Controls if cb.Checked and cb.Text != "Select all"]
        new_plan_names = [plan_name[1:] if plan_name.startswith("*") else plan_name for plan_name in new_plan_names]  # Remove asterisk from approved plan names
        new_exam_names = [cb.Text for cb in form.exams_gb.Controls if cb.Checked and cb.Text != "Select all"]
    
    if new_plan_names is None:
        new_plans = [plan]  # Copy within same plan
    else:
        if isinstance(new_plan_names, str):
            new_plan_names = [new_plan_names]
        existing_plan_names = [p.Name for p in case.TreatmentPlans]
        print(new_plan_names)
        new_plans = [case.TreatmentPlans[plan_name] for plan_name in new_plan_names if plan_name in existing_plan_names]
    if new_exam_names is None:
        new_exams = []
    else:
        if isinstance(new_exam_names, str):
            new_exam_names = [new_exam_names]
        existing_exam_names = [e.Name for e in case.Examinations]
        new_exams = [case.Examinations[exam_name] for exam_name in new_exam_names if exam_name in existing_exam_names]

    warnings = ""

    imported = old_beam_set.HasImportedDose()

    # Ensure machine name is recognized
    machine_name = old_beam_set.MachineReference.MachineName
    if imported:
        if machine_name.endswith("_imported"):
            machine_name = machine_name[:-9]
        else:
            fx = old_beam_set.FractionationPattern
            if fx is None or fx.NumberOfFractions > 5:
                machine_name = "ELEKTA"
            elif old_beam_set.Presciption is not None and old_beam_set.Prescription.PrimaryDosePrescription.DoseValue >= 600 * fx.NumberOfFractions:
                machine_name = "SBRT 6MV"
            else:
                machine_name = "ELEKTA"
        warnings += "Machine '{}' is not commissioned, so new beam set will use machine '{}'.".format(old_beam_set.MachineReference.MachineName, machine_name)

    tx_technique = get_tx_technique(old_beam_set)

    # Create new plan on each exam, and add new plans to new plans list
    ext = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"][0]
    old_plan_by = case.TreatmentPlans[old_plan_name].PlannedBy
    for e in new_exams:
        new_plan_name = name_item(old_beam_set_name, [p.Name for p in case.TreatmentPlans], 16)
        new_plan_comments = "New plan for '{}' on '{}'".format(old_beam_set_id, e.Name)
        new_plan = case.AddNewPlan(PlanName=new_plan_name, PlannedBy=old_plan_by, Comment=new_plan_comments, ExaminationName=e.Name, AllowDuplicateNames=False)
        dg = plan.GetDoseGrid()
        ext.CreateExternalGeometry(Examination=new_plan.GetStructureSet().OnExamination)
        new_plan.SetDefaultDoseGrid(VoxelSize=dg.VoxelSize)
        new_plans.append(new_plan)

    # Copy beam set to each new plan
    print("ready to copy")
    print(new_plans)
    for new_plan in new_plans:
        print(new_plan)
        if hasattr(new_plan, "Review") and new_plan.Review is not None and new_plan.Review.ApprovalStatus == "Approved":
            print("approved")
            new_plan = case.TreatmentPlans[copy_plan_without_changes(new_plan.Name)]
            print(new_plan)

        new_beam_set_name = name_item(old_beam_set.DicomPlanLabel, [beam_set.DicomPlanLabel for beam_set in new_plan.BeamSets])  # Unique beam set name in plan
        new_beam_set = new_plan.AddNewBeamSet(Name=new_beam_set_name, ExaminationName=new_plan.GetStructureSet().OnExamination.Name, MachineName=machine_name, Modality=old_beam_set.Modality, TreatmentTechnique=tx_technique, PatientPosition=old_beam_set.PatientPosition, NumberOfFractions=old_beam_set.FractionationPattern.NumberOfFractions, CreateSetupBeams=old_beam_set.PatientSetup.UseSetupBeams, Comment="Copy of {}".format(old_beam_set.DicomPlanLabel))

        # Unique beam number
        beam_num = 1
        for c in patient.Cases:
            for p in c.TreatmentPlans:
                for bs in p.BeamSets:
                    for b in bs.Beams:
                        beam_num = max(beam_num, b.Number + 1)
                    for sb in bs.PatientSetup.SetupBeams:
                        beam_num = max(beam_num, sb.Number + 1)

        ## Copy beam set
        if not imported and plan.GetStructureSet().OnExamination.Name == new_plan.GetStructureSet().OnExamination.Name:
            new_beam_set.CopyBeamsFromBeamSet(BeamSetToCopyFrom=old_beam_set, BeamsToCopy=[b.Name for b in old_beam_set.Beams])
        # CopyBeamsFromBeamSet doesn't work for uncommissioned machines, so manually add beams and copy segments
        # Code modified from RS support's CopyBeamSet script
        else:
            for i, beam in enumerate(old_beam_set.Beams):
                iso_data = new_beam_set.CreateDefaultIsocenterData(Position=beam.Isocenter.Position)
                iso_data["Name"] = iso_data["NameOfIsocenterToRef"] = beam.Isocenter.Annotation.Name
                energy = beam.MachineReference.Energy
                if old_beam_set.Modality == "Electrons":
                    new_beam = new_beam_set.CreateElectronBeam(Energy=energy, Name=beam.Name, GantryAngle=beam.GantryAngle, CouchAngle=beam.CouchAngle, ApplicatorName=beam.Applicator.ElectronApplicatorName, InsertName=beam.Applicator.Insert.Name, IsAddCutoutChecked=True, IsocenterData=iso_data)
                    new_beam.Applicator.Insert.Contour = beam.Applicator.Insert.Contour
                    new_beam.BeamMU = beam.BeamMU
                    new_beam.Description = beam.Description
                #elif tx_technique == "VMAT":
                    #new_beam = new_beam_set.CreateArcBeam(Energy=energy, Name=beam.Name, ArcRotationDirection=beam.ArcRotationDirection, ArcStopGantryAngle=beam.ArcStopGantryAngle, GantryAngle=beam.GantryAngle, CouchAngle=beam.CouchAngle, CollimatorAngle=beam.InitialCollimatorAngle, IsocenterData=iso_data)
                    
                    """
                    Does not work without license rayWave :(
                    gantry_angles, couch_angles, coll_angles = [], [], []
                    for s in beam.Segments:
                        gantry_angle = beam.GantryAngle + s.DeltaGantryAngle
                        if gantry_angle < 0:
                            gantry_angle += 360
                        gantry_angles.append(gantry_angle)

                        couch_angle = beam.CouchAngle + s.DeltaCouchAngle
                        if couch_angle < 0:
                            couch_angle += 360
                        couch_angles.append(couch_angle)

                        coll_angles.append(s.CollimatorAngle)

                    new_beam.SetArcTrajectory(GantryAngles=gantry_angles, CouchAngles=couch_angles, CollimatorAngles=coll_angles, ArcRotationDirection=new_beam.ArcRotationDirection)
                    """
                    
                elif tx_technique != "VMAT":
                    for s in beam.Segments:
                        name = name_item(beam.Name, [b.Name for b in new_beam_set.Beams], 16)
                        new_beam = new_beam_set.CreatePhotonBeam(Energy=energy, Name=name, GantryAngle=beam.GantryAngle, CouchAngle=beam.CouchAngle, CollimatorAngle=s.CollimatorAngle, IsocenterData=iso_data)  
                        new_beam.BeamMU = round(beam.BeamMU * s.RelativeWeight, 2)
                        new_beam.CreateRectangularField()
                        new_beam.Segments[0].JawPositions = s.JawPositions
                        new_beam.Segments[0].LeafPositions = s.LeafPositions
                    if beam.Segments.Count > 1:
                        new_beam_set.MergeBeamSegments(TargetBeamName=new_beam_set.Beams[i].Name, MergeBeamNames=[b.Name for b in new_beam_set.Beams][(i + 1):])
                    new_beam_set.Beams[i].Description = beam.Description

                else:
                    new_beam = new_beam_set.CreateArcBeam(ArcStopGantryAngle=beam.ArcStopGantryAngle, ArcRotationDirection=beam.ArcRotationDirection, Energy=energy, Name=beam.Name, GantryAngle=beam.GantryAngle, CouchAngle=beam.CouchAngle, CollimatorAngle=beam.InitialCollimatorAngle, IsocenterData=iso_data)
                    new_beam.BeamMU = beam.BeamMU
                    new_beam.Description = beam.Description

                    old_beam_set.ComputeDoseOnAdditionalSets(ExaminationNames=[new_plan.GetStructureSet().OnExamination.Name], FractionNumbers=[0])

        # Rename and renumber new beams
        for b in new_beam_set.Beams:
            b.Number = beam_num
            b.Name = str(beam_num)
            beam_num += 1

        # Manually copy setup beams from old beam set
        if old_beam_set.PatientSetup.UseSetupBeams and old_beam_set.PatientSetup.SetupBeams.Count > 0:
            old_sbs = old_beam_set.PatientSetup.SetupBeams
            #new_beam_set.RemoveSetupBeams()
            new_beam_set.UpdateSetupBeams(ResetSetupBeams=True, SetupBeamsGantryAngles=[sb.GantryAngle for sb in old_sbs])  # Clear the setup beams created when the beam set was added, to ensure no extraneous setup beams in new beam set
            for i, old_sb in enumerate(old_sbs):
                new_sb = new_beam_set.PatientSetup.SetupBeams[i]
                new_sb.Number = beam_num
                new_sb.Name = str(beam_num)
                new_sb.Description = old_sb.Description
                beam_num += 1

        old_plan_opt = [opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].DicomPlanLabel == old_beam_set.DicomPlanLabel][0]  # Get PlanOptimizations with a single optimized beam set - the old beam set (assume only one)
        new_plan_opt = [opt for opt in new_plan.PlanOptimizations if opt.OptimizedBeamSets.Count == 1 and opt.OptimizedBeamSets[0].DicomPlanLabel == new_beam_set_name][0]  # Get PlanOptimizations with a single optimized beam set - the new beam set (assume only one)
        
        # Copy optimization parameters, if applicable
        if tx_technique == "VMAT":
            # Copy optimization parameters (the only ones that CRMC ever uses)
            copy_plan_opt(old_plan_opt, new_plan_opt)

        # Copy Rx's
        if old_beam_set.Prescription is not None:
            with CompositeAction("Copy Prescriptions"):
                autoscale = old_plan_opt.AutoScaleToPrescription  # Assume all Rx's for the beam set have the same autoscale setting
                for old_rx in old_beam_set.Prescription.DosePrescriptions:  # Copy all Rx's, not just the primary Rx
                    if hasattr(old_rx, "OnStructure"):
                        if old_rx.PrescriptionType == "DoseAtPoint":  # Rx to POI
                            new_beam_set.AddDosePrescriptionToPoi(PoiName=old_rx.OnStructure.Name, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel, AutoScaleDose=autoscale)
                        else:  # Rx to ROI
                            new_beam_set.AddDosePrescriptionToRoi(RoiName=old_rx.OnStructure.Name, DoseVolume=old_rx.DoseVolume, PrescriptionType=old_rx.PrescriptionType, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel, AutoScaleDose=autoscale)
                    elif old_rx.OnDoseSpecificationPoint is None:  # Rx to DSP
                        new_beam_set.AddDosePrescriptionToSite(Description=old_rx.Description, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel, AutoScaleDose=autoscale)
                    else:  # Rx to site that is not a DSP
                        new_beam_set.AddDosePrescriptionToSite(Description=old_rx.Description, NameOfDoseSpecificationPoint=old_rx.OnDoseSpecificationPoint.Name, DoseValue=old_rx.DoseValue, RelativePrescriptionLevel=old_rx.RelativePrescriptionLevel, AutoScaleDose=autoscale)

        # Compute dose
        if new_beam_set.Modality != "Electrons" and new_beam_set.Beams.Count > 0:
            if any(beam.Segments.Count == 0 for beam in new_beam_set.Beams):
                MessageBox.Show("Cannot compute dose because beam(s) do not have valid segments.", "No Segments")
                sys.exit()
            else:
                new_beam_set.ComputeDose(ComputeBeamDoses=True, DoseAlgorithm=new_beam_set.AccurateDoseAlgorithm.DoseAlgorithm, ForceRecompute=True)
