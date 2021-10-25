# For GUI (MessageBox displays errors)
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import sys

from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


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
    new_plan_name: str
        Name of the plan to which to copy the beam set
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

    old_beam_set_id = kwargs.get("old_beam_set_id")
    new_plan_name = kwargs.get("new_plan_name")

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)
    if old_beam_set_id is None:
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
    else:
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
    if new_plan_name is None:
        new_plan = plan  # Copy within same plan
    else:
        try:
            new_plan = case.TreatmentPlans[new_plan_name]
        except:
            raise ValueError("Invalid argument for `new_plan_name`: there is no plan '{}' in the current case.".format(new_plan_name))

    warnings = ""

    # Alert user and exit script with an error if the plan is approved
    if new_plan.Review is not None and new_plan.Review.ApprovalStatus == "Approved":
        MessageBox.Show("Plan is approved, so a beam set cannot be added. Click OK to abort the script.", "Plan Is Approved")
        sys.exit(1)

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

    ## Create new beam set on new planning exam

    new_beam_set_name = name_item(old_beam_set.DicomPlanLabel, [beam_set.DicomPlanLabel for beam_set in new_plan.BeamSets])  # Unique beam set name in plan
    tx_technique = get_tx_technique(old_beam_set)
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
                
            #new_beam.BeamMU = beam.BeamMU
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

