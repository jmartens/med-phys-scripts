"""
This attached script is provided as a tool and not as a RaySearch endorsed script for
clinical use.  The use of this script in its whole or in part is done so
without any guarantees of accuracy or expected outcomes. Please verify all results. Further,
per the Raystation Instructions for Use, ALL scripts MUST be verified by the user prior to
clinical use.

Copy Plan to New CT or Merge Beamsets 2.2.1

It can copy a photon or electron plan from one CT to another.  It will not copy wedges.
It does support VMAT and RayStation Versions 4.7 and 5.  You must ensure that an External
exists on the new CT.  - No, you don't: script creates it, if necessary.  - Kaley

It can also merge photon or electron beamsets from one plan to another. It is assumed that
the plans are on the same examination
"""

import sys
import clr

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from random import choice
from re import IGNORECASE, search
from string import ascii_uppercase
from System.Drawing import *
from System.Windows.Forms import *

from connect import *


case = None


class CopyPlanToNewCTOrMergeBeamSetsForm(Form):

    def __init__(self):
        self.Text = "Copy Plan To New CT Or Merge Beam Sets"
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.StartPosition = FormStartPosition.CenterScreen

        tab_ctrl = TabControl()
        tab_ctrl.Size = Size(292, 336)
        self.Controls.Add(tab_ctrl)

        plan_names = sorted(plan.Name for plan in case.TreatmentPlans)

        ## "Copy Plan to New CT" tab

        y = 15
        tab_pg = TabPage()
        tab_pg.Text = "Copy Plan to New CT"
        tab_ctrl.Controls.Add(tab_pg)

        if case.Examinations.Count == 1:
            lbl = Label()
            lbl.Location = Point(15, y)
            lbl.Text = "There are no other examinations to copy to."
            lbl.Size = Size(300, lbl.Height)
            tab_pg.Controls.Add(lbl)
        else:
            # "Plan to Copy" input

            lbl = Label()
            lbl.Text = "Plan to Copy"
            lbl.Size = Size(91, lbl.Height)
            lbl.Location = Point(15, y)
            tab_pg.Controls.Add(lbl)
            y += lbl.Height

            self.p2cCB = ComboBox()
            self.p2cCB.DropDownStyle = ComboBoxStyle.DropDownList
            self.p2cCB.Size = Size(187, self.p2cCB.Height)
            self.p2cCB.Location = Point(15, y)
            self.p2cCB.Items.AddRange(plan_names)
            self.p2cCB.SelectedIndexChanged += self.p2cCB_SelectedIndexChanged
            tab_pg.Controls.Add(self.p2cCB)
            y += self.p2cCB.Height + 15

            # "New Plan Name" input

            self.PlanNameLbl = Label()
            self.PlanNameLbl.Text = "New Plan Name"
            self.PlanNameLbl.Size = Size(110, self.PlanNameLbl.Height)
            self.PlanNameLbl.Location = Point(15, y)
            tab_pg.Controls.Add(self.PlanNameLbl)
            y += self.PlanNameLbl.Height

            self.NewPlanNameTextBox = TextBox()
            self.NewPlanNameTextBox.Size = Size(187, 23)
            self.NewPlanNameTextBox.TextChanged += self.shouldEnableSubmit1
            self.NewPlanNameTextBox.Location = Point(15, y)
            tab_pg.Controls.Add(self.NewPlanNameTextBox)
            y += self.NewPlanNameTextBox.Height + 15

            # "New Dataset" input

            self.DatasetLbl = Label()
            self.DatasetLbl.Text = "New Dataset"
            self.DatasetLbl.Size = Size(110, self.DatasetLbl.Height)
            self.DatasetLbl.Location = Point(15, y)
            tab_pg.Controls.Add(self.DatasetLbl)
            y += self.DatasetLbl.Height

            self.NewDatasetCB = ComboBox()
            self.NewDatasetCB.DropDownStyle = ComboBoxStyle.DropDownList
            self.NewDatasetCB.Size = Size(187, self.NewDatasetCB.Height)
            self.NewDatasetCB.SelectedIndexChanged += self.shouldEnableSubmit1
            self.NewDatasetCB.Location = Point(15, y)
            self.NewDatasetCB.Items.AddRange(sorted(exam.Name for exam in case.Examinations))
            tab_pg.Controls.Add(self.NewDatasetCB)
            y += self.NewDatasetCB.Height + 15

            # "Submit" and "Cancel" buttons

            self.SubmitButton = Button()
            self.SubmitButton.Text = "Submit"
            self.SubmitButton.Size = Size(75, self.SubmitButton.Height)
            self.SubmitButton.Enabled = False
            self.SubmitButton.Click += self.SubmitButton_Click
            self.SubmitButton.Location = Point(15, y)
            tab_pg.Controls.Add(self.SubmitButton)

            self.CancelButton = Button()
            self.CancelButton.Text = "Cancel"
            self.CancelButton.Size = Size(75, self.CancelButton.Height)
            self.CancelButton.Click += self.CancelButton_Click
            self.CancelButton.Location = Point(20 + self.SubmitButton.Width, y)
            tab_pg.Controls.Add(self.CancelButton)

        ## "Copy Beam Set" tab

        y = 15
        tab_pg = TabPage()
        tab_pg.Text = "Copy Beam Set"
        tab_ctrl.Controls.Add(tab_pg)

        # "Plan to Copy From" input

        lbl = Label()
        lbl.Text = "Plan to Copy From"
        lbl.Size = Size(187, lbl.Height)
        lbl.Location = Point(15, y)
        tab_pg.Controls.Add(lbl)
        y += lbl.Height

        self.p2cfCB = ComboBox()
        self.p2cfCB.DropDownStyle = ComboBoxStyle.DropDownList
        self.p2cfCB.Size = Size(187, self.p2cfCB.Height)
        self.p2cfCB.SelectedIndexChanged += self.p2cfCB_SelectedIndexChanged
        self.p2cfCB.Location = Point(15, y)
        self.p2cfCB.Items.AddRange(plan_names)
        tab_pg.Controls.Add(self.p2cfCB)
        y += self.p2cfCB.Height + 15

        # "BeamSet(s) to Copy" input

        self.DatasetLbl2 = Label()
        self.DatasetLbl2.Text = "BeamSet(s) to Copy"
        self.DatasetLbl2.Size = Size(110, self.DatasetLbl2.Height)
        self.DatasetLbl2.Location = Point(15, y)
        tab_pg.Controls.Add(self.DatasetLbl2)
        y += self.DatasetLbl2.Height

        self.bsLB = ListBox()
        self.bsLB.Size = Size(187, 90)
        self.bsLB.SelectedIndexChanged += self.shouldEnableSubmit2
        self.bsLB.SelectionMode = SelectionMode.MultiExtended
        self.bsLB.Location = Point(15, y)
        tab_pg.Controls.Add(self.bsLB) 
        y += self.bsLB.Height + 15

        # "Plan to Copy To" input

        self.BeamSetLbl = Label()
        self.BeamSetLbl.Text = "Plan to Copy To"
        self.BeamSetLbl.Size = Size(110, self.BeamSetLbl.Height)
        self.BeamSetLbl.Location = Point(15, y)
        tab_pg.Controls.Add(self.BeamSetLbl)
        y += self.BeamSetLbl.Height

        self.p2c2CB = ComboBox()
        self.p2c2CB.DropDownStyle = ComboBoxStyle.DropDownList
        self.p2c2CB.Size = Size(187, self.p2c2CB.Height)
        self.p2c2CB.SelectedIndexChanged += self.shouldEnableSubmit2
        self.p2c2CB.Location = Point(15, y)
        tab_pg.Controls.Add(self.p2c2CB) 
        y += self.p2c2CB.Height + 15

        # "Submit" and "Cancel" buttons

        self.SubmitButton2 = Button()
        self.SubmitButton2.Text = "Submit"
        self.SubmitButton2.Size = Size(75, self.SubmitButton2.Height)
        self.SubmitButton2.Enabled = False
        self.SubmitButton2.Click += self.SubmitButton2_Click
        self.SubmitButton2.Location = Point(15, y)
        tab_pg.Controls.Add(self.SubmitButton2)

        self.CancelButton2 = Button()
        self.CancelButton2.Text = "Cancel"
        self.CancelButton2.Size = Size(75, self.CancelButton2.Height)
        self.CancelButton2.Click += self.CancelButton_Click
        self.CancelButton2.Location = Point(20 + self.SubmitButton2.Width, y)
        tab_pg.Controls.Add(self.CancelButton2)          

        self.mode = ""

        tab_ctrl.Size = self.ClientSize
        self.ShowDialog()

    def SubmitButton_Click(self, sender, e):
        self.mode = "copy"
        self.DialogResult = DialogResult.OK

    def CancelButton_Click(self, sender, e):
        self.DialogResult = DialogResult.Cancel

    def p2cCB_SelectedIndexChanged(self, sender, event):
        if self.p2cCB.Text != "":
            exam_names = [exam.Name for exam in case.Examinations if exam.Name != case.TreatmentPlans[self.p2cCB.Text].GetStructureSet().OnExamination.Name]
            self.NewDatasetCB.Items.Clear()
            for exam_name in sorted(exam_names):
                self.NewDatasetCB.Items.Add(exam_name)
        self.shouldEnableSubmit1()

    def p2cfCB_SelectedIndexChanged(self, sender, e):
        if self.p2cfCB.Text != "":
            bsns = [b.DicomPlanLabel for b in case.TreatmentPlans[self.p2cfCB.Text].BeamSets]
            self.bsLB.Items.Clear()
            for bsn in sorted(bsns):
                self.bsLB.Items.Add(bsn)

            self.p2c2CB.Items.Clear()
            self.p2c2CB.Items.AddRange(sorted(plan.Name for plan in case.TreatmentPlans if plan.Name != self.p2cfCB.Text))
        self.shouldEnableSubmit2()

    def SubmitButton2_Click(self, sender, e):
        self.mode = "merge"
        self.DialogResult = DialogResult.OK

    def shouldEnableSubmit1(self, sender=None, event=None):  # Sometimes not called as event handler
        if (self.NewPlanNameTextBox.Text != ""
                and self.NewDatasetCB.SelectedItem is not None
                and self.p2cCB.SelectedItem is not None):
            self.SubmitButton.Enabled = True

        else:
            self.SubmitButton.Enabled = False

    def shouldEnableSubmit2(self, sender=None, event=None):  # Sometimes not called as event handler
        if (self.p2c2CB.SelectedItem is not None
                and self.bsLB.SelectedItems.Count > 0
                and self.p2cfCB.SelectedItem is not None):
            self.SubmitButton2.Enabled = True

        else:
            self.SubmitButton2.Enabled = False


def TreatmentTechnique(bs):
    TT = "?"
    if bs.Modality == "Photons":
        if bs.PlanGenerationTechnique == "Imrt":

            if bs.DeliveryTechnique == "SMLC":
                TT = "SMLC"
            if bs.DeliveryTechnique == "DynamicArc":
                TT = "VMAT"
            if bs.DeliveryTechnique == "DMLC":
                TT = "DMLC"

        if bs.PlanGenerationTechnique == "Conformal":

            if bs.DeliveryTechnique == "SMLC":
                TT = "SMLC" # Changed from "Conformal". Failing with forward plans.
            if bs.DeliveryTechnique == "Arc":
                TT = "Conformal Arc"

    if bs.Modality == "Electrons":
        if bs.PlanGenerationTechnique == "Conformal":

            if bs.DeliveryTechnique == "SMLC":
                TT = "ApplicatorAndCutout"

    return TT


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


# -----------------------
# OBSERVE: The script would probably generate the wrong isocenter if patient is not HFS.
# OBSERVE: The script doesn't work for VMAT beams. No beams are copied.  - Actually, it does. It just doesn't copy MLC positions. - Kaley
# -----------------------
def copy_plan_to_ct_or_merge_beam_sets(**kwargs):
    """It can copy a photon or electron plan from one CT to another.  It will not copy wedges.
    It does support VMAT and RayStation Versions 4.7 and 5.  You must ensure that an External
    exists on the new CT.  - No, you don't: script creates it, if necessary.  - Kaley

    It can also merge photon or electron beamsets from one plan to another. It is assumed that
    the plans are on the same examination

    Optionally, user chooses settings from a GUI.

    Keyword Arguments
    -----------------
    gui: bool
        True if user should choose settings from a GUI, False otherwise
        If True, all other keyword arguments are ignored
        Defaults to False

    mode: str
        "copy" if we are copying a plan to a new CT, "merge" if we are merging a beam set from one plan to another
        Defaults to "copy"

    originalPlanName: ScriptObject
        Name of the plan to copy
        Used with either mode
        Plan must exist in current case
        Defaults to current plan
        
    replanName: str
        Used with either mode
        For mode "copy":
            Name of the new plan to create
            Defaults to `originalPlanName`, made unique
        For mode "merge":
            Name of the plan to merge the beam set(s) to
            Must exist in current case
            Required

    replanCTName: str
        Name of the exam to copy the plan to
        Used with mode "copy"
        Required

    beamSetNames: List[str]
        Name(s) of the beam set(s) to merge
        Used with mode "merge"
        Defaults to all beam set names in original plan
    """

    global case

    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)

    if case.TreatmentPlans.Count == 0:
        MessageBox.Show("There are no plans in the current case. Click OK to abort the script.", "No Plans")
        sys.exit(1)
    
    gui = kwargs.get("gui", False)
    if gui:
        dialog = CopyPlanToNewCTOrMergeBeamSetsForm()
        if dialog.DialogResult == DialogResult.Cancel:
            sys.exit()
        mode = dialog.mode
        if mode == "copy":
            originalPlan = case.TreatmentPlans[dialog.p2cCB.Text]
            beamSetNames = [bs.DicomPlanLabel for bs in originalPlan.BeamSets]
            replanCTName = dialog.NewDatasetCB.SelectedItem
            replanName = dialog.NewPlanNameTextBox.Text
        else:
            originalPlan = case.TreatmentPlans[dialog.p2cfCB.Text]
            beamSetNames = dialog.bsLB.SelectedItems
            replan = case.TreatmentPlans[dialog.p2c2CB.Text]
            replanCTName = replan.GetStructureSet().OnExamination.Name
    else:
        mode = kwargs.get("mode", "copy")
        originalPlanName = kwargs.get("originalPlanName")
        if originalPlanName is None:
            try:
                originalPlan = get_current("Plan")
            except:
                raise ValueError("There is no plan loaded, so keyword argument `originalPlanName` must be provided.")
        else:
            try:
                originalPlan = case.TreatmentPlans[originalPlanName]
            except:
                raise ValueError("There is no plan '{}' in the current case.".format(originalPlanName))
        if mode == "copy":
            replanName = kwargs.get("replanName", originalPlanName)
            replanName = name_item(replanName, [p.Name for p in case.TreatmentPlans])
            try:
                replanCTName = kwargs["replanCTName"]
                if replanCTName not in [e.Name for e in case.Examinations]:
                    raise ValueError("Examination '{}' does not exist in current case.".format(replanCTName))
            except:
                raise ValueError("`mode` is 'copy', so argument `replanCTName` must be provided")
            beamSetNames = [bs.DicomPlanLabel for bs in originalPlan.BeamSets]
        elif mode == "merge":
            try:
                replanName = kwargs["replanName"]
            except:
                raise ValueError("`mode` is 'merge', so argument `replanName` must be provided")
            replan = case.TreatmentPlans[replanName]
            replanCTName = replan.GetStructureSet().OnExamination.Name
            valid_bs_names = [bs.DicomPlanLabel for bs in originalPlan.BeamSets]
            try:
                beamSetNames = kwargs["beamSetNames"]
                invalid_bs_names = [bs_name for bs_name in beamSetNames if bs_name not in valid_bs_names]
                if invalid_bs_names:
                    raise ValueError("The following beam sets do not exist in plan '{}':\n{}".format(replanName, "\n\t-  ".join(invalid_bs_names)))
            except:
                beamSetNames = valid_bs_names
        else:
            raise ValueError("Invalid keyword argument `mode`. Must be 'copy' or 'merge'.")

    if originalPlan.TreatmentCourse.TotalDose.DoseValues is None:
        MessageBox.Show("Plan has no dose. Click OK to abort the script.", "No Dose")
        sys.exit(1)
    
    poIndex = 0
    opoIndex = 0

    reg = None

    examination = case.Examinations[replanCTName]
    ext = [roi for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"][0]
    if mode == "copy":
        originalPlanCTName = \
            originalPlan.GetStructureSet().OnExamination.Name

        # Does a registration exist? We will use this to maintain copy vs merge state
        reg = case.GetTransformForExaminations(
            FromExamination=originalPlanCTName,
            ToExamination=replanCTName
        )
        if reg is None:
            case.ComputeRigidImageRegistration(FloatingExaminationName=replanCTName, ReferenceExaminationName=originalPlanCTName, HighWeightOnBones=True)

            # Navigate to registration & allow user to make changes
            ui = get_current("ui")
            ui.TitleBar.MenuItem["Patient Modeling"].Button_Patient_Modeling.Click()
            ui = get_current("ui")  # New UI so that "Image Registration" tab is available
            ui.TabControl_Modules.TabItem["Image Registration"].Select()
            ui.ToolPanel.TabItem["Registrations"].Select()
            ui = get_current("ui")  # New UI so that list of registrations is available
            ui.ToolPanel.RegistrationList.TreeItem[replanCTName].Select()
            ui.ToolPanel.TabItem["Scripting"].Select()
            await_user_input("Review the rigid registration and make any necessary changes.")
            reg = case.GetTransformForExaminations(
                FromExamination=originalPlanCTName,
                ToExamination=replanCTName
            )

        # if a plan with the same name exists, rename the new plan to include 5 random characters
        for plancheck in case.TreatmentPlans:
            if replanName == plancheck.Name:
                replanName += "".join(choice(ascii_uppercase) for i in range(5))
                break

        # lets create the copy plan
        replan = case.AddNewPlan(PlanName=replanName, PlannedBy="", Comment="",
                                    ExaminationName=replanCTName,
                                    AllowDuplicateNames=False)
        dg = originalPlan.GetDoseGrid()
        if not replan.GetStructureSet().RoiGeometries[ext.Name].HasContours():
            ext.CreateExternalGeometry(Examination=case.Examinations[replanCTName])
        replan.SetDefaultDoseGrid(VoxelSize=dg.VoxelSize)
    # if not copy then must be merge
    else:
        inrange = True
        i = 0
        while inrange:
            try:
                replan.PlanOptimizations[i]
                i += 1
                poIndex = i
            except:
                inrange = False
        for (i,_po) in enumerate(originalPlan.PlanOptimizations):
            try:
                if _po.OptimizedBeamSets[0].DicomPlanLabel == originalPlan.BeamSets[beamSetNames[0]].DicomPlanLabel:
                    opoIndex = i
            except:
                pass

    for bs_name in beamSetNames:
        bs = originalPlan.BeamSets[bs_name]
        rx = bs.FractionDose.ForBeamSet.Prescription.PrimaryDosePrescription
        name = bs.DicomPlanLabel
        machineName = bs.MachineReference.MachineName
        modality = bs.Modality
        patientPosition = bs.PatientPosition
        fractions = bs.FractionationPattern.NumberOfFractions
        TT = TreatmentTechnique(bs)

        # if the beamset name exists we will add 5 random characters to maintain uniqueness
        for bscheck in replan.BeamSets:
            if bs.DicomPlanLabel == bscheck.DicomPlanLabel:
                name += "".join(choice(ascii_uppercase) for i in range(5))
                break

        # lets create the beamset
        rpbs = replan.AddNewBeamSet(
            Name=name,
            ExaminationName=replanCTName,
            MachineName=machineName,
            Modality=modality,
            TreatmentTechnique=TT,
            PatientPosition=patientPosition,
            NumberOfFractions=fractions,
            CreateSetupBeams=True,
            Comment="",
            )

        if modality == "Electrons":
            for beam in bs.Beams:
                CEBargs = {
                    "Energy" : beam.MachineReference.Energy,
                    "Name" : beam.Name,
                    "GantryAngle" : beam.GantryAngle,
                    "CouchAngle" : beam.CouchAngle,
                    "ApplicatorName" : beam.Applicator.ElectronApplicatorName,
                    "InsertName" : beam.Applicator.Insert.Name,
                    "IsAddCutoutChecked" : True
                }
                
                iso = beam.Isocenter.Position
                isocenter = {
                    "x":iso.x,
                    "y":iso.y,
                    "z":iso.z
                }

                # check if copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=originalPlanCTName,
                        ToExamination=replanCTName,
                        Point=isocenter
                    )
                    isocenter = {
                        "x":iso.x,
                        "y":iso.y,
                        "z":iso.z
                    }
                CEBargs["IsocenterData"] = bs.CreateDefaultIsocenterData(Position=isocenter)
                iso_name = name_item(CEBargs["IsocenterData"]["Name"], [b.Isocenter.Annotation.Name for b_set in replan.BeamSets for b in b_set.Beams if b_set.DicomPlanLabel != bs.DicomPlanLabel], 16)
                CEBargs["IsocenterData"]["NameOfIsocenterToRef"] = CEBargs["IsocenterData"]["Name"] = iso_name

                newBeam = rpbs.CreateElectronBeam(**CEBargs)
                if rpbs.Beams.Count > 1:
                    newBeam.SetIsocenter(Name=rpbs.Beams[0].Isocenter.Annotation.Name)
                contour = [{"x":c.x, "y":c.y} for c in beam.Applicator.Insert.Contour]
                newBeam.Applicator.Insert.Contour = contour
                newBeam.BeamMU = beam.BeamMU

        # not an electron plan is it also not a VMAT?
        elif TT != "VMAT":
            for beam in bs.Beams:
                CPBargs = {
                    "Energy": beam.MachineReference.Energy,
                    "Name": beam.Name,
                    "GantryAngle": beam.GantryAngle,
                    "CouchAngle": beam.CouchAngle,
                    "CollimatorAngle": beam.InitialCollimatorAngle,
                    }

                iso = beam.Isocenter.Position
                isocenter = {
                    "x":iso.x,
                    "y":iso.y,
                    "z":iso.z
                }

                # check if copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=originalPlanCTName,
                        ToExamination=replanCTName,
                        Point=isocenter
                    )
                    isocenter = {
                        "x":iso.x,
                        "y":iso.y,
                        "z":iso.z
                    }
                CPBargs["IsocenterData"] = bs.CreateDefaultIsocenterData(Position=isocenter)
                iso_name = name_item(CPBargs["IsocenterData"]["Name"], [b.Isocenter.Annotation.Name for b_set in replan.BeamSets for b in b_set.Beams if b_set.DicomPlanLabel != bs.DicomPlanLabel], 16)
                CPBargs["IsocenterData"]["NameOfIsocenterToRef"] = CPBargs["IsocenterData"]["Name"] = iso_name

                newBeam = rpbs.CreatePhotonBeam(**CPBargs)
                if rpbs.Beams.Count > 1:
                    newBeam.SetIsocenter(Name=rpbs.Beams[0].Isocenter.Annotation.Name)
                for (i, s) in enumerate(beam.Segments):
                    newBeam.CreateRectangularField()
                    newBeam.Segments[i].LeafPositions = s.LeafPositions
                    newBeam.Segments[i].JawPositions = s.JawPositions
                    newBeam.BeamMU = beam.BeamMU
                    newBeam.Segments[i].RelativeWeight = s.RelativeWeight

        # must be a VMAT
        else:
            for beam in bs.Beams:
                CABargs = {
                    "ArcStopGantryAngle": beam.ArcStopGantryAngle,
                    "ArcRotationDirection": beam.ArcRotationDirection,
                    "Energy": beam.MachineReference.Energy,
                    "Name": beam.Name,
                    "GantryAngle": beam.GantryAngle,
                    "CouchAngle": beam.CouchAngle,
                    "CollimatorAngle": beam.InitialCollimatorAngle,
                    }

                iso = beam.Isocenter.Position
                isocenter = {
                    "x":iso.x,
                    "y":iso.y,
                    "z":iso.z
                }

                # check if copy or merge
                if reg is not None:
                    iso = case.TransformPointFromExaminationToExamination(
                        FromExamination=originalPlanCTName,
                        ToExamination=replanCTName,
                        Point=isocenter
                    )
                    isocenter = {
                        "x":iso.x,
                        "y":iso.y,
                        "z":iso.z
                    }
                CABargs["IsocenterData"] = bs.CreateDefaultIsocenterData(Position=isocenter)
                iso_name = name_item(CABargs["IsocenterData"]["Name"], [b.Isocenter.Annotation.Name for b_set in replan.BeamSets for b in b_set.Beams if b_set.DicomPlanLabel != bs.DicomPlanLabel], 16)
                CABargs["IsocenterData"]["NameOfIsocenterToRef"] = CABargs["IsocenterData"]["Name"] = iso_name

                newBeam = rpbs.CreateArcBeam(**CABargs)
                if rpbs.Beams.Count > 1:
                    newBeam.SetIsocenter(Name=rpbs.Beams[0].Isocenter.Annotation.Name)

            # we cannot create controlpoints directly
            po = replan.PlanOptimizations[poIndex]
    
            # remove an roi named "dummyPTV" if it exists
            try:
                roi = case.PatientModel.RegionsOfInterest["dummyPTV"]
                roi.DeleteRoi()
            except:
                pass

            # create OUR dummyPTV roi
            with CompositeAction("Create dummy"):
                dummy = case.PatientModel.CreateRoi(Name="dummyPTV", Color="Red"
                        , Type="Ptv", TissueName=None, RoiMaterial=None)

                dummy.CreateSphereGeometry(Radius=2,
                                        Examination=examination,
                                        Center=isocenter)

            # remove any existing optimization objectives
            if po.Objective != None:
                for f in po.Objective.ConstituentFunctions:
                    f.DeleteFunction()

            # add OUR optimization objectives
            with CompositeAction("Add objective"):
                function = po.AddOptimizationFunction(
                    FunctionType="UniformDose",
                    RoiName="dummyPTV",
                    IsConstraint=False,
                    RestrictAllBeamsIndividually=False,
                    RestrictToBeam=None,
                    IsRobust=False,
                    RestrictToBeamSet=None,
                    )

                if rx is not None:
                    function.DoseFunctionParameters.DoseLevel = rx.DoseValue
                function.DoseFunctionParameters.Weight = 90

            # set the arc spacing and create the segments
            #with CompositeAction("Dummy optimization"):
            po.OptimizationParameters.Algorithm.MaxNumberOfIterations = 2
            po.OptimizationParameters.DoseCalculation.IterationsInPreparationsPhase = \
                1
            for (i, settings) in \
                enumerate(originalPlan.PlanOptimizations[opoIndex].OptimizationParameters.TreatmentSetupSettings[0].BeamSettings):
                replan.PlanOptimizations[poIndex].OptimizationParameters.TreatmentSetupSettings[0].BeamSettings[i].ArcConversionPropertiesPerBeam.FinalArcGantrySpacing = \
                    settings.ArcConversionPropertiesPerBeam.FinalArcGantrySpacing
            try:
                po.RunOptimization()
            except Exception as e:  # External geometry may have holes
                if search("The external ROI contains holes", str(e)):
                    # Ask if user wants to remove holes and retry the computation
                    res = MessageBox.Show("Dose could not be computed on image set '{}' because the external geometry contains holes. Would you like to Remove Holes and try again?".format(replanCTName), "Holy External!", MessageBoxButtons.YesNo)
                    if res == DialogResult.Yes:
                        case.PatientModel.StructureSets[replanCTName].SimplifyContours(RoiNames=[ext.Name], RemoveHoles3D=True, RemoveSmallContours=False, ReduceMaxNumberOfPointsInContours=False, ResolveOverlappingContours=False)
                        po.RunOptimization()

            poIndex += 1

            # lets copy the old segments to the new segments
            for (i, beam) in enumerate(bs.Beams):
                rpbs.Beams[i].BeamMU = beam.BeamMU

                for (j, s) in enumerate(beam.Segments):
                    rpbs.Beams[i].Segments[j].LeafPositions = \
                        s.LeafPositions
                    rpbs.Beams[i].Segments[j].JawPositions = s.JawPositions
                    rpbs.Beams[i].Segments[j].DoseRate = s.DoseRate
                    rpbs.Beams[i].Segments[j].RelativeWeight = \
                        s.RelativeWeight

    # lets compute the beams
    for bs in replan.BeamSets:
        if bs.Modality != "Electrons":
            if bs.Beams.Count != 0:
                if any(b.Segments.Count == 0 for b in bs.Beams):
                    MessageBox.Show("Cannot compute dose because beam(s) do not have valid segments.", "No Segments")
                    sys.exit()
                else:
                    bs.ComputeDose(ComputeBeamDoses=True, DoseAlgorithm="CCDose", ForceRecompute=True)
        else:       
            MessageBox.Show("Electron BeamSet - set histories, prescription and compute.", "Electron Beam Set")

    # Delete dummy PTV
    dummy.DeleteRoi()

    return replanName
