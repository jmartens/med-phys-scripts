t_path = r"\\vs20filesvr01\groups\CANCER\Physics"  # T: drive


# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))
from collections import OrderedDict
from re import IGNORECASE, match, search, sub

import pandas as pd  # Clinical Goals template data is read in as DataFrame
from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *

from CopyPlanWithoutChangesScript import copy_plan_without_changes


# Global so multiple functions can easily access
case = plan = beam_set = rx = rx_val = fx = None
is_sabr = False

# General and specific ROI names
# Some goals in templates have "general" ROI names (e.g., "Stomach and intestines") that correspond to multiple possible "specific" ROI names
# Specific names from original TG-263 nomenclature spreadsheet (not the one w/ CRMC-created names)
specific_rois = {
    "Aorta and major vessels": ["A_Aorta", "A_Aorta_Asc", "A_Aorta_Desc", "A_Coronary", "A_Pulmonary", "GreatVes", "V_Pulmonary", "V_Venacava", "V_Venacava_I", "V_Venacava_S"],
    "Bowel_Large": ["Anus", "Bowel_Large", "Colon", "Colon_Ascending", "Colon_Descending", "Colon_Sigmoid", "Colon_Transverse", "Rectum"],
    "Bowel_Small": ["Bowel_Small", "Duodenum", "Ileum", "Jejunum", "Jejunum_Ileum"]
}
specific_rois["Bowel"] = ["Bag_Bowel", "Bowel", "Spc_Bowel"] + specific_rois["Bowel_Large"] + specific_rois["Bowel_Small"]
specific_rois["Stomach and intestines"] = specific_rois["Bowel"] + ["Stomach"]


def format_list(l):
    # Helper function that returns a nicely formatted string of elements in a list
    # E.g., format_list(["A", "B", None]) -> "A, B, and None"

    if len(l) == 1:
        return l[0]
    if len(l) == 2:
        return "{} and {}".format(l[0], l[1])
    l_ = ["'{}'".format(item) if isinstance(item, str) else item for item in l]
    return "{}, and {}".format(", ".join(l_[:-1]), l_[-1])


def format_warnings(warnings_dict):
    # Helper function that nicely formats a dictionary of strings into one long string
    # E.g., format_warnings({"A": ["B", "C"], "D": ["E"]}) ->
    #     -  A
    #         -  B
    #         -  C
    #     -  D
    #         - E

    warnings_str = ""
    for k, v in warnings_dict.items():
        warnings_str += "\n\t-  {}".format(k)
        for val in sorted(list(set(v))):
            warnings_str += "\n\t\t-  {}".format(val)
    warnings_str += "\n"
    return warnings_str


class AddClinicalGoalsForm(Form):
    # Form that allows user to select MD, treatment technique, and body site from a GUI, to be used as template selection criteria
    # User also selects the types of template to apply (e.g., clinical goals)

    def __init__(self, data):
        self.Text = "Add Clinical Goals"  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        y = 15  # Vertical coordinate of next control

        self.data = data
        #self.data = OrderedDict([(name, data[name]) for name in sorted(data)])

        # Remove existing goals?
        self.clear_existing_cb = CheckBox()
        self.clear_existing_cb.AutoSize = True
        self.clear_existing_cb.Checked = True
        self.clear_existing_cb.Location = Point(15, y)
        self.clear_existing_cb.Text = "Clear existing goals"
        self.Controls.Add(self.clear_existing_cb)
        y += self.clear_existing_cb.Height + 15

        ## ListBox

        # Default selected template depends on current plan type
        default_template = "Mobius Conventional"
        if fx is not None and rx_val is not None and beam_set.PlanGenerationTechnique == "Imrt" and beam_set.DeliveryTechnique == "DynamicArc" and rx_val / fx >= 600:  # SABR/SBRT/SRS >=6 Gy/fx
            if fx == 1:
                default_template = "Mobius 1 Fx SRS"
            elif fx == 3:
                default_template = "Mobius 3 Fx SBRT"
            elif fx == 5:
                default_template = "Mobius 5 Fx SBRT"

        # Create and populate ListBox
        self.choose_templates_lb = ListBox()
        self.choose_templates_lb.Height = self.choose_templates_lb.PreferredHeight
        self.choose_templates_lb.Location = Point(15, y)
        self.choose_templates_lb.SelectionMode = SelectionMode.MultiExtended
        self.choose_templates_lb.Items.AddRange([name for name in self.data.keys() if not name.endswith("DNU")])
        selected_idx = list(self.data.keys()).index(default_template)
        self.choose_templates_lb.SetSelected(selected_idx - 1, True)
        self.choose_templates_lb.Location = Point(15, y)
        self.choose_templates_lb.Width = max(TextRenderer.MeasureText(template_name, self.choose_templates_lb.Font).Width for template_name in self.data.keys())
        self.choose_templates_lb.Height = self.choose_templates_lb.PreferredHeight
        self.choose_templates_lb.Visible = True
        self.template_names = OrderedDict([(default_template, self.data[default_template])])
        self.Controls.Add(self.choose_templates_lb)
        y += self.choose_templates_lb.Height + 15

        # OK button
        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Location = Point(15, y)
        self.ok.Text = "OK"
        self.AcceptButton = self.ok
        self.Controls.Add(self.ok)

        self.ShowDialog()  # Launch window

    def set_ok_enabled(self, sender=None, event=None):
        # Enable or disable "OK" button
        # Enable only if at least one template is selected

        self.ok.Enabled = self.choose_templates_lb.SelectedItems.Count > 0

    def ok_clicked(self, sender, event):
        # Event handler for "OK" button click

        self.DialogResult = DialogResult.OK


def add_clinical_goals(gui=True, **kwargs):
    """Apply clinical goals template(s) to plan

    If plan is approved, allow user to either exit script or add goals to a copy of the plan

    Read clinical goals from "template" (worksheet) in Excel workbook "Clinical Goals.xlsx"

    Default selected template name is the Mobius template that matches the current plan type

    Goals that depend on the Rx use the PrimaryDosePrescription for the beam set

    Add spreadsheet goals as well as:
    * For all the beam set's Rxs to volume of PTV:
        - D95%, V95%, D100%, and V100% goals for the Rx PTV
        - If the Rx PTV is derived from a CTV, D100% and V100% for that CTV
    * If Rx(s) exist, Dmax (D0.03) (from PrimaryDosePrescription)
    These goals have no numbered priority.

    Notes column in Excel file may specify:
    * Body site, if template applies to multiple body sites
    * Rx, if template applies to multiple Rx's
    * "Ipsilateral" or "Contralateral"
    * Other info unused by this function

    If Rx's are specified in a template name, only the goals with the plan Rx are applied. If plan Rx does not match any of the Rx's in the template, user may choose to scale goals to plan Rx and, if applicable, the template Rx's to use.
    
    If there is ambiguity in which template(s) to apply, user chooses template(s) from a GUI.
    
    Any template name ending in "DNU" is ignored.
    
    Add comments to plan saying which goals were applied. E.g.:
    Clinical Goals template(s) were applied:
    1. Mobius 5 Fx SBRT
    2. RTOG 0813 SBRT Lung

    If a GUI is used, visualization priority in the Clinical Goals list matches the number of the template in the plan comments.
    Otherwise, visualization priority does not exist unless the goal contains a relative volume. Since RS does not support clinical goals with relative volume, an absolute volume is used, and the priority is 1 to indicate that the template should be reapplied every time the geometry is changed.
    If goal appears in multiple templates, priority is the first template (alphabetically) in which it exists.

    Positional Arguments
    --------------------
    gui: bool
        True if the user should choose settings from a GUI, False otherwise
        If True, keyword arguments are ignored
        Defaults to True

    Keyword Arguments
    -----------------
    clear_existing: bool
        True if all existing clinical goals should be cleared before applying template(s), False otherwise
        Defaults to True
    template_names: List[str]
        List of names of clinical goals templates to apply
        Defaults to all

    Assumptions
    -----------
    No plan contains multiple nodal PTVs.
    Primary Rx is not to nodal PTV.
    Nodal PTV name contains "PTVn".
    In the clinical goals spreadsheet, Rx to primary PTV is "Rx" or "Rxp", and Rx to nodal PTV is "Rxn".
    Keyword argument `gui` is False only if this function is called from another script.
    """

    global case, plan, beam_set, rx, rx_val, fx

    # Get current variables
    try:
        beam_set = get_current("BeamSet")
    except:
        MessageBox.Show("The current plan has no beam sets. Click OK to abort script.", "No Beam Sets")
        sys.exit(1)
    patient = get_current("Patient")
    case = get_current("Case")
    plan = get_current("Plan")

    if plan.Review is not None and plan.Review.ApprovalStatus == "Approved":
        res = MessageBox.Show("Plan is approved, so clinical goals cannot be added. Would you like to add goals to a copy of the plan?", "Plan Is Approved", MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit(1)
        new_plan_name = copy_plan_without_changes()
        patient.Save()
        case.TreatmentPlans[new_plan_name].SetCurrent()
        plan = get_current("Plan")
        plan.BeamSets[beam_set.DicomPlanLabel].SetCurrent()
        beam_set = get_current("BeamSet")

    struct_set = plan.GetStructureSet()  # Geometries on the planning exam

    # Ensure that this is a photon beam set
    if beam_set.Modality != "Photons":
        MessageBox.Show("The current beam set is not a photon beam set. Click OK to abort script.", "Incorrect Modality")
        sys.exit(1)  # Exit with an error

    # Ensure that beam set machine is commissioned
    machine = beam_set.MachineReference.MachineName
    if get_current("MachineDB").GetTreatmentMachine(machineName=machine) is None:
        MessageBox.Show("Machine '{}' is uncommissioned. Click OK to abort script.".format(machine), "Uncommissioned Machine")
        sys.exit(1)

    warnings = ""  # Warnings to display at end of script (if there were any)

    # Rx and # fx
    if gui:
        rx = beam_set.Prescription
        if rx is not None:
            rx = rx.PrimaryDosePrescription
            rx_val = rx.DoseValue
        fx = beam_set.FractionationPattern
        if fx is not None:
            fx = fx.NumberOfFractions
    else:
        rx_val = 100

    # Read data from "Clinical Goals" spreadsheet
    # Read all sheets, ignoring "Planning Priority" column
    # Dictionary of sheet name : DataFrame
    filename = r"{}\Scripts\Data\Clinical Goals.xlsx".format(t_path)
    data = pd.read_excel(filename, sheet_name=None, engine="openpyxl", usecols=["ROI", "Goal", "Notes"])  # Default xlrd engine does not support xlsx
    
    # Get options from user
    if gui:
        form = AddClinicalGoalsForm(data)
        if form.DialogResult != DialogResult.OK:  # "OK" button was not clicked
            sys.exit()
        clear_existing = form.clear_existing_cb.Checked
        template_names = list(form.choose_templates_lb.SelectedItems)
    else:
        clear_existing = kwargs.get("clear_existing", True)
        template_names = kwargs.get("template_names", list(data.keys()))
        template_names = [name for name in template_names if name in data.keys() and not name.endswith("DNU")]

    # Text of checked RadioButtons
    # Clear existing Clinical Goals
    if clear_existing:
        with CompositeAction("Clear Clinical Goals"):
            while plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count > 0:
                plan.TreatmentCourse.EvaluationSetup.DeleteClinicalGoal(FunctionToRemove=plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[0])

    # If Rx is specified, add Dmax goal
    if rx_val is not None:
        ext = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]  # Select external ROI
        if ext:  # If there is an external (there will only be one), add Dmax goal
            d_max = 1.25 if is_sabr else 1.1
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ext[0], GoalCriteria="AtMost", GoalType="DoseAtAbsoluteVolume", ParameterValue=0.03, AcceptanceLevel=d_max * rx_val)  # e.g., D0.03 < 4400 for 4000 cGy non-SBRT plan
            except:  # Clinical goal already exists
                pass
        if gui:
            dose_rxs = [(dose_rx.DoseValue, dose_rx.OnStructure.Name) for dose_rx in beam_set.Prescription.DosePrescriptions if dose_rx.PrescriptionType == "DoseAtVolume" and dose_rx.OnStructure.Type == "Ptv"]
        else:
            dose_rxs = [(rx_val, "PTV")]
        for dose_rx in dose_rxs:
            # If Rx is to volume of PTV, add PTV D95%, V95%, D100%, and V100%
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=dose_rx[1], GoalCriteria="AtLeast", GoalType="DoseAtVolume", ParameterValue=0.95, AcceptanceLevel=dose_rx[0])  # D95% >= Rx
            except:  # Clinical goal already exists
                pass
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=dose_rx[1], GoalCriteria="AtLeast", GoalType="VolumeAtDose", ParameterValue=0.95 * dose_rx[0], AcceptanceLevel=1)  # V95% >= 100%
            except:
                pass
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=dose_rx[1], GoalCriteria="AtLeast", GoalType="VolumeAtDose", ParameterValue=dose_rx[0], AcceptanceLevel=0.95)  # V100% >= 95%
            except:
                pass
            try:
                plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=dose_rx[1], GoalCriteria="AtLeast", GoalType="DoseAtVolume", ParameterValue=1, AcceptanceLevel=0.95 * dose_rx[0])  # D100% >= 95%
            except:
                pass
            # If PTV is derived from CTV, add CTV D100% and V100%
            if gui:
                ptv = case.PatientModel.RegionsOfInterest[dose_rx[1]]
                if ptv.DerivedRoiExpression is not None:
                    ctvs = [r for r in struct_set.RoiGeometries[dose_rx[1]].GetDependentRois() if case.PatientModel.RegionsOfInterest[r].Type == "Ctv"]
                else:
                    ctvs = []
            else:
                ctvs = ["CTV"]
            for ctv in ctvs:
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ctv, GoalCriteria="AtLeast", GoalType="DoseAtVolume", ParameterValue=1, AcceptanceLevel=dose_rx[0])  # D100% >= 100%
                except:
                    pass
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(RoiName=ctv, GoalCriteria="AtLeast", GoalType="VolumeAtDose", ParameterValue=dose_rx[0], AcceptanceLevel=1)  # V100% >= 100%
                except:
                    pass

    # Information that will be displayed as warnings later
    # All invalid goals are in format "<ROI name>: <goal>", e.g., "Liver:  V21 Gy < (v-700) cc", except ipsi/contra goals ("<ROI name>: <goal> <Ipsilateral|Contralateral>")
    invalid_goals = OrderedDict()  # Goals in template that are in an invalid format
    empty_spare = OrderedDict()  # Volume-to-spare goals that cannot be added due to empty geometry
    lg_spare_vol = OrderedDict()  # Volume to spare is larger than ROI volume
    no_ipsi_contra = OrderedDict()  # Whether or not ipsilateral/contralateral goals could not be added due to indeterminable Rx side
    no_nodal_ptv = OrderedDict()  # If a nodal PTV does not exist, goals for a nodal PTV 

    ## Determine Rx side, used when adding ispi/contra goals
    
    # Get initial laser isocenter, just in case it is needed
    ini_laser_iso = [poi for poi in case.PatientModel.PointsOfInterest if poi.Type == "InitialLaserIsocenter"]
    rx_ctr = struct_set.PoiGeometries[ini_laser_iso[0].Name].Point.x if ini_laser_iso else None
    
    if hasattr(rx, "OnStructure"):
        struct = rx.OnStructure
        if struct.OrganData is not None and struct.OrganData.OrganType == "Target":  # Rx is to ROI
            rx_ctr = struct_set.RoiGeometries[struct.Name].GetCenterOfRoi().x  # R-L center of ROI
        else:  # Rx is to POI
            rx_ctr = struct_set.PoiGeometries[struct.Name].Point.x
    elif hasattr(rx, "OnDoseSpecificationPoint"):  # Rx is to site
        if rx.OnDoseSpecificationPoint is not None:  # Rx is to DSP
            rx_ctr = rx.OnDoseSpecificationPoint.Coordinates.x
        else:  # Rx is to site that is not a DSP
            dose_dist = plan.TreatmentCourse.TotalDose
            if dose_dist.DoseValues is not None and dose_dist.DoseValues.DoseData is not None:
                rx_ctr = dose_dist.GetCoordinateOfMaxDose().x

    ## Apply templates
    for template_name in template_names:
        goals = data[template_name]

        # "Fine-tune" the goals to apply
        # Check fractionation, Rx, body site, side, etc.

        ## Add goals   
                
        goals["ROI"] = pd.Series(goals["ROI"]).fillna(method="ffill")  # Autofill ROI name (due to vertically merged cells in spreadsheet)
    
        invalid_goals_template, empty_spare_template, lg_spare_vol_template, no_ipsi_contra_template, no_nodal_ptv_template = [], [], [], [], []
        roi_regex = "^{}(_[LR])?(\^.+)?( \(\d+\))?$"
        for _, row in goals.iterrows():  # Iterate over each row in DataFrame
            args = {}  # dict of arguments for ApplyTemplates
            roi = row["ROI"]  # e.g., "Lens"
            rois = []
            for r in case.PatientModel.RegionsOfInterest:
                if match(roi_regex.format(roi), r.Name, IGNORECASE) or (roi in specific_rois and any(match(roi_regex.format(specific_roi), r.Name, IGNORECASE) for specific_roi in specific_rois[roi])):
                    rois.append(r.Name)
                else:
                    # See if ROI in template is a PRV but doesn't specify a numerical expansion
                    prv_in_roi = search("PRV", roi, IGNORECASE)
                    if prv_in_roi is None or prv_in_roi.end() != len(roi):
                        continue
                    prv_in_r = search("PRV", r.Name, IGNORECASE)
                    if prv_in_r is None:
                        continue
                    roi_base = roi[:prv_in_roi.start()].strip("_")
                    r_base = r.Name[:prv_in_r.start()].strip("_")
                    if roi_base.lower() != r_base.lower():
                        continue
                    rois.append(r.Name)
            if not rois:  # ROI in goal does not exist in case
                continue

            goal = sub("\s", "", row["Goal"])  # Remove spaces in goal
            invalid_goal = "{}:\t{}".format(roi, row["Goal"])

            # If present, notes may be Rx, body site, body side, or info irrelevant to script
            notes = row["Notes"]
            if not pd.isna(notes):  # Notes exist
                # Goal only applies to specific Rx
                m = match("([\d\.]+) Gy", notes)
                if m is not None: 
                    notes = int(float(m.group(1)) * 100)  # Extract the number and convert to cGy
                    if rx != notes:  
                        continue

                # Goal only applies to certain Fx(s)
                elif notes.endswith("Fx"):
                    fxs = [int(elem.strip(",")) for elem in notes[:-3].split(" ") if elem.strip(",").isdigit()]
                    if fx not in fxs:
                        continue
                
                # Ipsilateral objects have same sign on x-coordinate (so product is positive); contralateral have opposite signs (so product is negative)
                elif notes in ["Ipsilateral", "Contralateral"]:
                    if rx_ctr is None:
                        no_ipsi_contra_template.append("{} ({})".format(invalid_goal, notes))
                    else:
                        rois = [r for r in rois if (notes == "Ipsilateral" and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x > 0) or (notes == "Contralateral" and rx_ctr * struct_set.RoiGeometries[r].GetCenterOfRoi().x < 0)]  # Select the ipsilateral or contralateral matching ROIs
                # Otherwise, irrelevant info

            # Visualization Priority (note that this is NOT the same as planning priority)
            if gui:
                """
                if pd.isna(row["Visualization Priority"]):
                    if pd.isna(row["Planning Priority"]):
                        args["Priority"] = 1
                    else:
                        args["Priority"] = row["Planning Priority"]
                else:
                    args["Priority"] = row["Visualization Priority"]
                """
                args["Priority"] = template_names.index(template_name) + 1
            
            ## Parse dose and volume amounts from goal. Then add clinical goal for volume or dose.

            # Regexes to match goal
            dose_amt_regex = """(
                                    (?P<dose_pct_rx>[\d.]+%)?
                                    (?P<dose_rx>Rx[pn]?)|
                                    (?P<dose_amt>[\d.]+)
                                    (?P<dose_unit>c?Gy)
                                )"""  # e.g., 95%Rx or 20Gy
            dose_types_regex = "(?P<dose_type>max|min|mean|median)"
            vol_amt_regex = """(
                                    (?P<vol_amt>[\d.]+)
                                    (?P<vol_unit>%|cc)|
                                    (\(v-(?P<spare_amt>[\d.]+)\)cc)
                            )"""  # e.g., 67%, 0.03cc, or v-700cc
            sign_regex = "(?P<sign><|>)"  # > or <

            dose_regex = """D
                            ({}|{})
                            {}
                            {}""".format(dose_types_regex, vol_amt_regex, sign_regex, dose_amt_regex)  # e.g., D0.03cc<110%Rx, Dmedian<20Gy

            vol_regex = """V
                            {}
                            {}
                            {}
                        """.format(dose_amt_regex, sign_regex, vol_amt_regex)  # e.g., V20Gy<67%

            # Need separate regexes b/c we can't have duplicate group names in a single regex
            # Remove whitespace from regex (left in above for readability) before matching
            vol_match = match(sub("\s", "", vol_regex), goal)
            dose_match = match(sub("\s", "", dose_regex), goal)
            m = vol_match if vol_match is not None else dose_match  # If it's not a volume, should be a dose

            if not m:  # Invalid goal format -> add goal to invalid goals list and move on to next goal
                invalid_goals_template.append(invalid_goal)
                continue

            args["GoalCriteria"] = "AtMost" if m.group("sign") == "<" else "AtLeast"  # GoalCriteria depends on sign

            # Extract dose: an absolute amount or a % of Rx
            dose_rx = m.group("dose_rx")
            if dose_rx:  # % of Rx
                if rx_val is None:
                    continue
                if not gui:
                    args["Priority"] = 1
                dose_pct_rx = m.group("dose_pct_rx")  # % of Rx
                if dose_pct_rx is None:  # Group not present. % of Rx is just specified as "Rx"
                    dose_pct_rx = 100
                else:  # A % is specified, so make sure format is valid
                    try:
                        dose_pct_rx = float(dose_pct_rx[:-1])  # Remove percent sign and convert to float
                        if dose_pct_rx < 0:  # % out of range -> add goal to invalid goals list and move on to next goal
                            invalid_goals_template.append(invalid_goal)
                            continue
                    except:  # % is non numeric -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                # Find appropriate Rx (to primary or nodal PTV)
                if gui:
                    if dose_rx == "Rxn":  # Use 2ry Rx (to nodal PTV)
                        rx_n = [rx_n for rx_n in beam_set.Prescription.DosePrescriptions if "PTVn" in rx_n.OnStructure.Name]
                        if rx_n:  # Found a nodal PTV (should never be more than one)
                            dose_rx = rx_n[0]
                        else:  # There is no nodal PTV, so add goal to list of goals that could not be added to nodal PTV, and move on to next goal
                            no_nodal_ptv_template.append(invalid_goal)
                    else:  # Primary Rx
                        dose_rx = rx
                else:
                    dose_rx = rx_val
                dose_amt = dose_pct_rx / 100 * rx_val  # Get absolute dose based on % Rx
            else:  # Absolute dose
                try:
                    dose_amt = float(m.group("dose_amt"))  # Account for scaling to template Rx if user selected this option (remember that `scaling_factor` is 1 otherwise)
                except:  # Given dose amount is non-numeric  -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue
                if m.group("dose_unit") == "Gy":  # Covert dose from Gy to cGy
                    dose_amt *= 100
            if dose_amt < 0 or dose_amt > 100000:  # Dose amount out of range  -> add goal to invalid goals list and move on to next goal
                invalid_goals_template.append(invalid_goal)
                continue

            # Extract volume: an absolute amount, a % of ROI volume, or an absolute amount to spare
            dose_type = vol_unit = spare_amt = None
            vol_amt = m.group("vol_amt")
            if vol_amt:  # Absolute volume or % of ROI volume
                try:
                    vol_amt = float(vol_amt)
                except:  # Given volume is non numeric -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue

                vol_unit = m.group("vol_unit")
                if vol_unit == "%":  # If relative volume, adjust volume amount
                    if vol_amt > 100:  # Given volume is out of range -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    vol_amt /= 100  # Convert percent to proportion

                if vol_amt < 0 or vol_amt > 100000:  # Volume amount out of range supported by RS -> add goal to invalid goals list and move on to next goal
                    invalid_goals_template.append(invalid_goal)
                    continue
            else:  # Volume to spare or dose type
                spare_amt = m.group("spare_amt")
                if spare_amt:  # Volume to spare
                    try:
                        spare_amt = float(spare_amt)
                    except:  # Volume amount is non numeric -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    
                    geom = struct_set.RoiGeometries[roi]
                    if not geom.HasContours():  # Cannot add volume to spare goal for empty geometry -> add goal to list of vol-to-spare goals for empty geometries
                        empty_spare_template.append(invalid_goal)
                        continue
                    if spare_amt < 0:  # Negative spare amount -> add goal to invalid goals list and move on to next goal
                        invalid_goals_template.append(invalid_goal)
                        continue
                    if spare_amt > geom.GetRoiVolume():
                        lg_spare_vol_template.append(invalid_goal)
                        continue
                
                    if not gui:
                        args["Priority"] = 1  # Hack: priority 1 means goal should be changed after template is applied
               
                else:  # Dose type: Dmax, Dmean, or Dmedian
                    dose_type = m.group("dose_type")

            # D...
            if goal.startswith("D"):
                # Dmax = D0.035
                if dose_type == "max":
                    args["GoalType"] = "DoseAtAbsoluteVolume"
                    args["ParameterValue"] = 0.03
                # Dmin = Max volume at that dose is everything but 0.035 cc
                elif dose_type == "min":
                    args["GoalType"] = "AbsoluteVolumeAtDose"
                    args["ParameterValue"] = dose_amt
                    args["AcceptanceLevel"] = struct_set.RoiGeometries[roi].GetRoiVolume() - 0.035
                # Dmean => "AverageDose"
                elif dose_type == "mean":
                    args["GoalType"] = "AverageDose"
                # Dmedian = D50%
                elif dose_type == "median":
                    args["GoalType"] = "DoseAtVolume"
                    args["ParameterValue"] = 0.5
                # Absolute or relative dose
                else:
                    args["ParameterValue"] = vol_amt
                    if vol_unit == "%":
                        args["GoalType"] = "DoseAtVolume"
                    else:
                        args["GoalType"] = "DoseAtAbsoluteVolume"
                args["AcceptanceLevel"] = dose_amt
            # V...
            else:
                args["ParameterValue"] = dose_amt
                if vol_unit == "%":
                    args["GoalType"] = "VolumeAtDose"
                else:
                    args["GoalType"] = "AbsoluteVolumeAtDose"
                if not spare_amt:
                    args["AcceptanceLevel"] = vol_amt
                
            # Add Clinical Goals
            for roi in rois:
                roi_args = args.copy()
                roi_args["RoiName"] = roi
                if spare_amt:
                    if gui:
                        total_vol = struct_set.RoiGeometries[roi].GetRoiVolume()
                        roi_args["AcceptanceLevel"] = total_vol - spare_amt
                    else:
                        roi_args["AcceptanceLevel"] = spare_amt
                        args["Priority"] = 1
                try:
                    plan.TreatmentCourse.EvaluationSetup.AddClinicalGoal(**roi_args)
                except:
                    pass

        if invalid_goals_template:
            invalid_goals[template_name] = invalid_goals_template
        if empty_spare_template:
            empty_spare[template_name] = empty_spare_template
        if lg_spare_vol_template:
            lg_spare_vol[template_name] = lg_spare_vol_template
        if no_ipsi_contra_template:
            no_ipsi_contra[template_name] = no_ipsi_contra_template
        if no_nodal_ptv_template:
            no_nodal_ptv[template_name] = no_nodal_ptv_template

    # Add warnings about clinical goals that were not added
    if invalid_goals:
        warnings += "The following clinical goals could not be parsed so were not added:"
        warnings += format_warnings(invalid_goals)
    if empty_spare:
        warnings += "The following clinical goals could not be added due to empty geometries:"
        warnings += format_warnings(empty_spare)
    if lg_spare_vol:
        warnings += "The following clinical goals could not be added because the volume to spare is larger than the ROI volume:"
        warnings += format_warnings(lg_spare_vol)
    if no_ipsi_contra:
        warnings += "There is no Rx, so ipsilateral/contralateral structures could not be determined. Ipsilateral/contralateral clinical goals were not added:"
        warnings += format_warnings(no_ipsi_contra)
    if no_nodal_ptv:
        warnings += "No nodal PTV was found, so the following clinical goals were not added:"
        warnings += format_warnings(no_nodal_ptv)

    # Add template names to plan comments
    if gui:
        new_comments = "Clinical Goals template(s) were applied:\n{}".format("\n".join(["{}. {}".format(template_names.index(name) + 1, name) for name in template_names]))
        if plan.Comments == "":
            plan.Comments = new_comments
        else:
            plan.Comments = "{}\n{}".format(plan.Comments, new_comments)

    # Display warnings if there were any
    if warnings != "":
        MessageBox.Show(warnings, "Warnings")
        
    if gui:
        sys.exit()  # For some reason, script won't exit on its own if warnings are displayed
