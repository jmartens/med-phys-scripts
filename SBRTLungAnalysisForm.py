"""Generate a PDF report of selected RTOG 0813 statistics for SBRT lung plan(s)

Use function `sbrt_lung_analysis_` to both create an open the report (this is in RS script SBRTLungAnalysis)
Use function `sbrt_lung_analysis` to create, but not open, the report (other scripts call this function)
"""

# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import sys
from collections import OrderedDict
from re import search, sub

import pandas as pd  # Interpolation data from RTOG 0813 is read in as a DataFrame
from connect import *  # Interact w/ RS

# Report uses ReportLab to create a PDF
from reportlab.lib.colors import obj_R_G_B, toColor, black, blue, grey, green, lightgrey, orange, red, white, yellow
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.platypus.flowables import KeepTogether
from reportlab.platypus.tables import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

from scipy import interpolate  # Interpolate stats for a given PTV volume

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


# Don't set these yet so that sbrt_lung_analysis functions can be used by other scripts that don't necessarily need these variables
# But must be global so multiple functions have easy access
case = plan = beam_set = None
report_name = None
plan_names = OrderedDict()


# ReportLab Paragraph styles
styles = getSampleStyleSheet()  # Base styles (e.g., "Heading1", "Normal")
hdg = ParagraphStyle(name="hdg", fontName="Helvetica-Bold", fontSize=16, alignment=TA_CENTER)  # For patient name
subhdg = ParagraphStyle(name="subhdg", parent=styles["Normal"], fontName="Helvetica", fontSize=16, alignment=TA_CENTER)  # For MR# and "RTOG 0813 SBRT Lung Analysis"
tbl_title = ParagraphStyle(name="tbl_title", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=16)  # For table titles ("Plan Stats" and "Interpolated Cutoffs")
tbl_hdg = ParagraphStyle(name="tbl_hdg", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, alignment=TA_CENTER, leading=15)  # For table headings; leading necessary due to possible subscripts
tbl_data = ParagraphStyle(name="tbl_data", parent=styles["Normal"], fontName="Helvetica", fontSize=10, alignment=TA_CENTER, leading=15)  # For table cells

# ReportLab Spacer objects to reuse for nice formatting
width, _ = landscape(letter)  # Need the width (11") for Spacer width
spcr_sm = Spacer(width, 0.1 * inch)  # Small
spcr_md = Spacer(width, 0.2 * inch)  # Medium
spcr_lg = Spacer(width, 0.3 * inch)  # Large


def create_roi_if_absent(roi_name, roi_type):
    # CRMC standard ROI colors
    colors = pd.read_csv(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\TG263 Nomenclature with CRMC Colors.csv", index_col="TG263-Primary Name", usecols=["TG263-Primary Name", "Color"])["Color"]

    roi = get_latest_roi(roi_name, unapproved_only=True)
    if roi is None:
        try:
            color = colors[roi_name].replace(";", ",")  # E.g., "255; 1; 2; 3" -> "255, 1, 2, 3"
        except:
            color = "255, 255, 255, 255"
        roi_name = name_item(roi_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)
        roi = case.PatientModel.CreateRoi(Name=roi_name, Type=roi_type, Color=color)
    return roi


def get_latest_roi(base_roi_name, **kwargs):
    # Helper function that returns the ROI with the given "base name" and largest copy number
    # kwargs:
    # unapproved_only: If True, consider only the ROIs that are not part of any approved structure set in the case (defaults to False)
    # non_empty_only: If True, consider only the ROIs with geometries on the exam (defaults to False)
    # exam: If `nonempty_only` or `unapproved_only` are True, the exam on which to check for a geometry/approval (defaults to current exam)

    unapproved_only = kwargs.get("unapproved_only", False)
    non_empty_only = kwargs.get("non_empty_only", False)
    exam = kwargs.get("exam", get_current("Examination"))

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
            m = search(" \((\d+)\)$".format(base_roi_name), roi_name)
            if m:  # There is a copy number
                grp = m.group()
                length = min(16 - len(grp), len(base_roi_name))
                if roi_name[:length] == base_roi_name[:length]:
                    copy_num = int(m.group(1))
        if copy_num > latest_copy_num:
            latest_copy_num = copy_num
            latest_roi = roi
    return latest_roi


def doses_on_addl_set(exam_names=None):
    eval_doses = {}
    for fe in case.TreatmentDelivery.FractionEvaluations:
        for doe in fe.DoseOnExaminations:
            exam_name = doe.OnExamination.Name
            if exam_names is None or exam_name in exam_names:
                for de in doe.DoseEvaluations:
                    if not de.Name and hasattr(de, "ForBeamSet"):  # Ensure it's a dose on additional set
                        de_name = "{} on {}".format(de.ForBeamSet.DicomPlanLabel, exam_name)
                        eval_doses[de_name] = de
    return eval_doses


def get_text_color(bkgrd_color):
    # Helper function that returns the appropriate text color (black or white) from the given background color (a reportlab.lib.colors.Color object)

    r, g, b = obj_R_G_B(bkgrd_color)
    return white if r * 0.00117 + g * 0.0023 + b * 0.00045 <= 0.72941 else black


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


def create_iso_roi(dose_dist, idl):
    # Helper function that creates ROI/geometry from `idl` isodose in the given dose distribution
    # Return the ROI

    iso_roi_name = "IDL_{:.0f}".format(idl)
    iso_roi = create_roi_if_absent(iso_roi_name, "Control")
    iso_roi.CreateRoiGeometryFromDose(DoseDistribution=dose_dist, ThresholdLevel=idl)
    dose_dist.UpdateDoseGridStructures()  # Must run to update new geometry
    return iso_roi


class SBRTLungAnalysisForm(Form):
    """Form that allows the user to choose settings for generation of a PDF report w/ RTOG 0813 statistics

    The user chooses:
        - PTV (default is PTV in Rx of current beam set) 
        - RTOG 0813 statistics (default is all)
        - Plan(s) (default is current plan) (External, E-PTV_Ev20, and Lungs-CTV must be contoured on the planning exam)
        - Doses on additional sets, if present (default is none) (External, E-PTV_Ev20, and Lungs-CTV must be contoured on the additional set)
    """
    
    def __init__(self, ptv_names, default_ptv_name, stats, default_stats, plan_names, default_plan_names, eval_doses, default_eval_doses):
        self.plan_names, self.eval_doses = plan_names, eval_doses

        self.AutoSize = True  # Enlarge form to accommodate controls, if necessary
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # Disallow user resizing
        self.StartPosition = FormStartPosition.CenterScreen  # Launch window in middle of screen
        self.Text = "SBRT Lung Analysis"
        y = 15  # y-coordinate of next control

        # PTV radiobuttons
        self.ptv_gb = GroupBox()
        self.ptv_gb.AutoSize = True  # Enlarge GroupBox to accommodate radiobutton widths, if necessary
        self.ptv_gb.Location = Point(15, y)
        self.ptv_gb.Text = "PTV:"
        rb_y = 15
        for ptv_name in ptv_names:
            rb = RadioButton()
            rb.AutoSize = True  # Enlarge RadioButton to accommodate text width, if necessary
            rb.Click += self.ptv_name_clicked
            rb.Checked = ptv_name == default_ptv_name
            rb.Location = Point(15, rb_y)
            rb.Text = ptv_name
            self.ptv_gb.Controls.Add(rb)
            rb_y += rb.Height
        self.ptv_gb.Height = len(ptv_names) * 20
        self.Controls.Add(self.ptv_gb)
        y += self.ptv_gb.Height + 15
        
        # Add stats checkboxes
        self.stats_gb = GroupBox()
        self.stats_gb.AutoSize = True
        self.stats_gb.Location = Point(15, y)
        self.stats_gb.Text = "Stat(s):"
        cb_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = "Select all"
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        if len(default_stats) == len(stats):
            cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        elif not stats:
            cb.CheckState = cb.Tag = CheckState.Unchecked
        else:
            cb.CheckState = cb.Tag = CheckState.Indeterminate
        cb.Click += self.select_all_clicked
        self.stats_gb.Controls.Add(cb)
        cb_y += cb.Height
        for stat in stats:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)
            cb.Checked = stat in default_stats
            cb.Click += self.checkbox_clicked
            cb.Text = stat
            self.stats_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.stats_gb.Height = len(default_stats) * 20
        self.Controls.Add(self.stats_gb)
        y += self.stats_gb.Height + 15

        # Add plans checkboxes
        self.plans_gb = GroupBox()
        self.plans_gb.AutoSize = True
        self.plans_gb.Location = Point(15, y)
        self.plans_gb.Text = "Plan(s):"
        cb_y = 15
        cb = CheckBox()
        cb.AutoSize = True
        cb.Text = "Select all"
        cb.Location = Point(15, cb_y)
        cb.ThreeState = True
        if len(default_plan_names) == len(plan_names):
            cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
        elif not default_plan_names:
            cb.CheckState = cb.Tag = CheckState.Unchecked
        else:
            cb.CheckState = cb.Tag = CheckState.Indeterminate
        cb.Click += self.select_all_clicked
        self.plans_gb.Controls.Add(cb)
        cb_y += cb.Height
        for plan_name in self.plan_names:
            cb = CheckBox()
            cb.AutoSize = True
            cb.Location = Point(30, cb_y)
            cb.Checked = plan_name in default_plan_names
            cb.Click += self.checkbox_clicked
            cb.Text = plan_name
            self.plans_gb.Controls.Add(cb)
            cb_y += cb.Height
        self.plans_gb.Height = len(self.plan_names) * 20
        self.Controls.Add(self.plans_gb)
        y += self.plans_gb.Height + 15

        # Add eval doses (doses on additional set) checkboxes
        if self.eval_doses:  # There is dose computed on additional set
            # Add eval dose checkboxes
            self.eval_doses_gb = GroupBox()
            self.eval_doses_gb.AutoSize = True
            self.eval_doses_gb.Location = Point(15, y)
            self.eval_doses_gb.Text = "Dose(s) on additional set:"
            cb_y = 15
            cb = CheckBox()
            cb.AutoSize = True
            cb.Text = "Select all"
            cb.Location = Point(15, cb_y)
            cb.ThreeState = True
            if len(default_eval_doses) == len(self.eval_doses):
                cb.CheckState = cb.Tag = CheckState.Checked  # Tag attribute is used to keep track of previous check state, so next check state can be determined
            elif not default_eval_doses:
                cb.CheckState = cb.Tag = CheckState.Unchecked
            else:
                cb.CheckState = cb.Tag = CheckState.Indeterminate
            cb.Click += self.select_all_clicked
            self.eval_doses_gb.Controls.Add(cb)
            cb_y += cb.Height
            for name in self.eval_doses:
                cb = CheckBox()
                cb.AutoSize = True
                cb.Location = Point(30, cb_y)
                cb.Checked = name in default_eval_doses
                cb.Click += self.checkbox_clicked
                cb.Text = name
                self.eval_doses_gb.Controls.Add(cb)
                cb_y += cb.Height
            self.eval_doses_gb.Height = len(self.eval_doses) * 20
            self.Controls.Add(self.eval_doses_gb)
            y += self.eval_doses_gb.Height + 15

        # Add "Generate Report" button
        self.generate_btn = Button()
        self.generate_btn.AutoSize = True
        self.generate_btn.Click += self.generate_clicked
        self.generate_btn.Location = Point(15, y)
        self.generate_btn.Text = "Generate Report"
        self.AcceptButton = self.generate_btn
        self.Controls.Add(self.generate_btn)

        self.ptv_name_clicked()
        self.set_generate_enabled()
        self.ShowDialog()  # Launch window

    def ptv_name_clicked(self, sender=None, event=None):
        # Disable plan and eval dose names for which there is no PTV geometry
        ptv_name = [rb.Text for rb in self.ptv_gb.Controls if rb.Checked][0]
        for cb in list(self.plans_gb.Controls)[1:]:
            plan_name = cb.Text
            if case.TreatmentPlans[plan_name].GetStructureSet().RoiGeometries[ptv_name].HasContours():
                cb.Enabled = True
            else:
                cb.Enabled = cb.Checked = False
        self.checkbox_clicked(cb)
        if self.eval_doses:
            for cb in list(self.eval_doses_gb.Controls)[1:]:
                exam_name = cb.Text.split(" on ")[1]
                if case.PatientModel.StructureSets[exam_name].RoiGeometries[ptv_name].HasContours():
                    cb.Enabled = True
                else:
                    cb.Enabled = cb.Checked = False
            self.checkbox_clicked(cb)

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
        self.set_generate_enabled()

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
        self.set_generate_enabled()

    def set_generate_enabled(self):
        # Helper method that enables or disables "Generate Report" button
        # Enable button if a PTV is selected, at least one stat is selected, and at least one plan or eval dose is selected

        ptv_cked = any(cb.Checked for cb in self.ptv_gb.Controls)
        stat_cked = any(cb.Checked for cb in self.stats_gb.Controls)
        plan_cked = any(cb.Checked for cb in self.plans_gb.Controls)
        eval_dose_cked = any(cb.Checked for cb in self.eval_doses_gb.Controls) if hasattr(self, "eval_doses_gb") else False
        self.generate_btn.Enabled = ptv_cked and stat_cked and (plan_cked or eval_dose_cked)

    def generate_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


def sbrt_lung_analysis_():
    """Generate PDF report using `sbrt_lung_analysis`, open report, and exit script
    Use GUI
    """

    sbrt_lung_analysis(gui=True)

    # Open report
    reader_paths = [r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe", r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"]  # Paths to Adobe Reader on RS servers
    for reader_path in reader_paths:
        try:
            os.system(r'START /B "{}" "{}"'.format(reader_path, filename))
            break
        except:
            continue
        

def sbrt_lung_analysis(**kwargs):
    """Generate a PDF report with selected RTOG 0813 statistics for the selected plan(s) and/or dose(s) on additional set

    If a GUI is used, display errors in a dialog box and abort script.
    Otherwise, return the error message when it is generated (return None if the function completes)

    Keyword Arguments
    -----------------
    gui: bool
        True if the user should choose settings from a GUI, False otherwise
        Defaults to False
    ptv_name: str
        Name of the PTV to use for statistic calculations
        If `gui` is True, the name of the default checked PTV
        Defaults to the PTV that the current beam set's primary Rx is to
    stats: List[str]
        List of statistic names to compute
        If `gui` is True, the default checked stats
        Choose from the following:
            - "Traditional CI (R100%)"
            - "Paddick CI"
            - "GI (R50%)"
            - "D2cm [%]"
            - "Max dose @ appreciable volume [%]"
            - "V20Gy [%]"
            - "Max dose to External is inside PTV"
        Defaults to all of the above
    plan_names: List[str]
        List of plan names for which to compute the stats
        If `gui` is True, the default checked plan names
        Defaults to all plan names with normal tissue ("E-PTV_Ev20"), external, and "Lungs-CTV" geometries on their planning exams
    eval_dose_names: List[str]
        List of doses on additional set for which to compute the stats
        If `gui` is True, the default checked eval dose names
        Each eval dose name is in the format "<beam set name> on <exam name>" (e.g., "SBRT Rt Lung on "Inspiration 12.9.20")
        Defaults to all eval doses with normal tissue ("E-PTV_Ev20"), external, and "Lungs-CTV" geometries on their exams

    PDF report is in "T:\Physics\Scripts\Output Files\SBRTLungAnalysis"
    Report contains:
    - Table of color-coded computed stats for each selected plan and eval dose
    - If any interpolated stats were selected, table of RTOG 0813 interpolation stats with a row added for each plan and eval dose
    Assume no PTV volume > 163 cc (largest RTOG interpolation PTV volume)
    For simplicity and consistency, all colors are System.Drawing.Color objects, not reportlab.lib.colors.Color objects. They are converted to the latter only when necessary.
    """

    global case, plan, beam_set, report_name

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort script.", "No Patient Loaded")
        sys.exit(1)  # Exit script with an error
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
    try:
        beam_set = get_current("BeamSet")
    except:
        MessageBox.Show("There are no beam sets in the current plan. Click OK to abort script.", "No Beam Sets")
        sys.exit(1)  # Exit script with an error
    exam = plan.GetStructureSet().OnExamination

    gui = kwargs.get("gui", False)

    # Report name
    pt_name = "{}, {}".format(patient.Name.split("^")[0], patient.Name.split("^")[1])
    report_name = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\SBRTLungAnalysis\{} SBRT Lung Analysis.pdf".format(pt_name)
    
    # Ensure Rx exists
    rx = beam_set.Prescription.PrimaryDosePrescription
    if rx is None:
        msg = "Beam set has no prescription."
        if not gui:
            return msg
        MessageBox.Show(msg, "No Prescription")
        sys.exit(1)

    # Ensure beam set is SBRT lung
    fx_pattern = beam_set.FractionationPattern
    if beam_set.Modality != "Photons" or fx_pattern is None or fx_pattern.NumberOfFractions > 5 or rx.DoseValue < 600 * fx_pattern.NumberOfFractions or case.BodySite not in ["", "Thorax"]:
        MessageBox.Show("This is not an SBRT plan. Click OK to abort the script.", "SBRT Lung Analysis")
        sys.exit(1)

    # Exams w/ the necessary geometries (necessary to select plans and eval doses)
    exam_names = {}
    for e in case.Examinations:
        roi_geoms = case.PatientModel.StructureSets[e.Name].RoiGeometries
        normal_tissue = get_latest_roi("E-PTV_Ev20", exam=e, nonempty_only=True)
        ext = [geom.OfRoi.Name for geom in roi_geoms if geom.OfRoi.Type == "External"]
        lungs_ctv = get_latest_roi("Lungs-CTV", exam=e, nonempty_only=True)
        if lungs_ctv is None:
            lungs_ctv = get_latest_roi("Lungs-ITV", exam=e, nonempty_only=True)
        if normal_tissue is not None and ext and lungs_ctv is not None:
            exam_names[e.Name] = [normal_tissue.Name, ext[0], lungs_ctv.Name]

    # All PTV names
    all_ptv_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type.upper() == "PTV"]
    if not all_ptv_names:
        msg = "There are no PTVs in the current case."
        if not gui:
            return msg
        MessageBox.Show(msg, "No PTVs")
        sys.exit(1)
    
    # Default PTV name
    ptv_name = kwargs.get("ptv_name", rx.OnStructure.Name)

    # All plan stats
    all_stats = ["Traditional CI (R100%)", "Paddick CI", "GI (R50%)", "D2cm [%]", "Max dose @ appreciable volume [%]", "V20Gy [%]", "Max dose to External is inside PTV"]

    # Default plan stats
    stats = kwargs.get("stats", all_stats)

    # All plan names
    all_plan_names = [p.Name for p in case.TreatmentPlans if p.TreatmentCourse is not None and p.TreatmentCourse.TotalDose is not None and p.TreatmentCourse.TotalDose.DoseValues is not None and plan.GetStructureSet().OnExamination.Name in exam_names]
    if not all_plan_names:
        msg = "There are no plans that have dose, and a geometry for each E-PTV_Ev20, External, and either Lungs-CTV or Lungs-ITV."
        if not gui:
            return msg
        MessageBox.Show(msg, "No Viable Plans")
        sys.exit(1)

    # Default plan names
    if "plan_names" in kwargs:
        plan_names = [plan_name for plan_name in kwargs.get("plan_names") if plan_name in plan_names]
    else:
        plan_names = all_plan_names[:]

    # All eval doses
    all_eval_doses = doses_on_addl_set(exam_names)

    # Default eval doses
    if "eval_dose_names" in kwargs:
        eval_doses = {name: all_eval_doses[name] for name in kwargs.get("eval_dose_names") if name in all_eval_doses}
    else:
        eval_doses = all_eval_doses.copy()

    # Display GUI if applicable
    if gui:
        form = SBRTLungAnalysisForm(all_ptv_names, ptv_name, all_stats, stats, all_plan_names, plan_names, list(all_eval_doses.keys()), list(eval_doses.keys()))
        if form.DialogResult != DialogResult.OK:  # User exited GUI
            sys.exit()
        ptv_name = [rb.Text for rb in form.ptv_gb.Controls if rb.Checked][0]
        stats = [cb.Text for cb in list(form.stats_gb.Controls)[1:] if cb.Checked]
        plan_names = [cb.Text for cb in list(form.plans_gb.Controls)[1:] if cb.Checked]
        eval_doses = {cb.Text: all_eval_doses[cb.Text] for cb in list(form.eval_doses_gb.Controls)[1:] if cb.Checked} if all_eval_doses else {}

    # Read in data
    filename = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Data\RTOG0813.xlsx"
    data = pd.read_excel(filename, engine="openpyxl")  # RTOG 0813 stats for interpolation

    # Stats that will be interpolated
    interp_stats = [stat for stat in stats if stat in ["GI (R50%)", "D2cm [%]", "V20Gy [%]"]]

    ## Prepare PDF
    pdf = SimpleDocTemplate(report_name, pagesize=landscape(letter), bottomMargin=0.2 * inch, leftMargin=0.25 * inch, rightMargin=0.2 * inch, topMargin=0.2 * inch)  # 8.5 x 11", 0.2" top and bottom margin, 0.25" left and right margin

    # Heading
    hdr = Paragraph(pt_name, style=hdg)
    mrn = Paragraph("MR#: {}".format(patient.PatientID), style=subhdg)
    desc = Paragraph("RTOG 0813 SBRT Lung Analysis", style=subhdg)

    # Plan/Eval dose key is a ReportLab Table
    # A single row: orange square, "Plan" text, blue square, "Evaluation Dose" text
    # Include each key item only if some of that items are selected
    key_1_data, key_1_style = [], []
    if plan_names:
        key_1_data.extend(["", "Plan"])
        key_1_style.extend([("BACKGROUND", (0, 0), (0, 0), orange), ("BOX", (0, 0), (0, 0), 0.5, black)])  # 1st cell: orange background, black outline
        if eval_doses:
            key_1_data.extend(["", "Evaluation dose"]) 
            key_1_style.extend([("BACKGROUND", (2, 0), (2, 0), blue), ("BOX", (2, 0), (2, 0), 0.5, black)])  # 3rd cell: blue background, black outline
    elif eval_doses:
        key_1_data.extend(["", "Evaluation dose"])
        key_1_style.extend([("BACKGROUND", (0, 0), (0, 0), blue), ("BOX", (0, 0), (0, 0), 0.5, black)])  # 1st cell: blue background, black outline
    key_1 = Table([key_1_data], colWidths=[0.2 * inch, 1.25 * inch] * 2, rowHeights=[0.2 * inch], style=TableStyle(key_1_style), hAlign="LEFT")  # Left-align the table

    # Deviation key is a ReportLab Table
    key_2_data = ["", "No deviation", "", "Minor deviation", "", "Major deviation"]  # A single row: green square, "No deviation" text, yellow square, "Minor deviation" text, red square, "Major deviation" text
    key_2_style = [
        # 1st cell: green background, black outline
        ("BACKGROUND", (0, 0), (0, 0), green),
        ("BOX", (0, 0), (0, 0), 0.5, black),
        # 3rd cell: yellow background, black outline
        ("BACKGROUND", (2, 0), (2, 0), yellow),
        ("BOX", (2, 0), (2, 0), 0.5, black),
        # 5th cell: red background, black outline
        ("BACKGROUND", (4, 0), (4, 0), red),
        ("BOX", (4, 0), (4, 0), 0.5, black)
    ]
    key_2 = Table([key_2_data], colWidths=[0.2 * inch, 1.25 * inch] * 3, rowHeights=[0.2 * inch], style=TableStyle(key_2_style), hAlign="LEFT")  # Left-align the table

    ## Plan stats table

    # Table title
    plan_stats_title = Paragraph("Plan Stats", style=tbl_title)

    # Header row
    plan_stats_data = ["Plan / Evaluation dose", "PTV vol [cc]"] + stats
    plan_stats_data = [sub("(\d+(%|cm|Gy))", "<sub>\g<1></sub>", text) for text in plan_stats_data]  # Surround each subscript with HTML "<sub>" tags
    plan_stats_data = [[Paragraph(text, style=tbl_hdg) for text in plan_stats_data]]
    plan_stats_style = [  # Center-align, middle-align, and black outline for all cells. Gray background for header row
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 0), (-1, 0), lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE")
    ]

    ## Compute all PTV vols (necessary now so we know where in tables to place plan rows)
    # List of (plan / eval dose name, PTV vol)
    ptv_vols = {}
    for plan_name in plan_names:
        ptv_vol = case.TreatmentPlans[plan_name].GetStructureSet().RoiGeometries[ptv_name].GetRoiVolume()
        if ptv_vol in ptv_vols:
            ptv_vols[ptv_vol].append(plan_name)
        else:
            ptv_vols[ptv_vol] = [plan_name]
    for name in eval_doses:
        ptv_vol = case.PatientModel.StructureSets[name.split(" on ")[1]].RoiGeometries[ptv_name].GetRoiVolume()
        if ptv_vol in ptv_vols:
            ptv_vols[ptv_vol].append(name)
        else:
            ptv_vols[ptv_vol] = [name]
    ptv_vols = OrderedDict({vol: sorted(ptv_vols[vol]) for vol in sorted(ptv_vols)})
    
    ## Colors
    # Equally spaced hues based on number of hues needed
    # Constrain to brighter hues so that color is obvious 
    r, g, b = obj_R_G_B(orange)
    a = [0.502 + (0.9 - 0.502) / len(plan_names) * (len(plan_names) - i - 1) for i in range(len(plan_names))]  # A components of colors
    plan_colors = [toColor("rgba({}, {}, {}, {})".format(r, g, b, a_)) for a_ in a]  # Plan rows will be an orange hue

    r, g, b = obj_R_G_B(blue)
    a = [0.502 + (0.9 - 0.502) / len(eval_doses) * (len(eval_doses) - i - 1) for i in range(len(eval_doses))]  # A components of colors
    eval_dose_colors = [toColor("rgba({}, {}, {}, {})".format(r, g, b, a_)) for a_ in a]  # Eval dose rows will be an orange hue

    ## Interpolated cutoffs table

    # Table title
    interp_cutoffs_title = Paragraph("Interpolated Cutoffs", style=tbl_title)

    # Header row
    interp_cutoffs_data = ["PTV vol [cc]"]
    for stat in interp_stats:
        interp_cutoffs_data.extend(["{} None".format(stat), "{} Minor".format(stat)])
    interp_cutoffs_data = [[Paragraph(sub("(\d+(%|cm|Gy))", "<sub>\g<1></sub>", text), style=tbl_hdg) for text in interp_cutoffs_data]]  # Surround each subscript with HTML "<sub>" tags
    interp_cutoffs_style = plan_stats_style[:]  # Interp vals table has same style as plan stats table

    # Add plans and eval doses to plan stats table and interp cutoffs table
    plan_stats_idx = 0  # Row number in plan stats table
    for i, row in data.iterrows():
        if plan_stats_idx != len(plan_names) + len(eval_doses):  # Some plans / eval doses haven't yet been added to the tables
            ptv_vol, names = list(ptv_vols.items())[plan_stats_idx]
            if ptv_vol <= row["PTV vol [cc]"]:  # Plan / eval dose row goes immediately before this interp cutoffs row
                for name in names:
                    plan_stats_idx += 1
                    plan_stats_row = [name, round(ptv_vol, 2)]  # e.g., ["SBRT Lung", 38.89]
                    interp_cutoffs_row = [Paragraph(str(round(ptv_vol, 2)), style=tbl_data)]  # e.g., [38.89]
                
                    if name in plan_names:
                        color = plan_colors.pop()  # Get next plan row color
                        dose_dist = case.TreatmentPlans[name].TreatmentCourse.TotalDose
                        exam_name = case.TreatmentPlans[name].GetStructureSet().OnExamination.Name
                        rx = sum(bs.Prescription.PrimaryDosePrescription.DoseValue for bs in case.TreatmentPlans[name].BeamSets if bs.Prescription.PrimaryDosePrescription is not None)  # Sum of Rx's from all beam sets that have an Rx
                        v20_vol = 2000
                    else:  # Eval dose
                        # Dose distribution is fractional!
                        color = eval_dose_colors.pop()  # Get next eval dose row color
                        dose_dist = eval_doses[name]
                        exam_name = name.split(" on ")[1]
                        rx = float(dose_dist.ForBeamSet.Prescription.PrimaryDosePrescription.DoseValue) / dose_dist.ForBeamSet.FractionationPattern.NumberOfFractions
                        v20_vol = 2000.0 / dose_dist.ForBeamSet.FractionationPattern.NumberOfFractions
                    struct_set = case.PatientModel.StructureSets[exam_name]
                    text_color = get_text_color(color)  # Color text black or white (based on background color `color`)?

                    plan_stats_style.append(("BACKGROUND", (0, plan_stats_idx), (1, plan_stats_idx), color))
                    bk_color = None  # Background color for the individual stat value

                    # Compute each stat and add appropriately colored cell to plan stats table
                    for j, stat in enumerate(stats):
                        if stat == "Traditional CI (R100%)":
                            iso_roi = create_iso_roi(dose_dist, rx)  # Create ROI from 100% isodose
                            
                            plan_stat_val = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume() / ptv_vol  # Volume of 100% isodose geometry as proportion of PTV volume
                            plan_stat_val = round(plan_stat_val, 2)  # Round CI to 2 decimal places
                            bk_color = yellow if plan_stat_val < 1 else green if plan_stat_val <= 1.2 else yellow if plan_stat_val <= 1.5 else red
                            
                            iso_roi.DeleteRoi()

                        elif stat == "Paddick CI":
                            iso_roi = create_iso_roi(dose_dist, rx)  # Create ROI from 100% isodose
                            # Create intersection of PTV and 100% isodose
                            intersect_roi_name = "PTV&IDL_{}".format(int(rx))
                            intersect_roi = create_roi_if_absent(intersect_roi_name, "Control")
                            intersect_roi.CreateAlgebraGeometry(Examination=case.Examinations[exam_name], ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [ptv_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [iso_roi.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Intersection", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                            #dose_dist.UpdateDoseGridStructures()  # Must run to update new geometry

                            iso_roi_vol = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume()
                            intersect_roi_vol = struct_set.RoiGeometries[intersect_roi.Name].GetRoiVolume()
                            
                            plan_stat_val = intersect_roi_vol * intersect_roi_vol / (ptv_vol * iso_roi_vol)
                            plan_stat_val = round(plan_stat_val, 2)  # Round CI to 2 decimal places
                            bk_color = yellow if plan_stat_val < 1 else green if plan_stat_val <= 1.2 else yellow if plan_stat_val <= 1.5 else red
                            
                            intersect_roi.DeleteRoi()
                            iso_roi.DeleteRoi()
                        
                        elif stat == "GI (R50%)":
                            iso_roi = create_iso_roi(dose_dist, 0.5 * rx)  # 50% isodose
                            
                            plan_stat_val = struct_set.RoiGeometries[iso_roi.Name].GetRoiVolume() / ptv_vol  # Volume of 50% isodose geometry as proportion of PTV volume
                            plan_stat_val = round(plan_stat_val, 2)  # Round GI to 2 decimal places
                            
                            iso_roi.DeleteRoi()

                        elif stat == "D2cm [%]":
                            normal_tissue_name = exam_names[exam_name][0]
                            normal_tissue_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[normal_tissue_name].GetRoiVolume()  # Volume of everything except 2-cm expansion of PTV
                            rel_vol = 0.035 / normal_tissue_vol  # Appreciable volume (0.035 cc) as proportion of normal tissue volume
                            dose_at_rel_vol = dose_dist.GetDoseAtRelativeVolumes(RoiName=normal_tissue_name, RelativeVolumes=[rel_vol])[0]

                            plan_stat_val = dose_at_rel_vol / rx * 100  # Dose at relative volume, as percent of Rx
                            plan_stat_val = round(plan_stat_val, 1)  #Round D2cm to a single decimal place

                        elif stat == "Max dose @ appreciable volume [%]":
                            ext_name = exam_names[exam_name][1]
                            ext_vol = case.PatientModel.StructureSets[exam_name].RoiGeometries[ext_name].GetRoiVolume()  # Volume of external
                            rel_vol = 0.035 / ext_vol  # Appreciable volume (0.035 cc) as proportion of external volume
                            dose_at_rel_vol = dose_dist.GetDoseAtRelativeVolumes(RoiName=ext_name, RelativeVolumes=[rel_vol])[0]
                            
                            plan_stat_val = dose_at_rel_vol / rx * 100  # Dose at relative volume, as percent of Rx
                            plan_stat_val = round(plan_stat_val, 1)  # Round to a single decimal place
                            bk_color = yellow if plan_stat_val < 123 else green if plan_stat_val <= 130 else yellow if plan_stat_val <= 135 else red

                        elif stat == "V20Gy [%]":
                            lungs_ctv_name = exam_names[exam_name][2]
                            rel_vol = dose_dist.GetRelativeVolumeAtDoseValues(RoiName=lungs_ctv_name, DoseValues=[v20_vol])[0]  # Proportion of Lungs-CTV that receives 2000 cGy total dose
                            
                            plan_stat_val = rel_vol * 100  # Volume as a percent
                            plan_stat_val = round(plan_stat_val, 2)  # Round V20Gy to 2 decimal places

                        else:  # Max dose to External is inside PTV
                            ext_name = exam_names[exam_name][1]
                            ext_max = dose_dist.GetDoseStatistic(RoiName=ext_name, DoseType="Max")
                            ptv_max = dose_dist.GetDoseStatistic(RoiName=ptv_name, DoseType="Max")
                            
                            if ext_max == 0 or ptv_max == 0:  # External or PTV geometry is empty, or geometry has been updated since last voxel volume computation
                                plan_stat_val = "N/A"
                                bk_color = grey
                            elif ext_max == ptv_max:  # Same max dose, so assume it's the same point
                                plan_stat_val = "Yes"
                                bk_color = green
                            else:  # Different max doses, so assume they're at different points
                                plan_stat_val = "No"
                                bk_color = red

                        # If this stat is to be interpolated, add the None and Minor interpolated values to the row in interp cutoffs table
                        if stat in interp_stats:
                            none_dev = float(interpolate.interp1d(data["PTV vol [cc]"], data["{} None".format(stat)])(ptv_vol))
                            minor_dev = float(interpolate.interp1d(data["PTV vol [cc]"], data["{} Minor".format(stat)])(ptv_vol))
                            bk_color = green if plan_stat_val < none_dev else yellow if plan_stat_val < minor_dev else red
                            # Display no decimal places for V20Gy. Else, display 2 decimal places.
                            if stat == "V20Gy [%]":
                                none_dev, minor_dev = int(none_dev), int(minor_dev)
                            else:
                                none_dev, minor_dev = round(none_dev, 2), round(minor_dev, 2)
                            interp_cutoffs_row.extend([Paragraph("<{}".format(none_dev), style=tbl_data), Paragraph("<{}".format(minor_dev), style=tbl_data)])

                        # Add stat to plan stats row
                        plan_stats_row.append(plan_stat_val)
                        plan_stats_style.append(("BACKGROUND", (j + 2, plan_stats_idx), (j + 2, plan_stats_idx), bk_color))
                        plan_stats_style.append(("TEXTCOLOR", (j + 2, plan_stats_idx), (j + 2, plan_stats_idx), get_text_color(bk_color)))

                    # Add row to plan stats table
                    plan_stats_row = [Paragraph(str(text), style=tbl_data) for text in plan_stats_row]
                    plan_stats_data.append(plan_stats_row)

                    # Add row to interp cutoffs table
                    interp_cutoffs_data.append(interp_cutoffs_row)
                    interp_cutoffs_style.append(("BACKGROUND", (0, i + plan_stats_idx), (-1, i + plan_stats_idx), color))
                    interp_cutoffs_style.append(("TEXTCOLOR", (0, i + plan_stats_idx), (-1, i + plan_stats_idx), text_color))

        # Add non-plan row to interp cutoffs table
        interp_cutoffs_row = [round(row["PTV vol [cc]"], 1)]  # e.g., [1.8]
        for stat in interp_stats:
            for dev in ["None", "Minor"]:
                interp_cutoffs_val = row["{} {}".format(stat, dev)]  # e.g., "V20Gy [%] None"
                # If V20Gy, display without decimal place. Otherwise, display a single decimal place
                if stat == "V20Gy [%]":
                    interp_cutoffs_row.append(int(interp_cutoffs_val))
                else:
                    interp_cutoffs_row.append(round(interp_cutoffs_val, 1))
        interp_cutoffs_row = [Paragraph(str(text), style=tbl_data) for text in interp_cutoffs_row]
        interp_cutoffs_data.append(interp_cutoffs_row)

    # Finally create tables
    plan_stats_tbl = Table(plan_stats_data, style=TableStyle(plan_stats_style))
    interp_cutoffs_tbl = Table(interp_cutoffs_data, style=TableStyle(interp_cutoffs_style))

    elems = [KeepTogether([hdr, spcr_sm, mrn, spcr_sm, desc]), spcr_lg, KeepTogether([key_1, spcr_sm, key_2]), spcr_lg, KeepTogether([plan_stats_title, spcr_sm, plan_stats_tbl])]  # List of elements to build PDF from
    
    # Add interp cutoffs table to PDF only if interp stat(s) were selected
    # Would be more efficient to just not build this table, but...
    if interp_stats:
        elems.extend([spcr_lg, KeepTogether([interp_cutoffs_title, spcr_sm, interp_cutoffs_tbl])])

    # Build PDF report
    try:
        pdf.build(elems)
    except:
        MessageBox.Show("An SBRT lung analysis report is already open for this case. Close it and then rerun the script.")
        sys.exit(1)
