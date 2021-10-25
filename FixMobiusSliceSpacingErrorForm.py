# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import os
import re
import shutil
import sys
sys.path.append(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\RayStation")

import pydicom  # Read/write/manipulate DICOM data
from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *

from CopyPlanToCTOrMergeBeamSetsForm import copy_plan_to_ct_or_merge_beam_sets
from PrepareExamsScript import prepare_exams


case = None  # Global so multiple functions can easily access it


class FixMobiusSliceSpacingErrorForm(Form):
    """Form that gets a non-negative floating point value from the user

    Show user an example of where to find the value to enter
    Submit button is disabled if text field value is invalid
    """

    def __init__(self):
        # Set up Form
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Fix Mobius Slice Spacing Error"
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)

        self.float = ""  # The value
        x = y = 15

        # Instructions for user
        instrs = Label()
        instrs.AutoSize = True
        instrs.AutoSizeMode = AutoSizeMode.GrowAndShrink
        instrs.Location = Point(x, y)
        instrs.Text = "Enter the value that caused the error. In the below screenshot, this value is 405.5."
        instrs_width = TextRenderer.MeasureText(instrs.Text, instrs.Font).Width
        instrs.MinimumSize = Size(instrs_width, 0)

        # Screenshot example of where to find the value in Mobius
        pb = PictureBox()
        pb.AutoSize = True
        pb.Image = Image.FromFile(r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Images\FixMobiusSliceSpacingErrorImage.png")
        y = instrs.Location.Y + instrs.Height
        pb.Location = Point(x, y)

        # Prompt for value
        val_lbl = Label()
        val_lbl.AutoSize = True
        val_lbl.AutoSizeMode = AutoSizeMode.GrowAndShrink
        y = pb.Location.Y + pb.Height + 80
        val_lbl.Location = Point(x, y)
        val_lbl_width = TextRenderer.MeasureText(val_lbl.Text, val_lbl.Font).Width
        val_lbl.Text = "Value:"

        # Input field for value
        self.tb = TextBox()
        self.tb.TextChanged += self.set_submit_enabled
        x = val_lbl.Location.X + val_lbl_width + 40
        self.tb.Location = Point(x, y)

        # Submit button
        self.submit_btn = Button()
        self.submit_btn.AutoSize = True
        self.submit_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.submit_btn.Click += self.submit_clicked
        self.submit_btn.Enabled = False
        x = val_lbl.Location.X
        y = self.tb.Location.Y + self.tb.Height + 15
        self.submit_btn.Location = Point(x, y)
        self.submit_btn.Text = "Submit"
        self.AcceptButton = self.submit_btn

        self.Controls.AddRange([instrs, pb, val_lbl, self.tb, self.submit_btn])
        self.ShowDialog()  # Display Form

    def set_submit_enabled(self, sender, event):
        try:
            self.float = float(self.tb.Text)
            self.submit_btn.Enabled = True
        except:
            self.submit_btn.Enabled = False

    def submit_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


class ChooseGatedGroupForm(Form):
    def __init__(self, grps):
        # Set up Form
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Fix Mobius Slice Spacing Error"
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        y = 15  # Vertical coordinate of next control

        l = Label()
        l.AutoSize = True
        l.Location = Point(15, y)
        l.Text = "Case contains multiple 4DCT groups with exams in the same FoR as the planning exam. Choose the group you want to use."
        self.Controls.Add(l)
        y += l.Height

        self.gb = GroupBox()
        self.gb.AutoSize = True
        self.gb.Location = Point(15, y)
        rb_y = 20
        for grp in grps:
            rb = RadioButton()
            rb.AutoSize = True
            rb.Checked = False
            rb.Click += self.set_ok_enabled
            rb.Location = Point(15, rb_y)
            rb.Text = grp.Name
            self.gb.Controls.Add(rb)
            rb_y += rb.Height
        self.Controls.Add(self.gb)
        y += self.gb.Height + 15

        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Enabled = False
        self.ok.Location = Point(15, y)
        self.ok.Text = "OK"
        self.Controls.Add(self.ok)

    def set_ok_enabled(self, sender, event):
        self.ok.Enabled = any(rb.Checked for rb in self.gb.Controls)

    def ok_clicked(self, sender, event):
        self.grp = [rb.Text for rb in self.gb.Controls if rb.Checked][0]
        self.DialogResult = DialogResult.OK


def compute_new_id(study_or_series):
    # Helper function that creates a unique DICOM StudyInstanceUID or SeriesInstanceUID for an examination
    # `study_or_series`: either "Study" or "Series"
    
    all_ids = [exam.GetAcquisitionDataFromDicom()["{}Module".format(study_or_series)]["{}InstanceUID".format(study_or_series)] for exam in case.Examinations]
    new_id = all_ids[-1]
    while True:
        if new_id not in all_ids:
            return new_id
        dot_idx = new_id.rfind(".")
        new_id = "{}{}".format(new_id[:(dot_idx + 1)], int(new_id[(dot_idx + 1):]) + 1)


def name_item(item, l, max_len=sys.maxsize):
    print(max_len)
    copy_num = 0
    old_item = item
    while item in l:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
        print(item)
    return item[:max_len]


def fix_mobius_slice_spacing_error():
    """Create a new plan that does not give a slice spacing error in Mobius

    This error occurs on some SBRT plans that are planned on an AVG exam:
    "CT: Image Position (Patient) (0020, 0032), CT slice spacing must be uniform; slice __/__: expected __, got __, delta > 0.1 mm"

    Assume planning exam is an average of gated images and that Rx is dose to volume

    Get problem slice location from a user GUI
    Export images, remove problematic slices, re-import, create new AVG, and copy plan (including Rx and Clinical Goals) onto new AVG
    The external copy function performs a dummy optimization to create control points. Optimization settings are not copied to the new plan
    The user must export the new plan to Mobius
    """

    global case

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
    patient_db = get_current("PatientDB")
    struct_set = plan.GetStructureSet()
    planning_exam = struct_set.OnExamination

    # Get offending slice location from user
    form = FixMobiusSliceSpacingErrorForm()
    if form.DialogResult != DialogResult.OK:
        sys.exit()  # User exited Form
    slice_loc = form.float

    # Create export folder
    folder = r"\\vs20filesvr01\groups\CANCER\Physics\Temp\FixMobiusSliceSpacingErrorScript"
    # Delete folder if it already exists
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.mkdir(folder)

    # Select gated images only
    gated_grps = [grp for grp in case.ExaminationGroups if grp.Type == "Collection4dct" and grp.Items[0].Examination.EquipmentInfo.FrameOfReference == planning_exam.EquipmentInfo.FrameOfReference]  # 4DCT group in same FoR as planning exam
    if not gated_grps:
        MessageBox.Show("There is no 4DCT gated group in the same FoR as the planning exam. Click OK to abort the script.")
        sys.exit(1)
    elif len(gated_grps) == 1:
        grp = grps[0]
    else:  # Found multiple gated groups, so user chooses the correct one from a GUI
        form = ChooseGatedGroupForm(gated_grps)
        if form.DialogResult != DialogResult.OK:
            sys.exit()
        grp = case.ExaminationGroups[form.grp]
    gated = [item.Examination.Name for item in grp]

    # Export gated images
    patient.Save()  # Must save changes before export
    try:
        case.ScriptableDicomExport(ExportFolderPath=folder, Examinations=gated, IgnorePreConditionWarnings=False)
    except SystemError as e:
        res = MessageBox.Show("The script generated the following warning(s): {}. Continue?".format(e), "Warnings", MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit(1)
        case.ScriptableDicomExport(ExportFolderPath=folder, Examinations=gated, IgnorePreConditionWarnings=True)

    # Compute new series IDs
    all_series_ids = [exam.Series[0].ImportedDicomUID for exam in case.Examinations]
    ids = {}
    for exam in gated:
        old_series_id = new_series_id = exam.Series[0].ImportedDicomUID
        dot_idx = new_series_id.rfind(".")
        first_part = new_series_id[:dot_idx]
        second_part = int(new_series_id[(dot_idx + 1):])
        while new_series_id in all_series_ids:
            second_part += 1
            new_series_id = "{}.{}".format(first_part, second_part)
        ids[old_series_id] = (new_series_id, exam.Name)
    
    # Delete all slices up to and including the offending slice
    # Change Series Instance UID and Series Instance UID so RayStation doesn't think these are the same images. 
    for f in os.listdir(folder):
        path = os.path.join(folder, f)  # Make absolute path to DICOM file
        dcm = pydicom.dcmread(path)
        slice_loc_dcm = float(dcm.ImagePositionPatient[2])
        if slice_loc_dcm >= slice_loc:  # Delete file if it is in the problematic region
            os.remove(path)
        else:
            dcm.SeriesInstanceUID = ids[dcm.SeriesInstanceUID]
            dcm.save_as(path)  # Overwrite original DICOM file
    
    # Import images w/o bad slices
    study_id = planning_exam.GetAcquisitionDataFromDicom()["StudyModule"]["StudyInstanceUID"]
    series = patient_db.QuerySeriesFromPath(Path=folder, SearchCriterias={"PatientID": patient.PatientID, "StudyInstanceUID": study_id})
    patient.ImportDataFromPath(Path=folder, SeriesOrInstances=series, CaseName=case.CaseName)

    # Rename imported exam
    avg_name = prepare_exams(study_id)

    # Copy plan to imported exam
    new_exam = case.Examinations[avg_name]
    copy_plan_to_ct_or_merge_beam_sets(replanName=plan.Name, replanCTName=avg_name)
    new_plan = list(case.TreatmentPlans)[-1]

    # Copy geometries
    geom_names = [geom.OfRoi.Name for geom in struct_set.RoiGeometries if geom.HasContours()]
    case.PatientModel.CopyRoiGeometries(SourceExamination=planning_exam, TargetExaminationNames=[new_exam.Name], RoiNames=geom_names)
    for poi_geom in struct_set.PoiGeometries:
        if all(abs(val) < 1000 for val in poi_geom.Point.values()):
            new_plan.GetStructureSet().PoiGeometries[poi_geom.OfPoi.Name].Point = poi_geom.Point

    # Crop ROI geometries to FOV
    fov_name = name_item("Field-Of-View", roi_names)
    fov = case.PatientModel.CreateRoi(Name=fov_name, Type="FieldOfView")
    fov.CreateFieldOfViewROI(ExaminationName=new_exam.Name)
    for geom_name in geom_names:  # Intersect FOV and each geometry on exam, if possible
        case.PatientModel.RegionsOfInterest[geom_name].CreateAlgebraGeometry(Examination=new_exam, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [geom_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [fov_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Intersection", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

    # Delete FOV ROI
    fov.DeleteRoi()

    # Fix dose grid
    new_plan.SetDefaultDoseGrid(VoxelSize=plan.GetDoseGrid().VoxelSize)

    # Recompute dose since structures were changed
    for i, bs in enumerate(new_plan.BeamSets):
        bs.ComputeDose(ComputeBeamDoses=True, DoseAlgorithm=plan.BeamSets[i].AccurateDoseAlgorithm.DoseAlgorithm)

    # Delete temp folder
    shutil.rmtree(folder)
