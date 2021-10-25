# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import re
import sys

from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


case = plan = None


def add_date_to_exam_name(exam):
    # Helper function that adds a date to the end of an exam name, if a date does not exist

    regex = "(\d{1,2}[\./\\_ ]\d{1,2}[\./\\_ ](\d{2}|\d{4}))|(\d{6}|\d{8})"
    if not re.search(regex, exam.Name):
        dcm = exam.GetAcquisitionDataFromDicom()
        if dcm["SeriesModule"]["SeriesDateTime"]:
            exam.Name = "{} {}".format(exam.Name, dcm["SeriesModule"]["SeriesDateTime"].ToString("d"))
        elif dcm["StudyModule"]["StudyDateTime"]:
            exam.Name = "{} {}".format(exam.Name, dcm["StudyModule"]["StudyDateTime"].ToString("d"))    


class QACTAdaptiveAnalysisForm(Form):
    """Windows Form that allows user to select parameters for an analysis of whether a replan is needed on an adaptive exam

    User chooses:
    - TPCT (reference exam)
    - QACT (floating)
    - Geometries to copy versus deform

    Support geometries are not allowed to be deformed
    """

    def __init__(self):
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "QACT Adaptive Analysis"
        self.MouseClick += self.mouse_click
        y = 15  # Vertical coordinate of next control

        # Make sure data is sufficient

        # External ROI exists
        ext_name = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "External"]
        if not ext_name:
            MessageBox.Show("There is no external ROI. Click OK to abort script.", "No External ROI")
            sys.exit(1)
        self.ext_name = ext_name[0]
        
        # Multiple examinations exist
        if case.Examinations.Count == 1:  # Only the TPCT exists
            MessageBox.Show("Case contains only one exam. Click OK to abort the script.", "No Possible QACTs")
            sys.exit(1)
        
        # External is contoured on TPCT
        ss = plan.GetStructureSet()
        self.tpct = ss.OnExamination  # TPCT is planning exam
        if not ss.RoiGeometries[self.ext_name].HasContours():  # No external on TPCT
            MessageBox.Show("There is no external geometry on the TPCT. Click OK to abort the script.", "No External Geometry")
            sys.exit(1)

        # Possible QACTs are not the TPCT, not registered to TPCT in opposite direction, and not in same frame of reference as TPCT
        ext_approved = self.ext_name in [struct.OfRoi.Name for ss_ in case.PatientModel.StructureSets for approved_set in ss_.ApprovedStructureSets for struct in approved_set.ApprovedRoiStructures]
        qacts = []
        for exam in case.Examinations:
            if exam.Name == self.tpct.Name:  # Exam is the TPCT
                continue
            if exam.EquipmentInfo.FrameOfReference == self.tpct.EquipmentInfo.FrameOfReference:  # Exam is in same FoR as TPCT
                continue
            if self.tpct.Name in [reg.RegistrationSource.ToExamination.Name for reg in case.Registrations if reg.RegistrationSource.FromExamination.Name == exam.Name]:  # Exam is registered to TPCT
                continue
            if ext_approved and not ss.RoiGeometries[self.ext_name].HasContours():
                continue
            qacts.append(exam)
        if not qacts:  # No other exams w/ an external
            criteria = ["It is not the TPCT (planning exam for current plan).", 
                        "It is not registered with the TPCT in the opposite direction (QACT -> TPCT).",
                        "It is not in the same frame of reference as the TPCT."]
            if ext_approved:
                criteria.append("It has external contours.")
            msg = "There are no possible QACTs. All of the following must be true of an examination used as a QACT:\n- {}\nClick OK to cancel the script.".format("\n- ".join(criteria))
            MessageBox.Show(msg, "No Possible QACTs")
            sys.exit(1)

        # TPCT name
        tpct_lbl = Label()
        tpct_lbl.AutoSize = True
        tpct_lbl.Font = Font(tpct_lbl.Font, FontStyle.Bold)
        tpct_lbl.Location = Point(15, y)
        tpct_lbl.Text = "TPCT:    "
        self.Controls.Add(tpct_lbl)
        tpct_name_lbl = Label()
        tpct_name_lbl.AutoSize = True
        tpct_name_lbl.Location = Point(15 + tpct_lbl.Width, y)
        tpct_name_lbl.Text = self.tpct.Name
        y += tpct_lbl.Height + 15
        self.Controls.Add(tpct_name_lbl)

        # QACT
        qact_lbl = Label()
        qact_lbl.AutoSize = True
        qact_lbl.Font = Font(qact_lbl.Font, FontStyle.Bold)
        qact_lbl.Location = Point(15, y)
        qact_lbl.Text = "QACT:    "
        self.Controls.Add(qact_lbl)
        if len(qacts) == 1:  # If only one possible QACT, display the name
            self.qact = qacts[0].Name
            qact_name_lbl = Label()
            qact_name_lbl.AutoSize = True
            qact_name_lbl.Location = Point(15 + qact_lbl.Width, y)
            qact_name_lbl.Text = self.qact
            self.Controls.Add(qact_name_lbl)
        else:  # If multiple possible QACTs, allow user to choose
            qacts = [qact.Name for qact in qacts]
            self.qact_combo = ComboBox()
            self.qact_combo.DropDownStyle = ComboBoxStyle.DropDownList
            self.qact_combo.Items.AddRange(qacts)
            self.qact_combo.Location = Point(15 + qact_lbl.Width, y)
            self.qact_combo.SelectedIndex = 0
            self.qact_combo.Width = max(TextRenderer.MeasureText(qact, self.qact_combo.Font).Width for qact in qacts) + 25
            self.Controls.Add(self.qact_combo)
        y += tpct_lbl.Height + 15

        geom_names = [geom.OfRoi.Name for geom in ss.RoiGeometries if geom.HasContours()]
        lbl_width = self.get_lbl_width(geom_names) * 2  # Double the neecessary width for better formatting
        
        # "Geometry" label
        geom_lbl = Label()
        geom_lbl.AutoSize = True
        geom_lbl.MinimumSize = Size(lbl_width, 0)
        geom_lbl.Font = Font(geom_lbl.Font, FontStyle.Bold)
        geom_lbl.Location = Point(15, y)
        geom_lbl.Text = "Geometry"
        self.Controls.Add(geom_lbl)

        # "Copy" label
        copy_lbl = Label()
        copy_lbl.AutoSize = True
        copy_lbl.Font = Font(copy_lbl.Font, FontStyle.Bold)
        copy_x = 15 + lbl_width
        copy_lbl.Location = Point(copy_x, y)
        copy_lbl.Text = "Copy"
        self.Controls.Add(copy_lbl)

        # "Deform" label
        deform_lbl = Label()
        deform_lbl.AutoSize = True
        deform_lbl.Font = Font(deform_lbl.Font, FontStyle.Bold)
        deform_x = copy_x + int(copy_lbl.Width * 1.5)
        deform_lbl.Location = Point(deform_x, y)
        deform_lbl.Text = "Deform"
        self.Controls.Add(deform_lbl)
        y += int(deform_lbl.Height * 1.25)  # Point() cannot take a float

        # "Select all" copy checkbox
        self.copy_all_cb = CheckBox()
        self.copy_all_cb.AutoSize = True
        self.copy_all_cb.IsThreeState = True
        self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        self.copy_all_cb.MinimumSize = Size(deform_x - copy_x, 0)  # At least as wide as deform label + space between deform and copy labels
        self.copy_all_cb.Click += self.copy_all_clicked
        self.copy_all_cb.Location = Point(copy_x, y)
        self.copy_all_cb.Text = "All"
        self.Controls.Add(self.copy_all_cb)

        # "Select all" deform checkbox
        self.deform_all_cb = CheckBox()
        self.deform_all_cb.AutoSize = True
        self.deform_all_cb.IsThreeState = True
        self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked
        self.deform_all_cb.MinimumSize = Size(deform_x - copy_x, 0)  # At least as wide as deform label + space between deform and copy labels
        self.deform_all_cb.Click += self.deform_all_clicked
        self.deform_all_cb.Location = Point(deform_x, y)
        self.deform_all_cb.Text = "All"
        self.Controls.Add(self.deform_all_cb)
        y += int(self.deform_all_cb.Height * 1.25)  # Point() cannot take a float

        # Zebra-striped geometry rows (alternate LightGray and the default Label background color)
        colors = [Color.LightGray, geom_lbl.BackColor]
        self.geom_cbs = {}  # ROI: [copy checkbox, deform checkbox]
        for i, geom_name in enumerate(geom_names):
            color = colors[i % 2]  # Alternate the two colors
            # Label
            geom_name_lbl = Label()
            geom_name_lbl.AutoSize = True
            geom_name_lbl.MinimumSize = Size(lbl_width, 0)
            geom_name_lbl.BackColor = color
            geom_name_lbl.Location = Point(15, y)
            geom_name_lbl.Text = geom_name
            self.Controls.Add(geom_name_lbl)

            # "Copy" checkbox
            copy_cb = CheckBox()
            copy_cb.AutoSize = True
            copy_cb.Checked = True
            copy_cb.MinimumSize = Size(deform_x - copy_x, 0)
            copy_cb.BackColor = color
            copy_cb.Click += self.copy_clicked
            copy_cb.Location = Point(copy_x, y)
            copy_cb.Name = "Copy{}".format(geom_name)  # e.g., "CopyPTV"
            self.Controls.Add(copy_cb)

            # "Deform" checkbox
            if ss.RoiGeometries[geom_name].OfRoi.Type == "Support":
                self.geom_cbs[geom_name] = [copy_cb, None]  # Cannot deform a support geometry
                placeholder = Label()
                placeholder.BackColor = color
                placeholder.Location = Point(deform_x, y)
                placeholder.Size = Size(deform_x - copy_x, copy_cb.Height)  # Placeholder is same size as a copy checkbox
                self.Controls.Add(placeholder)
            else:
                deform_cb = CheckBox()
                deform_cb.AutoSize = True
                deform_cb.Checked = False
                deform_cb.MinimumSize = Size(deform_x - copy_x, 0)
                deform_cb.BackColor = color
                deform_cb.Click += self.deform_clicked
                deform_cb.Location = Point(deform_x, y)
                deform_cb.Name = "Deform{}".format(geom_name)  # e.g., "DeformPTV"
                self.Controls.Add(deform_cb)
                self.geom_cbs[geom_name] = [copy_cb, deform_cb]

            y += int(deform_cb.Height * 1.25)  # Point() cannot take a float
        y += 15

        # Form is at least as wide as title text plus room for "X" button, etc.
        min_width = TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 150
        self.MinimumSize = Size(min_width, self.Height)

        # "Start" button
        self.start_btn = Button()
        self.start_btn.AutoSize = True
        self.start_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.start_btn.Click += self.start_clicked
        self.start_btn.Enabled = False
        self.start_btn.Text = "Start"
        b_x = self.ClientSize.Width - self.start_btn.Width  # Right align
        self.start_btn.Location = Point(b_x, y)
        self.Controls.Add(self.start_btn)
        self.AcceptButton = self.start_btn

        self.set_start_enabled()

        self.ShowDialog()  # Launch window
    
    def get_lbl_width(self, lbl_txt):
        # Return the widest label width necessary for any of the strings in list lbl_txt
        lbl_width = 0
        for txt in lbl_txt:
            l = Label()
            l.Text = txt
            lbl_width = max(lbl_width, l.Width)
        return lbl_width
    
    def mouse_click(self, sender, event):
        # Remove focus when user clicks the mouse
        self.ActiveControl = None
    
    def set_copy_checkstate(self):
        # When a copy checkbox is checked, set the appropriate "select all" checkstate
        copy_cbs_cked = [cbs[0].Checked for cbs in self.geom_cbs.values()]
        if all(copy_cbs_cked):
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        elif any(copy_cbs_cked):
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Indeterminate
        else:
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Unchecked

    def set_deform_checkstate(self):
        # When a deform checkbox is checked, set the appropriate "select all" checkstate
        deform_cbs = [cbs[1] for cbs in self.geom_cbs.values() if cbs[1] is not None]  # Select all deform checkboxes that exist
        deform_cbs_cked = [cb.Checked for cb in deform_cbs]
        if all(deform_cbs_cked):
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Checked
        elif any(deform_cbs_cked):
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Indeterminate
        else:
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked

    def copy_clicked(self, sender, event):
        # When a copy checkbox is checked, extract ROI name and uncheck that ROI's deform checkbox if it exists
        # Set checkstate of "select all" checkboxes
        geom = sender.Name[4:]  # ROI name is the part of the checkbox name after "Copy"
        deform_cb = self.geom_cbs[geom][1]
        if deform_cb is not None:
            deform_cb.Checked = False
        self.set_deform_checkstate()
        self.set_copy_checkstate()

    def deform_clicked(self, sender, event):
        # When a deform checkbox is checked, extract ROI name and uncheck that ROI's copy checkbox
        # Set checkstate of "select all" checkboxes
        geom = sender.Name[6:]  # ROI name is the part of the checkbox name after "Deform"
        copy_cb = self.geom_cbs[geom][0]
        copy_cb.Checked = False
        self.set_copy_checkstate()
        self.set_deform_checkstate()

    def copy_all_clicked(self, sender, event):
        # When "select all" copy checkbox is checked, set appropriate checkstate based on previous checkstate
        # For each ROI, uncheck the copy checkbox, or check the copy checkbox and uncheck the deform checkbox if it exists
        if self.copy_all_cb.Tag == CheckState.Checked:
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Unchecked
        else:
            self.copy_all_cb.CheckState = self.copy_all_cb.Tag = CheckState.Checked
        for cbs in self.geom_cbs.values():
            cbs[0].Checked = self.copy_all_cb.CheckState
            if self.copy_all_cb.CheckState and cbs[1] is not None and cbs[1].Checked:
                cbs[1].Checked = False
        self.set_deform_checkstate()
        self.set_start_enabled()

    def deform_all_clicked(self, sender, event):
        # When "select all" deform checkbox is checked, set appropriate checkstate based on previous checkstate
        # For each ROI, if the deform checkbox exists, uncheck it, or check it and uncheck the copy checkbox
        if self.deform_all_cb.Tag == CheckState.Checked:
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Unchecked
        else:
            self.deform_all_cb.CheckState = self.deform_all_cb.Tag = CheckState.Checked
        for cbs in self.geom_cbs.values():
            if cbs[1] is not None:
                cbs[1].Checked = self.deform_all_cb.CheckState
                if self.deform_all_cb.CheckState and cbs[0].Checked:
                    cbs[0].Checked = False
        self.set_copy_checkstate()
        self.set_start_enabled()

    def set_start_enabled(self):
        # Enable "Start" button if any checkboxes are checked
        copy_cbs_cked = [cbs[0].Checked for cbs in self.geom_cbs.values()]
        deform_cbs_cked = [cbs[1] for cbs in self.geom_cbs.values() if cbs[1] is not None and cbs[1].Checked]
        self.start_btn.Enabled = copy_cbs_cked or deform_cbs_cked

    def start_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


def qact_adaptive_analysis():
    """Perform an analysis on a TPCT and a QACT to determine whether an adaptive plan is needed

    Using a GUI, user chooses QACT (if there are multiple possible options) and geometries to copy or deform
    Note: Approved geometries on the QACT are not copied/deformed

    1. User chooses QACT (if there are options), ROIs to copy, and ROIs to deform, from a GUI.
    2. Rigidly register TPCT to QACT, if necessary. User can adjust registration if desired.
    3. If necessary, deform TPCT to QACT.
    4. Deform POI geometries from TPCT to QACT.
    5. Copy the specified ROI geometries from TPCT to QACT. Crop to QACT field=of-view.
    6. Deform the specified ROI geometries from TPCT to QACT.
    7. Resize QACT dose grid to include entire exam.
    8. If dose is computed on current plan, compute dose on the QACT for each of the plan's beam sets. Update dose grid structures for the new evaluation dose.
    9. Add "(TPCT)" and "(QACT)" to TPCT and QACT exam names, if necessary.

    Assumptions
    -----------
    No exam name contains a comma
    """

    global case, plan

    # Get current variables
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error
    try:
        plan = get_current("Plan")
    except:
        MessageBox.Show("There is no plan loaded. Click OK to abort script.", "No Plan Loaded")
        sys.exit(1)

    form = QACTAdaptiveAnalysisForm()
    if form.DialogResult != DialogResult.OK:  # User cancelled GUI window
        sys.exit()  

    # Extract attributes from Form
    tpct = form.tpct
    if hasattr(form, "qact_combo"):
        qact_name = form.qact_combo.SelectedItem
    else:
        qact_name = form.qact
    qact = case.Examinations[qact_name]
    if not case.PatientModel.StructureSets[qact.Name].RoiGeometries[form.ext_name].HasContours():
        case.PatientModel.RegionsOfInterest[form.ext_name].CreateExternalGeometry(Examination=qact)
    copy_geoms = [geom for geom, cbs in form.geom_cbs.items() if cbs[0].Checked]
    deform_geoms = [geom for geom, cbs in form.geom_cbs.items() if cbs[1] is not None and cbs[1].Checked]

    # Compute rigid registration, if necessary
    reg = [reg for reg in case.Registrations if reg.RegistrationSource.FromExamination.Name == tpct.Name and reg.RegistrationSource.ToExamination.Name == qact.Name]
    if reg:
        reg = reg[0]
    else:
        case.ComputeRigidImageRegistration(FloatingExaminationName=qact.Name, ReferenceExaminationName=tpct.Name, HighWeightOnBones=True)

        # Navigate to registration & allow user to make changes
        ui = get_current("ui")
        ui.TitleBar.MenuItem["Patient Modeling"].Button_Patient_Modeling.Click()
        ui = get_current("ui")  # New UI so that "Image Registration" tab is available
        ui.TabControl_Modules.TabItem["Image Registration"].Select()
        ui.ToolPanel.TabItem["Registrations"].Select()
        ui = get_current("ui")  # New UI so that list of registrations is available
        [tree_item for tree_item in ui.ToolPanel.RegistrationList.TreeItem if qact.Name in re.match(r"<.+'((.+(, )?)+)'>", str(tree_item)).group(1).split(", ")][0].Select()
        ui.ToolPanel.TabItem["Scripting"].Select()
        await_user_input("Review the rigid registration and make any necessary changes.")

    # Deformable registration
    grp_name = [srg.Name for srg in case.PatientModel.StructureRegistrationGroups for dsr in srg.DeformableStructureRegistrations if dsr.FromExamination.Name == tpct.Name and dsr.ToExamination.Name == qact.Name]  # Find deformable reg in structure reg group for TPCT -> QACT
    if grp_name:
        grp_name = grp_name[0]
    else:
        grp_name = "{} to {}".format(tpct.Name, qact.Name)
        case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=grp_name, ReferenceExaminationName=tpct.Name, TargetExaminationNames=[qact.Name])

    # Deform POI geometries
    #poi_names = [geom.OfPoi.Name for geom in case.PatientModel.StructureSets[tpct.Name].PoiGeometries if abs(geom.Point.x) < 1000]  # POI w/ no geometry has infinite coordinates
    #case.MapPoiGeometriesDeformably(PoiGeometryNames=poi_names, StructureRegistrationGroupNames=[grp_name], ReferenceExaminationNames=[tpct.Name], TargetExaminationNames=[qact.Name])

    # Copy ROI geometries
    if copy_geoms: 
        case.PatientModel.CopyRoiGeometries(SourceExamination=tpct, TargetExaminationNames=[qact.Name], RoiNames=copy_geoms)
        # Crop to FOV
        approved_fov_names = list(set([struct.OfRoi.Name for ss_ in case.PatientModel.StructureSets for approved_set in ss_.ApprovedStructureSets for struct in approved_set.ApprovedRoiStructures if struct.OfRoi.Type == "Field-Of-view"]))
        unapproved_fov_names = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type == "FieldOfView" and roi.Name not in approved_fov_names]
        if unapproved_fov_names:
            fov = case.PatientModel.RegionsOfInterest[unapproved_fov_names[0]]
        elif approved_fov_names:
            matches = [re.match("Field-Of-View( \((\d+)\))?", roi.Name, re.IGNORECASE) for roi in case.PatientModel.RegionsOfInterest if re.match("Field-Of-View( \((\d+)\))?", roi.Name, re.IGNORECASE)]
            if matches:
                match = sorted(matches, key=lambda match: 0 if not match.group(1) else int(match.group(2)))[-1]
                if match.group(1):
                    fov_name = "Field-Of-View ({})".format(match.group(2) + 1)
                else:
                    fov_name = "Field-Of-View (1)"
            fov = case.PatientModel.CreateRoi(Name=fov_name, Type="FieldOfView")
        else:
            fov = case.PatientModel.CreateRoi(Name="Field-Of-View", Type="FieldOfView")
        fov.CreateFieldOfViewROI(ExaminationName=qact.Name)
        for roi_name in copy_geoms:  # Intersect FOV and each geometry on exam, if possible
            case.PatientModel.RegionsOfInterest[roi_name].CreateAlgebraGeometry(Examination=qact, ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [roi_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [fov.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Intersection", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

    # Deform ROI geometries
    if deform_geoms:
        case.MapRoiGeometriesDeformably(RoiGeometryNames=deform_geoms, StructureRegistrationGroupNames=[grp_name], ReferenceExaminationNames=[tpct.Name], TargetExaminationNames=[qact.Name])
    
    # Resize dose grid
    dose_dist = plan.TreatmentCourse.TotalDose
    dg = plan.GetDoseGrid()
    minimum, maximum = tpct.Series[0].ImageStack.GetBoundingBox()
    for geom in case.PatientModel.StructureSets[qact.Name].RoiGeometries:
        if geom.HasContours():
            geom_min, geom_max = geom.GetBoundingBox()
            # If any infinite coordinates, skip this geometry (nothing will ever be over 1000...)
            if any([abs(coords[coord]) > 1000 for coords in [geom_min, geom_max] for coord in ["x", "y", "z"]]):
                continue
            minimum = {coord: min(minimum[coord], geom_min[coord]) for coord in minimum}
            maximum = {coord: max(maximum[coord], geom_max[coord]) for coord in maximum}
    num_voxels = {coord: (maximum[coord] - minimum[coord]) / dg.VoxelSize[coord] for coord in minimum}
    plan.UpdateDoseGrid(Corner=minimum, VoxelSize=dg.VoxelSize, NumberOfVoxels=num_voxels)
    dose_dist.UpdateDoseGridStructures()

    # Compute dose on QACT
    if dose_dist.DoseValues is None:
        MessageBox.Show("The dose distribution of the current plan has no dose, so dose will not be computed on the QACT.", "No Dose")
    else:
        # Create eval dose for each beam set
        for bs in plan.BeamSets:
            bs.ComputeDoseOnAdditionalSets(ExaminationNames=[qact.Name], FractionNumbers=[0])
        # Update dose grid structures for the new eval doses
        doe = [doe for doe in fe.DoseOnExaminations for fe in case.TreatmentDelivery.FractionEvaluations if doe.OnExamination.Name == qact.Name][0]  # All eval doses on QACT
        for i in range(plan.BeamSets.Count):
            list(doe.DoseEvaluations)[-i].UpdateDoseGridStructures()

    # Map dose (alternative to "compute dose")
    #deformable = case.Registrations[0].StructureRegistrations["Deformable Registration1"]
    #case.MapDose(DoseDistribution=dose_dist, StructureRegistration=deformable)

    # Rename TPCT and QACT if needed
    add_date_to_exam_name(tpct)
    add_date_to_exam_name(qact)
    if "TPCT" not in tpct.Name:
        tpct.Name += " (TPCT)"
    if "QACT" not in qact.Name:
        qact.Name += " (QACT)"

    # Delete unnecessary ROI
    fov.DeleteRoi()
