
# For GUI
import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import os
import re

from connect import *  # Interact w/ RS

# For GUI
from System.Drawing import *
from System.Windows.Forms import *


case = None


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


class StructurePropagation4DCTForm(Form):
    """Propagate ROI geometries from a reference image set to all images in the selected 4DCT group.
    Create ITV on all images in the 4DCT group. ITV is union of ITV on gated images in 4DCT group, and target geometry on the reference image set.
    Display estimated min, mid, and max phases, as well as maximum excursions of the selected target ROI in the gated images.

    User selects 4DCT group, reference image set, ROIs to propagate, and structure by which to determine excursions, from a GUI

    External, fixation, and support structures are not options to be copied.
    """

    def __init__(self):
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # No minimize/maximize controls, and no form resizing
        self.ClientSize = Size(650, 500)
        self.Text = "Structure Propagation 4DCT"
        self.TopMost = True
        y = 0  # Vertical coordinate of next control

        # IMAGE SELECTION
        self.group_box()
        self.gb.Location = Point(0, y)
        self.gb.MinimumSize = Size(650, 0)
        self.gb.Text = "Image selection"

        # 4DCT group data
        grp_names = [group.Name for group in case.ExaminationGroups]
        # Exit script with an error if no 4DCT groups exist
        if not grp_names:
            MessageBox.Show("There are no 4DCT groups exist in the current case. Click OK to abort script.", "No 4DCT Groups")
            sys.exit(1)
        
        # 4DCT group Label
        lbl = Label()
        lbl.Location = Point(15, 15)
        lbl.Text = "4DCT group:"
        self.gb.Controls.Add(lbl)

        # 4DCT group ComboBox
        self.grp_names_cb = ComboBox()
        self.grp_names_cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.grp_names_cb.Location = Point(200, 15)
        self.grp_names_cb.Items.AddRange(grp_names)
        self.gb.Controls.Add(self.grp_names_cb)

        # Reference image Label
        lbl = Label()
        lbl.Location = Point(15, 40)
        lbl.Text = "Reference image:"
        self.gb.Controls.Add(lbl)

        # Reference image ComboBox
        exam_names = [exam.Name for exam in case.Examinations]
        self.ref_img_cb = ComboBox()
        self.ref_img_cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.ref_img_cb.Location = Point(200, 40)
        self.ref_img_cb.Items.AddRange(exam_names)
        self.gb.Controls.Add(self.ref_img_cb)
        y += self.gb.Height

        self.Controls.Add(self.gb)

        self.grp_names_cb.DropDownWidth = self.grp_names_cb.Width = self.ref_img_cb.DropDownWidth = self.ref_img_cb.Width = max(TextRenderer.MeasureText(text, self.grp_names_cb.Font).Width for text in grp_names + exam_names) + 15

        # STRUCTURES TO PROPAGATE
        self.structs = self.data_grid()
        self.gb.Location = Point(0, y)
        self.gb.Text = "Structures to propagate"

        # TARGET
        self.targets = self.data_grid()
        self.targets.MultiSelect = False
        self.gb.Location = Point(330, y)
        self.gb.Text = "Center-of-mass ROI"
        y += 250  # Leave room for data grids

        # Results label
        self.result = Label()
        self.result.Location = Point(0, y)
        self.result.AutoSize = True
        self.result.Visible = False
        self.Controls.Add(self.result)

        # Button
        self.run_btn = Button()
        self.run_btn.AutoSize = True
        self.run_btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
        self.run_btn.Location = Point(600, 460)
        self.run_btn.Text = "Run"
        self.Controls.Add(self.run_btn)

        # Event handlers
        self.ref_img_cb.SelectedIndexChanged += self.ref_img_chged
        self.grp_names_cb.SelectedIndex = self.ref_img_cb.SelectedIndex = 0  # Default: select first group and first ref image in lists
        self.run_btn.Click += self.run_clicked

        self.ShowDialog()  # Launch window

    def group_box(self):
        self.gb = GroupBox()
        self.gb.AutoSize = True
        self.gb.AutoSizeMode = AutoSizeMode.GrowAndShrink

    def data_grid(self):
        self.group_box()
        self.gb.MinimumSize = Size(320, 0)

        dg = DataGridView()
        dg.AllowUserToAddRows = dg.AllowUserToDeleteRows = dg.AllowUserToResizeRows = False  # User cannot change rows
        dg.AutoGenerateColumns = False
        dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dg.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        dg.BackgroundColor = Color.White
        dg.ClientSize = Size(self.gb.Size.Width, dg.ColumnHeadersHeight)
        # 2 columns: "Name" and "Type"
        dg.ColumnCount = 2
        dg.Columns[0].Name = "Name"
        dg.Columns[1].Name = "Type"
        dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dg.Location = Point(0, 15)
        dg.ReadOnly = True  # User cannot type in dose grid
        dg.RowHeadersVisible = False  # No row numbers
        dg.SelectionMode = DataGridViewSelectionMode.FullRowSelect  # User selects whole row, not individual cell
        self.gb.Controls.Add(dg)
        self.Controls.Add(self.gb)
        return dg

    """
    def grp_name_chged(self, sender, event):
        self.ref_img_cb.Items.Clear()
        grp_name = self.grp_names_cb.SelectedItem
        ref_imgs = [item.Examination.Name for item in case.ExaminationGroups[grp_name].Items]
        self.ref_img_cb.Items.AddRange(ref_imgs)
        self.ref_img_cb.SelectedIndex = 0
        
        text = list(self.grp_names_cb.Items) + list(self.ref_img_cb.Items)
        width = self.grp_names_cb.DropDownWidth = self.ref_img_cb.DropDownWidth = 10 + max([TextRenderer.MeasureText(t, self.grp_names_cb.Font).Width for t in text])
        self.grp_names_cb.ClientSize = Size(width, self.grp_names_cb.ClientSize.Height)
        self.ref_img_cb.ClientSize = Size(width, self.ref_img_cb.ClientSize.Height)
    """

    def ref_img_chged(self, sender, event):
        # Helper method that changes lists of possible structures and targets
        # Structures are all non-external structures with volume >= 0.001 cc on the selected reference exam
        # Targets are all targets with volume >= 0.001 cc on the selected reference exam

        # Clear existing structure and target lists
        self.structs.Rows.Clear()
        self.targets.Rows.Clear()

        for geom in case.PatientModel.StructureSets[self.ref_img_cb.SelectedItem].RoiGeometries:
            if geom.OfRoi.Type not in ["External", "Fixation", "Support"] and geom.HasContours() and geom.GetRoiVolume() >= 0.001:  # Ignore external or empty geometry
                self.structs.Rows.Add([geom.OfRoi.Name, geom.OfRoi.Type])
                if geom.OfRoi.OrganData.OrganType == "Target":
                    self.targets.Rows.Add([geom.OfRoi.Name, geom.OfRoi.Type])

        # Autosize DataGridView according to column and row heights
        for dg in [self.structs, self.targets]:
            ht = dg.ColumnHeadersHeight + sum(row.Height for row in dg.Rows)
            dg.ClientSize = Size(320, min(ht, 200))

    def set_run_enabled(self):
        # Helper method that enabled or disables the "Run" button when data selection is changed
        # Enable only if at least one structure is selected and a target is selected

        self.run_btn.Enabled = self.structs.SelectedRows.Count > 0 and self.targets.SelectedRows.Count > 0

    def run_clicked(self, sender, event):
        grp_name = self.grp_names_cb.SelectedItem
        ref_img = self.ref_img_cb.SelectedItem
        structs = [row.Cells[0].Value for row in self.structs.SelectedRows]
        target = self.targets.SelectedRows[0].Cells[0].Value

        if target not in structs:  # User did not select the target to propagate
            # If the target is not derived, we can just add it to the structures list
            if not case.PatientModel.RegionsOfInterest[target].DerivedRoiExpression:
                structs.append(target)
            # If the target is derived, all of its dependent ROIs must propagate as well
            else:
                structs.extend([geom.OfRoi.Name for geom in case.PatientModel.StructureSets[ref_img].RoiGeometries[target].GetDependentRois() if geom.OfRoi.Name not in structs])

        grp_4d = case.ExaminationGroups[grp_name]
        target_imgs = [item.Examination.Name for item in grp_4d.Items]  # All images in the group except the reference image

        # Get the same DIR algorithms settings as if DIR were run from UI.
        # if the internal structure of the lung is of specific interest (and you don't care about structures outside),
        # it could be considered to change DeformationStrategy to 'InternalLung'
        default_dir_settings = case.PatientModel.GetAlgorithmSettingsForHybridDIR(ReferenceExaminationName=ref_img, TargetExaminationName=target_imgs[0], FinalResolution={'x': 0.25, 'y': 0.25, 'z': 0.25}, DiscardImageInformation=False, UsesControllingROIs=False, DeformationStrategy='Default')
        
        # Create deformable registration
        dir_grp_name = name_item("DIR for ROI Propagation", [srg.Name for srg in case.PatientModel.StructureRegistrationGroups])
        case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=dir_grp_name, ReferenceExaminationName=ref_img, TargetExaminationNames=target_imgs, AlgorithmSettings=default_dir_settings)

        # Map ROI geometries from reference to all images in group
        # Ignore ROIs that already have contours on the given target image
        for struct in structs:
            # Deform geometry to all target images w/o geometry for this ROI
            imgs = [img for img in target_imgs if img != ref_img and not case.PatientModel.StructureSets[img].RoiGeometries[struct].HasContours()]  # Exam names that don't have this ROI contoured
            if imgs:
                case.MapRoiGeometriesDeformably(RoiGeometryNames=[struct], StructureRegistrationGroupNames=[dir_grp_name] * len(imgs), ReferenceExaminationNames=[ref_img] * len(imgs), TargetExaminationNames=imgs)  # Map the geometry from the refernce to the taregets
            
            # If ROI is a target, create ITV
            roi = case.PatientModel.RegionsOfInterest[struct]
            if roi.OrganData.OrganType == "Target":
                gated_imgs = [img for img in target_imgs if "Gated" in case.Examinations[img].GetAcquisitionDataFromDicom()["SeriesModule"]["SeriesDescription"]]  # Only create ITV from gated images
                non_gated_imgs = [img for img in target_imgs if img not in gated_imgs]
                
                # Create "real" and temporary ITV. Temp ITV is from geometries on all phases in group. We will set the geometry for the "real" ITV later by copying the temp ITV geometry into it and then underiving it. The ITV cannot depend on itself
                real_itv_name = "i{}".format(roi.Type.upper())
                real_itv_name = name_item(real_itv_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)  # Get unique name for "real" ITV
                real_itv = case.PatientModel.CreateRoi(Name=real_itv_name, Color=roi.Color, Type=roi.Type)
                
                itv_name = "{}^Temp".format(real_itv_name)
                itv_name = name_item(itv_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)  # Get unique name for ITV
                itv = case.PatientModel.CreateRoi(Name=itv_name, Color=roi.Color, Type=roi.Type)
                itv.CreateITV(SourceRegionOfInterest=roi, ExaminationNames=gated_imgs, MarginSettingsData={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })  # ITV from geometries on gated images
                
                # Copy ITV to each non-gated phase in group
                if non_gated_imgs:
                    case.PatientModel.CopyRoiGeometries(SourceExamination=case.Examinations[gated_imgs[0]], TargetExaminationNames=non_gated_imgs, RoiNames=[struct])  # Doesn't matter which gated exam we copy from b/c geometry is the same on all
                
                # Union each ITV geometry with ITV geometry on the reference exam

                # Create copy of target on reference exam
                copied_target_name = "Copy of {}".format(struct)
                copied_target_name = name_item(copied_target_name, [r.Name for r in case.PatientModel.RegionsOfInterest], 16)  # Get unique name for copied target ROI
                copied_target = case.PatientModel.CreateRoi(Name=copied_target_name)  # ROI that is copy of the target ROI
                
                # Copied target's geometry is same as target's
                copied_target.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [struct], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                copied_target.UpdateDerivedGeometry(Examination=case.Examinations[ref_img])

                # Copy the copied target to other phases in group
                case.PatientModel.CopyRoiGeometries(SourceExamination=case.Examinations[ref_img], TargetExaminationNames=[img for img in gated_imgs if img != ref_img], RoiNames=[copied_target_name])

                # Union the copied target with the ITV geometry on each phase
                for img in gated_imgs:
                    if img != ref_img:
                        real_itv.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [itv.Name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [copied_target_name], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Union", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
                        real_itv.UpdateDerivedGeometry(Examination=case.Examinations[img])

                # Underived "real" ITV since we're deleting the ROI it depends on
                real_itv.DeleteExpression()

                # Delete the temp ITV and the copied target
                itv.DeleteRoi()
                copied_target.DeleteRoi()

        # Update derived geometries
        for target_img in target_imgs:
            geoms = case.PatientModel.StructureSets[target_img].RoiGeometries
            # Add external geometry if necessary
            ext = [geom for geom in geoms if geom.OfRoi.Type == "External"]
            if ext:
                ext = ext[0]
            else:
                ext = case.PatientModel.CreateRoi(Name="External", Color="white", Type="External")
                ext.CreateExternalGeometry(Examination=qact)
            derived_rois = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.DerivedRoiExpression and all(geoms[dep_roi].HasContours() for dep_roi in geoms[roi.Name].GetDependentRois())]  # Derived ROIs with contours for all dependent ROIs
            if derived_rois:
                case.PatientModel.UpdateDerivedGeometries(RoiNames=derived_rois, Examination=case.Examinations[target_img])

        # Transverse coordinate of target center-of-mass in all target images
        ctrs_of_mass = [case.PatientModel.StructureSets[img].RoiGeometries[target].GetCenterOfRoi() for img in list(set([ref_img] + target_imgs))]
        ctrs_of_mass_x = [ctr.x for ctr in ctrs_of_mass]
        ctrs_of_mass_y = [ctr.y for ctr in ctrs_of_mass]
        ctrs_of_mass_z = [ctr.z for ctr in ctrs_of_mass]
        max_idx = ctrs_of_mass_z.index(max(ctrs_of_mass_z))
        min_idx = ctrs_of_mass_z.index(min(ctrs_of_mass_z))
        mid_idx = ctrs_of_mass_z.index(sorted(ctrs_of_mass_z)[len(ctrs_of_mass_z) / 2])
        text = "Phases:\n"
        text = "    Max: '{}'.\n".format(grp_4d.Items[max_idx].Examination.Name)
        text += "    Min: '{}'.\n".format(grp_4d.Items[min_idx].Examination.Name)
        text += "    Mid: '{}'.\n".format(grp_4d.Items[mid_idx].Examination.Name)
        text += "\nMax excursion:\n"
        text += "    R-L: {:.2f} cm\n".format(max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_x for ctr_2 in ctrs_of_mass_x))
        text += "    I-S: {:.2f} cm\n".format(max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_y for ctr_2 in ctrs_of_mass_y))
        text += "    P-A: {:.2f} cm\n".format(max(abs(ctr - ctr_2) for ctr in ctrs_of_mass_z for ctr_2 in ctrs_of_mass_z))
        self.result.Text = text
        self.result.Visible = True


def structure_propagation_4dct():
    global case

    # Get current variables
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Case Loaded")
        sys.exit(1)

    StructurePropagation4DCTForm()
