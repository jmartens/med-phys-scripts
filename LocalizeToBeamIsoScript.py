import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import sys

from connect import *

from System.Drawing import *
from System.Windows.Forms import *


class ChooseBeamForm(Form):
    # Helper class that allows user to choose which beam's isocenter to localize to

    def __init__(self, beam_set):
        self.Text = "Choose Beam"  # Form title
        self.AutoSize = True  
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot minimize, maximize, or resize form, but they can cancel ("X out of") it
        self.StartPosition = FormStartPosition.CenterScreen  # Position form in middle of screen
        y = 15  # Vertical coordinate of next control
        
        l = Label()
        l.AutoSize = True
        l.Location = Point(15, y)
        l.Text = "Not all beams in the current beam set have the same isocenter position. Choose the beam isocenter to localize to."
        self.Controls.Add(l)
        y += l.Height

        for beam in beam_set.Beams:
            rb = RadioButton()
            rb.AutoSize = True
            rb.Click += self.set_ok_enabled
            rb.Location = Point(15, y)
            rb.Text = beam.Name
            self.Controls.Add(rb)
            y += rb.Height

        self.ok = Button()
        self.ok.Click += self.ok_clicked
        self.ok.Enabled = False
        self.ok.Location = Point(self.ClientSize.Width - 50, y)
        self.ok.Text = "OK"
        self.AcceptButton = self.ok
        self.Controls.Add(self.ok)

        self.ShowDialog()  # Launch window

    def set_ok_enabled(self, sender, event):
        self.ok.Enabled = True

    def ok_clicked(self, sender, event):
        self.DialogResult = DialogResult.OK


def name_item(item, l, max_len=sys.maxsize):
    # Helper function that generates a unique name for `item` in list `l`
    # Limit name to `max_len` characters
    # E.g., name_item("Isocenter Name A", ["Isocenter Name A", "Isocenter Na (1)", "Isocenter N (10)"]) -> "Isocenter Na (2)"

    copy_num = 0
    old_item = item
    while item in l:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
    return item[:max_len]


def same_coords(point_1, point_2):
    # Helper function that returns whether the given points (dictionaries or ExpandoObjects) have the same coordinates

    return all(val == point_2[coord] for coord, val in point_1.items())


def tab_item_to_str(tab_item):
    # Helper function that returns the text of an RS UI TabItem

    return str(tab_item).split("'")[1]


def localize_to_beam_iso():
    """Localize views to the beam isocenter of the current beam set

    If beams in the current beam set have different iso coordinates, user chooses the beam to use, from a GUI
    After perfforming actions in Patient Modeling > Structure Definition, return user to the tab from which they ran the script
    """

    # Get current objects
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)
    try:
        beam_set = get_current("BeamSet")
        if beam_set.Beams.Count == 0:  # Beam set has no beams (and therefore no possible isos)
            MessageBox.Show("The current beam set has no beams. Click OK to abort the script.", "No Beams")
            sys.exit(1)
    except:
        MessageBox.Show("There is no beam set loaded. Click OK to abort the script.", "No Beam Set Loaded")
        sys.exit(1)

    # Get iso coordinates
    iso = [beam.Isocenter.Position for beam in beam_set.Beams]  # Iso coordinates of all beams
    if len(iso) == 1 or all(same_coords(iso[0], other_iso) for other_iso in iso[1:]):  # There is only one beam, or all beams have same iso coordinates
        iso = iso[0]
    else:  # Not all beams have same iso coordinates, so allow user to choose the beam whose iso to localize to
        form = ChooseBeamForm(beam_set)
        if form.DialogResult != DialogResult.OK:
            sys.exit()
        iso = [beam_set.Beams[ctrl.Text].Isocenter.Position for ctrl in form.Controls if isinstance(ctrl, RadioButton) and ctrl.Checked][0]  # The beam whose name is the checked radio button

    # Determine which module/submodule/tab the user is currently in
    modules = {"Fallback Planning": "Automated Planning", "Patient Information": "Patient Data Management", "Image Registration": "Patient Modeling", "Plan Setup": "Plan Design", "Plan Optimization": "Plan Optimization", "Plan Evaluation": "Plan Evaluation", "QA Preparation": "QA Preparation"}
    submodules = {"Automatic tools": "Image Registration", "ROI Tools": "Structure Definition", "Deformation": "Deformable Registration"}
    toolbar_tabs = {"EXTRAS": "ROI Tools", "CURRENT POI": "POI Tools", "STRUCTURE SET APPROVAL": "Approval", "LEVEL / WINDOW": "Fusion"}

    ui = get_current("ui")
    module = modules[tab_item_to_str(ui.TabControl_Modules.TabItem[0])]
    submodule = submodules[tab_item_to_str(ui.TabControl_ToolBar.TabItem[0])] if module == "Patient Modeling" else None
    toolbar_tab = toolbar_tabs[tab_item_to_str(list(ui.TabControl_ToolBar.ToolBarGroup)[-1])] if submodule == "Structure Definition" else None

    # Switch to Patient Modeling > Structure Definition > ROI Tools
    ui.TitleBar.MenuItem["Patient Modeling"].Click()
    ui.TitleBar.MenuItem["Patient Modeling"].Popup.MenuItem["Structure Definition"].Click()
    ui = get_current("ui")
    ui.TabControl_ToolBar.TabItem["ROI Tools"].Select()
    ui = get_current("ui")

    # Create new ROI at iso
    roi_name = name_item("Iso ROI", [roi.Name for roi in case.PatientModel.RegionsOfInterest])
    roi = case.PatientModel.CreateRoi(Name=roi_name, Type="Control")
    roi.CreateSphereGeometry(Radius=1, Examination=beam_set.GetPlanningExamination(), Center=iso)

    # Localize to ROI
    ui.TabControl_ToolBar.ToolBarGroup["CURRENT ROI"].Button_Localize_ROI.Click()

    # Return user to previous module/submodule
    if module == "Patient Modeling":
        ui.TabControl_Modules.TabItem[submodule].Select()
        if submodule == "Structure Definition":
            ui = get_current("ui")
            ui.TabControl_ToolBar.TabItem[toolbar_tab].Select()
    else:
        getattr(ui.TitleBar.MenuItem[module], "Button_{}".format("_".join(module.split()))).Click()  # E.g., click "Button_Plan_Design" button

    # Delete iso ROI
    roi.DeleteRoi()
