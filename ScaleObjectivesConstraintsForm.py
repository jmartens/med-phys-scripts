import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import sys

from connect import *
from System.Drawing import *
from System.Windows.Forms import *


plan = beam_set = objectives = constraints = None


class ScaleObjectivesConstraintsForm(Form):
    def __init__(self):
        global objectives, constraints

        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink  # Adapt form size to controls
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow  # User cannot resize form
        self.StartPosition = FormStartPosition.CenterScreen  # Start form in middle of screen
        self.Text = "Scale Objectives and Constraints"
        self.MinimumSize = Size(TextRenderer.MeasureText(self.Text, SystemFonts.CaptionFont).Width + 100, 0)
        self.y = 15  # Vertical coordinate of next control

        # OK button
        # Don't add yet because y-coordinate depends on what controls are added
        self.ok = Button()
        self.ok.Text = "OK"
        self.Controls.Add(self.ok)
        self.AcceptButton = self.ok

        # Controls depend on whether objectives and/or (scalable) constraints exist
        plan_opt = [opt for opt in plan.PlanOptimizations if opt.OptimizedBeamSets[0].DicomPlanLabel == beam_set.DicomPlanLabel][0]
        objectives = plan_opt.Objective
        constraints = plan_opt.Constraints
        # No objectives or constraints
        # -> Alert user and enable OK button for aborting script
        if objectives is None and constraints.Count == 0:
            l = self.add_lbl("There are no objectives or constraints for the current beam set.\n\nClick OK to abort the script.")
            self.y += l.Height + 10
            self.ok.Enabled = True
        # No objectives, and no constraints that would change by dose scaling (uniform dose constraint does not have dose as a parameter)
        # -> Alert user and enable OK button for aborting script
        elif objectives is None and all(hasattr(constraint.DoseFunctionParameters, "PercentStdDeviation") for constraint in constraints):
            l = self.add_lbl("Scaling to a new dose would not change any objectives or constraints.\n\nClick OK to abort the script.")
            self.y += l.Height + 10
            self.ok.Enabled = True
        # There are objectives and/or constraints
        else:
            self.ok.Enabled = False

            # Script description/instructions
            l = self.add_lbl("Scale objectives and constraints to a new dose value.")
            self.y += l.Height + 10

            # Align textboxes by using uniform label width
            self.txt_box_x = max(TextRenderer.MeasureText(txt, Font(l.Font, FontStyle.Bold)).Width for txt in ["Reference dose:", "Dose to scale to:", "Scale factor:", "Preview:"]) + 5

            # Reference dose
            # Default value is current beam set Rx, if it is present
            rx = beam_set.Prescription.PrimaryDosePrescription
            if rx is not None:
                rx = int(rx.DoseValue)  # Rx can't be fractional, so this is not truncating
            self.ref_dose_tb = self.add_dose_input("Reference dose:", rx)
            self.y += self.ref_dose_tb.Height

            # Dose to scale to
            # No default value
            self.scale_dose_tb = self.add_dose_input("Dose to scale to:")
            self.y += self.scale_dose_tb.Height

            # Scale factor
            # "?" until both dose textboxes have valid values
            self.add_lbl("Scale factor:", bold=True)
            self.scale_factor_lbl = self.add_lbl("?", x=self.txt_box_x)
            self.y += self.scale_factor_lbl.Height + 20

            ## Preview
            # An example of scaling an objective/constraint
            l = self.add_lbl("Preview:", bold=True)
            self.y += l.Height
            # If there are objective(s), use the first objective as the example
            if objectives is not None:
                dfp = objectives.ConstituentFunctions[0].DoseFunctionParameters
            # If no objectives, use the first scalable constraint as the example
            else:
                dfp = [constraint for constraint in constraints if not hasattr(constraint.DoseFunctionParameters, "PercentStdDeviation")][0].DoseFunctionParameters
            if not hasattr(dfp, "FunctionType"):  # Dose falloff
                string = "Dose Fall-Off [H]__ cGy [L]__ cGy, Low dose distance {:.2f} cm".format(dfp.LowDoseDistance)
                self.doses = [int(dfp.HighDoseLevel), int(dfp.LowDoseLevel)]  # Current dose values in the example objective/constraint
            else:
                self.doses = [int(dfp.DoseLevel)]  # Current dose values in the example objective/constraint
                if dfp.FunctionType == "MinDose":
                    string = "Min dose __ cGy"
                elif dfp.FunctionType == "MaxDose":
                    string = "Max dose __ cGy"
                elif dfp.FunctionType == "MinDvh":
                    string = "Min DVH __ cGy to {}% volume".format(dfp.PercentVolume)
                elif dfp.FunctionType == "MaxDvh":
                    string = "Max DVH __ cGy to {}% volume".format(dfp.PercentVolume)
                elif dfp.FunctionType == "UniformDose":
                    string = "Uniform dose __ cGy"
                elif dfp.FunctionType == "MinEud":
                    string = "Min EUD __ cGy, Parameter A {:.0f}".format(dfp.EudParameterA)
                elif dfp.FunctionType == "MaxEud":
                    string = "Max EUD __ cGy, Parameter A {:.0f}".format(dfp.EudParameterA)
                else:  # UniformEud
                    string = "Target EUD __ cGy, Parameter A {:.0f}".format(dfp.EudParameterA)
            self.add_preview_objective(string, self.doses)  # Labels containing current values for the example objective/constraint
            l = self.add_lbl("will be changed to", x=70)
            self.y += l.Height
            self.dose_lbls = self.add_preview_objective(string, ["?"] * len(self.doses))  # Values for scaled example objective/constraint are "?" until a scale factor is computed
        
        self.ok.Location = Point(self.ClientSize.Width - 50, self.y)  # Right align
        self.ok.Click += self.ok_clicked

        self.ShowDialog()  # Launch window
        
    def add_lbl(self, lbl_txt, **kwargs):
        # Helper method that adds a label to the Form
        # kwargs:
        # x: x-coordinate of new label
        # bold: True if label text should be bold, False otherwise

        x = kwargs.get("x", 15)
        bold = kwargs.get("bold", False)

        l = Label()
        l.AutoSize = True
        if bold:
            l.Font = Font(l.Font, FontStyle.Bold)
        l.Location = Point(x, self.y)
        l.Text = lbl_txt
        self.Controls.Add(l)
        return l

    def add_dose_input(self, lbl_txt, dose=None):
        # Helper method that adds a set of controls for getting dose input from the user
        # Label with description of dose field, text box, and "cGy" label
        # lbl_txt: Description of field
        # dose: Starting value for text box. If None, text box starts out empty.

        # Label w/ description of field
        l = self.add_lbl(lbl_txt, bold=True)

        # TextBox
        tb = TextBox()
        tb.Location = Point(self.txt_box_x, self.y - 5)  # Approximately vertically centered with description label
        if dose is not None:
            tb.Text = str(dose)
        tb.TextChanged += self.dose_chged
        tb.Size = Size(40, l.Height)  # Wide enough for 6 digits
        self.Controls.Add(tb)

        # "cGy" label
        self.add_lbl("cGy", x=self.txt_box_x + 45)
        self.y += l.Height

        return tb

    def add_preview_objective(self, string, doses):
        # Helper method that adds an example objective/constraint with the given dose values
        # Dose values are bold
        # string: Text to display, with "__" standing in for each dose value
        # doses: List of dose values to substitute for "__" in `string`

        lbl_txts = string.split("__")  # E.g., "Dose Fall-Off [H]__ cGy [L]__ cGy, Low dose distance 1.00 cm" -> ["Dose Fall-Off [H]", " cGy [L]", " cGy, Low dose distance 1.00 cm"]
        dose_lbls = []
        x = 50  # Offset from left
        for i, txt in enumerate(lbl_txts):
            l = self.add_lbl(txt, x=x)
            self.Controls.Add(l)
            x += l.Width
            # There is dose to display
            if i != len(lbl_txts) - 1:
                l_2 = self.add_lbl(str(doses[i]), bold=True, x=x)
                l_2.Width = 45  # Wide enough for 6 digits
                self.Controls.Add(l_2)
                x += 45
                dose_lbls.append(l_2)
        self.y += l.Height
        return dose_lbls

    def dose_chged(self, sender, event):
        # Event handler for changing the text of a dose input field
        # Turn background red if input is invalid
        # If both dose values valid, compute and display scale factor, change preview output, and enable OK button
        # If either dose value invalid, display "?" for scale factor and preview output, and disable OK button

        # Reference dose
        self.ref_dose = self.ref_dose_tb.Text
        # If empty, don't turn red, but the value is still not valid
        if self.ref_dose == "":
            self.ref_dose_tb.BackColor = Color.White
            ref_dose_ok = False
        else:
            try:
                self.ref_dose = float(self.ref_dose)
                if self.ref_dose != int(self.ref_dose) or self.ref_dose <= 0 or self.ref_dose > 100000:  # Number is fractional or out of range
                    raise ValueError
                self.ref_dose_tb.BackColor = Color.White
                ref_dose_ok = True
            except ValueError:
                self.ref_dose_tb.BackColor = Color.Red
                ref_dose_ok = False

        # Dose to scale to
        self.scale_dose = self.scale_dose_tb.Text
        # If empty, don't turn red, but the value is still not valid
        if self.scale_dose == "":
            self.scale_dose_tb.BackColor = Color.White
            scale_dose_ok = False
        else:
            try:
                self.scale_dose = float(self.scale_dose)
                if self.scale_dose != int(self.scale_dose) or self.scale_dose < 0 or self.scale_dose > 100000:  # Number is fractional or out of range
                    raise ValueError
                self.scale_dose_tb.BackColor = Color.White
                scale_dose_ok = True
            except ValueError:
                self.scale_dose_tb.BackColor = Color.Red
                scale_dose_ok = False
        
        # Both dose values are valid
        if ref_dose_ok and scale_dose_ok:
            self.scale_factor = self.scale_dose / self.ref_dose
            
            # Display at most 3 decimal places, with no trailing zeros
            display = str(round(self.scale_factor, 3)).rstrip("0")
            if display.endswith("."):
                display = display[:-1]
            self.scale_factor_lbl.Text = display
            
            # Display rounded computed doses in preview
            for i, dose_lbl in enumerate(self.dose_lbls):
                dose_lbl.Text = "{:.0f}".format(self.scale_factor * self.doses[i])
            self.ok.Enabled = True
        # Invalid dose value(s)
        else:
            self.scale_factor_lbl.Text = "?"  # Can't compute scale factor
            for i, dose_lbl in enumerate(self.dose_lbls):  # Can't compute preview dose value(s)
                dose_lbl.Text = "?"
            self.ok_enabled = False

    def ok_clicked(self, sender, event):
        # Event handler for clicking the OK button

        if hasattr(self, "ref_dose"):  # self.ref_dose is created in dose_chged event handler, so only exists if dose text boxes exist
            self.DialogResult = DialogResult.OK
        else:  # There are no objectives and constraints, and a label on the form tells the user to click "OK" to abort the script
            self.DialogResult = DialogResult.Abort


def scale_objectives_constraints():
    """Scale the current beam set's objectives and constraints to a new dose value

    (If there is only one beam set in the current plan, these objectives and constraints apply to the plan as a whole.)
    Change objectives and onstraints in place: do not add objectives and constraints.
    User supplies reference dose (default is Rx dose, if it is provided) and dose to scale to, in a GUI.
    All doses in objectives and constraints are multiplied by a scale factor = dose to scale to / reference dose

    Assumptions
    -----------
    Each PlanOptimization applies to only one beam set
    """

    global plan, beam_set

    try:
        beam_set = get_current("BeamSet")
    except:
        MessageBox.Show("There is no beam set loaded.", "No Beam Set Loaded")
        sys.exit(1)
    plan = get_current("Plan")
    # If plan is approved, cannot change objectives and constraints
    if plan.Review is not None and plan.Review.ApprovalStatus == "Approved":
        MessageBox.Show("Plan is approved, so objectives and constraints canot be changed. Click OK to abort the script.", "Plan Is Approved")
        sys.exit(1)

    # Get scale factor from user input in a GUI
    form = ScaleObjectivesConstraintsForm()
    if form.DialogResult != DialogResult.OK:
        sys.exit()
    scale_factor = form.scale_factor
    
    # List of objectives and/or constraints
    objectives_constraints = list(constraints)
    if objectives is not None:
        objectives_constraints.extend(list(objectives.ConstituentFunctions))

    # Scale each objective or constraint
    for o_c in objectives_constraints:
        dfp = o_c.DoseFunctionParameters
        if not hasattr(dfp, "PercentStdDeviation"):  # The objective or constraint is scalable
            for attr in dir(dfp):
                # Multiply each dose parameter by the scale factor
                if "DoseLevel" in attr:
                    val = getattr(dfp, attr)
                    setattr(dfp, attr, val * scale_factor)
