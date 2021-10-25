# Network drive
z_path = r"\\vs19msqapp\MOSAIQ_APP\ESCAN"

# For GUI (MessagesBoxes display errors)
import clr
clr.AddReference("System.Windows.Forms")

import os
import shutil
import sys

from connect import *  # Interact w/ RS
from System.Windows.Forms import *  # For GUI


def create_qa_plan():
    """Create and export a DQA plan for the current beam set
    
    Export DICOM files to "Z:\TreatmentPlans\DQA"
    If export folder with the computed name already exists, delete and recreate it

    Detailed example of the code that names the new QA plan:
    Existing QA plans are "Test" and "Rectal Boost DQA"
    New QA plan for plan "Rectal Boost" would be "Rectal Boost DQA (1)", but name length must <=16 characters, so truncate the plan name to get QA plan name "Rectal B DQA (1)"
    If we create another QA plan for "Rectal Boost", it is named "Rectal B (2)"
    """

    # Get current objects
    try:
        beam_set = get_current("BeamSet") 
    except:
        MessageBox.Show("There is no beam set loaded. Click OK to abort the script.", "No Beam Set Loaded")
        sys.exit(1)
    patient = get_current("Patient")
    case = get_current("Case")
    plan = get_current("Plan")

    # Get unique DQA plan name while limiting to 16 characters
    qa_plan_name = "{} DQA".format(plan.Name[:12].strip())
    qa_plan_names = [qa_plan.BeamSet.DicomPlanLabel.lower() for qa_plan in plan.VerificationPlans]
    copy_num = 0
    while qa_plan_name.lower() in qa_plan_names:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        qa_plan_name = "{} DQA{}".format(plan.Name[:(12 - len(copy_num_str))].strip(), copy_num_str)

    # Create QA plan
    iso = {"x": 21.93, "y": 21.93, "z": 0}  # DICOM coordinates of "center" POI of Delta4 phantom
    dg = plan.GetDoseGrid().VoxelSize
    beam_set.CreateQAPlan(PhantomName="Delta4 Phantom", PhantomId="Delta4_2mm", QAPlanName=qa_plan_name, IsoCenter=iso, DoseGrid=dg, CouchAngle=0, ComputeDoseWhenPlanIsCreated=True)

    # Set colorwash to CRMC standard (adapted for smaller dose)
    case.CaseSettings.DoseColorMap.ColorTable = get_current("PatientDB").LoadTemplateColorMap(templateName="CRMC Standard Dose Colorwash").ColorMap.ColorTable

    # Ensure phantom electronics are not excessively irradiated
    qa_plan = list(plan.VerificationPlans)[-1]  # Latest QA plan
    point = {"x": 21.93, "y": 21.93, "z": 20}  # InterpolateDoseInPoint takes DICOM coordinates
    dose_at_electronics = qa_plan.BeamSet.FractionDose.InterpolateDoseInPoint(Point=point, PointFrameOfReference=qa_plan.BeamSet.FrameOfReference)
    if dose_at_electronics > 20:
        res = MessageBox.Show("Phantom electronics may be irradiated at {} > 20 cGy. Export anyway?".format(round(dose_at_electronics, 1)), "Create QA Plan", MessageBoxButtons.YesNo)
        # Exit script if user does not want to continue
        if res == DialogResult.No:
            sys.exit()

    # Create export folder
    patient.Save()  # Must save before any DICOM export
    pt_name = ", ".join(patient.Name.split("^")[:2])  # e.g., "Jones, Bill"
    qa_folder_name = "{} {}".format(pt_name, qa_plan_name)  # e.g., "Jones, Bill Prostate DQA"
    qa_folder_name = r"{}\TreatmentPlans\DQA\{}".format(z_path, qa_folder_name)  # Absolute path to export folder
    if os.path.isdir(qa_folder_name):
        shutil.rmtree(qa_folder_name)
    os.mkdir(qa_folder_name)
    
    # Export QA plan
    try:
        qa_plan.ScriptableQADicomExport(ExportFolderPath=qa_folder_name, QaPlanIdentity="Patient", ExportBeamSet=True, ExportBeamSetDose=True, ExportBeamSetBeamDose=True, IgnorePreConditionWarnings=False)
    except SystemError as e:
        res = MessageBox.Show("{}\nProceed?".format(e), "Create QA Plan", MessageBoxButtons.YesNo)
        if res == DialogResult.Yes:
            qa_plan.ScriptableQADicomExport(ExportFolderPath=qa_folder_name, QaPlanIdentity="Patient", ExportBeamSet=True, ExportBeamSetDose=True, ExportBeamSetBeamDose=True, IgnorePreConditionWarnings=True)
