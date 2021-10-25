t_path = r"\\vs20filesvr01\groups\CANCER\Physics"  # T: drive

import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))
from datetime import datetime
from os.path import getmtime

import pandas as pd
from connect import *
from System.Drawing import *
from System.Windows.Forms import *

from AddClinicalGoalsForm import add_clinical_goals, specific_rois


def update_clinical_goals_templates():
    """Update clinical goals templates in RayStation to match "Clinical Goals" spreadsheet
    
    Requires user interaction due to unscriptable actions delete and create clinical goals template
    If any ROIs in the template are not in the TG-263 spreadsheet, add them to "T:/Physics/Scripts/Output Files/UpdateClinicalGoalsTemplates/Not in TG-263 Spreadsheet.txt" and do not add goals for those ROIs
    
    Goals with a priority (1) need changing after the template is applied. These goals are one of two types:
        - Target goals. Dose values should be changed to that percent of the Rx. 
          E.g., "At least 90 cGy dose at 98.00% volume" with an Rx of 5000 cGy should be changed to "At least 4500 cGy dose at 98.00% volume
        - Volume-to-spare goals. These are implemented as absolute volume goals. The volume in the goal is the volume to spare, not the volume that belongs in the goal. Change the absolute volume to the ROI volume less the volume to spare.
          E.g., "At most 910 cGy dose at 700 cc" for an ROI with volume 1000 cc should be changed to "At most 910 cGy dose at 300 cc"

    """

    # Open templates file
    filename = "Clinical Goals.xlsx"
    rel_filepath = r"Scripts\Data"
    filepath = r"{}\{}\{}".format(t_path, rel_filepath, filename)
    try:
        goals = pd.read_excel(filepath, sheet_name=None, engine="openpyxl", usecols=["ROI", "Goal", "Notes"])  # Default xlrd engine does not support xlsx
    except FileNotFoundError:  # Alert user and abort script if spreadsheet does not exist
        msg = "Clinical goals spreadsheet '{}' does not exist at 'T:\{}'. Click OK to abort the script.".format(filename, rel_filepath)
        MessageBox.Show(msg, "No Clinical Goals Spreadsheet")
        sys.exit(1)

    output_filepath = r"{}\Scripts\Output Files\UpdateClinicalGoalsTemplates".format(t_path)
    last_run_filepath = r"{}\Last Run.txt".format(output_filepath)
    # Has clinical goals spreadsheet been modified since last script run?
    dt_fmt = "%Y-%m-%d %H:%M:%S.%f"
    try:
        with open(last_run_filepath, "r") as f: 
            script_time = datetime.strptime(f.read().strip(), dt_fmt)
            mod_time = datetime.fromtimestamp(getmtime(filepath))  # Spreadsheet modification date
            if script_time >= mod_time:
                msg = "The 'Clinical Goals' spreadsheet has not been modified since this script was last run, so the RS clinical goals templates will not be changed. This script takes a while, so are you sure you want to run it?"
                cont = MessageBox.Show(msg, "No Updates Necessary", MessageBoxButtons.YesNo)
                if cont == DialogResult.No:
                    sys.exit()
    except (FileNotFoundError, TypeError) as e:  # "Last Run" file does not exist, or value in file is not a valid datetime
        pass

    # Open dummy patient "Master Test"
    patient_db = get_current("PatientDB")
    pt = patient_db.QueryPatientInfo(Filter={"LastName": "Test", "FirstName": "Master", "PatientID": "123456789"})
    if pt == []:  # Test patient does not exist
        msg = "The patient Master Test (MR# 123456789) does not exist. Click OK to abort the script."
        MessageBox.Show(msg, "No Test Patient")
        sys.exit(1)
    pt = patient_db.LoadPatient(PatientInfo=pt[0])  # If test patient does exist, there will only be one
    pt.Cases["UpdateClinicalGoalsTemplates"].SetCurrent()
    case = get_current("Case")
    case.TreatmentPlans["Test"].SetCurrent()
    exam = get_current("Examination")

    # Remove "DNU" templates
    goals = {name: data for name, data in goals.items() if not name.endswith("DNU")}

    # Read in all TG-263 names
    filename = "TG263 Nomenclature with CRMC Colors.csv"
    filepath = r"{}\{}\{}".format(t_path, rel_filepath, filename)
    try:
        tg263 = pd.read_csv(filepath, usecols=["Target Type", "Major Category", "TG263-Primary Name", "Color"])
    except FileNotFoundError:  # Alert user and abort script if CSV file does not exist
        msg = "TG-263 names spreadsheet '{}' does not exist at 'T:\{}'. Click OK to abort the script.".format(filename, rel_filepath)
        MessageBox.Show(msg, "No TG-263 Names Spreadsheet")
        sys.exit(1)

    # Process each clinical goals template in the spreadsheet
    no_tg263 = []  # ROI names not in TG-263 spreadsheet

    # Size and center of ROI geometries to be created
    exam_min, exam_max = exam.Series[0].ImageStack.GetBoundingBox()
    exam_sz = {dim: exam_max[dim] - exam_min[dim] for dim in "xyz"}
    exam_ctr = {dim: (exam_min[dim] + exam_max[dim]) / 2 for dim in "xyz"}

    all_rois = ["CTV", "GTV", "PTV"]
    for data in goals.values():
        data["ROI"] = pd.Series(data["ROI"]).fillna(method="ffill")  # Autofill ROI name (due to vertically merged cells in spreadsheet)
        
        # Get all matching ROI names
        for roi in set(data["ROI"].values):
            if roi in specific_rois:
                rois = specific_rois[roi]
            else:
                rois = [roi]
            for r in rois:
                for suffix in "LR":
                    r_side = "{}_{}".format(r, suffix)#add_suffix(r, "_{}".format(suffix))
                    if r_side in tg263["TG263-Primary Name"].values:
                        rois.append(r_side)
            all_rois.extend(rois)

    # Create all ROIs
    #roi_geoms = case.PatientModel.StructureSets[exam.Name].RoiGeometries
    for roi in set(all_rois):
        if roi not in tg263["TG263-Primary Name"].values:
            no_tg263.append(roi)
            continue  # Skip this ROI

        row = tg263.loc[tg263["TG263-Primary Name"] == roi, ["Color", "Target Type", "Major Category"]].iloc[0, :]
        if row["Target Type"] == "Target":
            roi_type = row["Major Category"]
            color = "255,255,255,0" if roi_type == "CTV" else "255,255,165,0" if roi_type == "CTV" else "255,255,0,0"
            roi_type = roi_type[0].upper() + roi_type[1:].lower()
        else:
            if "PRV" in roi or "Ev" in roi:
                roi_type = "Control"
            else:
                roi_type = "Organ"
            color = ",".join(row["Color"].split("; "))
        if roi not in [r.Name for r in case.PatientModel.RegionsOfInterest]:
            r = case.PatientModel.CreateRoi(Name=roi, Color=color, Type=roi_type)
        else:
            r = case.PatientModel.RegionsOfInterest[roi]
            r.Color = color
        
        # Create geometry just in case clinical goal(s) require it (e.g., vol to spare)
        # Large geometry (same size as exam just in case vol to spare is large
        #if not roi_geoms[roi].HasContours():
        r.CreateBoxGeometry(Size=exam_sz, Examination=exam, Center=exam_ctr) 

    existing = [template["Name"] for template in patient_db.GetClinicalGoalTemplateInfo()]
    orig_existing = existing[:]
    for name in goals.keys():
        # Allow user to delete existing template, if it exists
        while name in existing:
            await_user_input("Please delete the template '{}' from RayStation.\nThen resume the script.".format(name))
            existing = [template["Name"] for template in patient_db.GetClinicalGoalTemplateInfo()]

        # Apply template
        add_clinical_goals(False, template_names=[name])

        # Allow user to save the template
        while name not in existing:
            await_user_input("Please save the goals as a template named '{}' in RayStation.\nClick OK in 'Save' dialog if it appears.\nThen resume the script.".format(name))
            existing = [template["Name"] for template in patient_db.GetClinicalGoalTemplateInfo()]

    # Delete any misnamed added templates
    necessary_templates = set(orig_existing + list(goals.keys()))
    extra = set(existing).difference(necessary_templates)
    while extra:
        await_user_input("Delete the following templates from RayStation: {}.\nThen resume the script.".format(", ".join(["'{}'".format(name) for name in extra])))
        existing = [template["Name"] for template in patient_db.GetClinicalGoalTemplateInfo()]
        extra = set(existing).difference(necessary_templates)

    # If there were some ROI(s) not in TG-263 spreadsheet, add them to the output file so Kaley can add them later
    if no_tg263:
        no_tg263_filepath = r"{}\not in TG-263 Spreadsheet.txt".format(output_filepath)
        with open(no_tg263_filepath, "w") as f:
            f.write("\n".join(sorted(list(set(no_tg263)))))

    # Write datetime to "Last Run" file
    with open(last_run_filepath, "w") as f:
        f.write(datetime.now().strftime(dt_fmt))
