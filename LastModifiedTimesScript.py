import clr
clr.AddReference("System.Windows.Forms")
import os
import sys
from datetime import datetime

from connect import *
from System.Windows.Forms import *


def write_mod_info(lbl, obj, handle):
    mod_info = obj.ModificationInfo
    if mod_info is None:
        handle.write("{}: Modification info unavailable because updates need saving\n".format(lbl))
        return

    users = {"aak0": "Alisa", "aps0": "Algis", "b045z": "Dr. Jiang", "bq800": "Zach", "jh80g": "Courtney", "jm84u": "Kaley", "jvs1": "Jonas", "krt0": "King", "nj49e": "Beshoi", "yb25b": "Mary"}
    user = mod_info.UserName.split("\\")[-1]
    if user in users:
        user = users[user]

    dt = mod_info.ModificationTime
    date = "{}/{}/{}".format(dt.Month, dt.Day, dt.Year % 2000)

    if dt.Hour == 0:
        time = "12:{:02d} AM".format(dt.Minute)
    elif dt.Hour < 12:
        time = "{}:{:02d} AM".format(dt.Hour, dt.Minute)
    elif dt.Hour == 12:
        time = "12:{:02d} PM".format(dt.Minute)
    else:
        time = "{}:{:02d} PM".format(dt.Hour % 12, dt.Minute)
    
    handle.write("{}: {} {} by {}\n".format(lbl, date, time, user))


def last_modified_times():
    """Create and open a TXT file listing the last modified dates and times for objects that have this information
        - Patient
        - Registrations (for current case, if a case is loaded)
        - Structure Sets (for current case, if a case is loaded)
        - Beam sets (for current plan, if a plan is loaded)
        - Plan dose (for current plan, if a plan is loaded)
            - Beam set dose (for each beam set in current plan)
                - Beam dose (for each beam in each beam set)
        - Evaluation doses (for current case, if a case is loaded)
    """

    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)

    # Filename
    pt_name = ", ".join(patient.Name.split("^")[:2])
    timestamp = datetime.now().strftime("%m-%d-%y %H_%M_%S")
    filepath = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\LastModifiedTimes\{} Last Modified Times {}.txt".format(pt_name, timestamp)

    with open(filepath, "w") as f:
        write_mod_info("Patient", patient, f)
        f.write("\n")

        case = get_current("Case")
        if case is not None:
            for r in case.Registrations:
                r_name = "{} to {}".format(r.RegistrationSource.FromExamination.Name, r.RegistrationSource.ToExamination.Name)
                write_mod_info("Registration '{}'".format(r_name), r, f)
            if case.Registrations.Count > 0:
                f.write("\n")

            for exam in case.Examinations:
                write_mod_info("Structure set on '{}'".format(exam.Name), case.PatientModel.StructureSets[exam.Name], f)
            if case.Examinations.Count > 0:
                f.write("\n")

            plan = get_current("Plan")
            if plan is not None:
                for bs in plan.BeamSets:
                    write_mod_info("Beam set '{}'".format(bs.DicomPlanLabel), bs, f)
                if plan.BeamSets.Count > 0:
                    f.write("\n")
                
                write_mod_info("Plan dose", plan.TreatmentCourse.TotalDose, f)
                for bs in plan.BeamSets:
                    f.write("\t")
                    write_mod_info("Beam set dose '{}'".format(bs.DicomPlanLabel), bs.FractionDose, f)
                    for i, b in enumerate(bs.Beams):
                        f.write("\t\t")
                        write_mod_info("Beam dose '{}'".format(b.Name), bs.FractionDose.BeamDoses[i], f)
                    if bs.Beams.Count > 0:
                        f.write("\n")
                if plan.BeamSets.Count > 0:
                    f.write("\n")

                for fe in case.TreatmentDelivery.FractionEvaluations:
                    for doe in fe.DoseOnExaminations:
                        for de in doe.DoseEvaluations:
                            if de.PerturbedDoseProperties is not None:
                                # Perturbed dose
                                rds = de.PerturbedDoseProperties.RelativeDensityShift
                                density = "{0:.1f} %".format(rds * 100)
                                iso = de.PerturbedDoseProperties.IsoCenterShift
                                isocenter = "({0:.1f}, {1:.1f}, {2:.1f})".format(iso.x, iso.z, -iso.y)
                                beam_set_name = de.ForBeamSet.DicomPlanLabel
                                eval_dose_name = "Perturbed dose of {0}: {1}, {2}".format(beam_set_name, density, isocenter)
                            elif de.Name != "":
                                # Usually a summed dose
                                eval_dose_name = de.Name
                            elif hasattr(de, "ByStructureRegistration"):
                                # Mapped dose
                                reg_name = de.ByStructureRegistration.Name
                                beam_set_name = de.OfDoseDistribution.ForBeamSet.DicomPlanLabel
                                eval_dose_name = "Deformed dose of {} by registration {1}".format(beam_set_name, reg_name)
                            else:
                                # Neither perturbed, summed, nor mapped dose
                                eval_dose_name = de.ForBeamSet.DicomPlanLabel
                            write_mod_info("Eval dose '{}'".format(eval_dose_name), de, f)

    os.system(r'START /B C:\Windows\notepad.exe "{}"'.format(filepath))
