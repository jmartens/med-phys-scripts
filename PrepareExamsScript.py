import datetime
import re
import sys

from connect import *  # Interact w/ RS


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


def prepare_exams(study_id=None):
    """Set imaging system, create 4DCT group (if necessary), and create AVG/MIP (if necessary)

    Parameters
    ----------
    study_id: str
        StudyInstanceUID of the study whose exams to prepare
        If None, use all exams with the latest date
        Defaults to None

    Rename exams:
        - Gated: "Gated __% <date>"
        - AVG: "AVG (Tx Planning) <date>"
        - MIP: MIP to "MIP <date>"
        - Non-gated: "3D <date>"
        - For other exams, if name does not include a plan or case name:
            * If there are plans on the exam, prepend name with name of first plan on the exam
            * If there are no plans on the exam, prepend name with case name
        - Any exam name may include a copy number to prevent duplicates
    
    4DCT group includes all gated examinations in the given study, that are not already part of a gated group
    Return the name of the average exam if it exists. Otherwise, return None
    """

    # Get current variables
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error

    # Call GetAcquisitionDataFromDicom ONCE for each exam, for speed
    dcm_data = {exam.Name: exam.GetAcquisitionDataFromDicom() for exam in case.Examinations}

    # We don't touch exams that have approved plans
    approved_planning_exam_names = list(set(plan.GetStructureSet().OnExamination.Name for plan in case.TreatmentPlans if plan.Review is not None and plan.Review.ApprovalStatus == "Approved"))

    # Select exams in the given study, or latest exams
    if study_id is not None:
        exams = [exam for exam in case.Examinations if exam.Name not in approved_planning_exam_names and dcm_data[exam.Name]["StudyModule"]["StudyInstanceUID"] == study_id]
    else:
        max_date = None
        exams = []
        for exam in case.Examinations:
            if exam.Name not in approved_planning_exam_names:
                date = dcm_data[exam.Name]["StudyModule"]["StudyDateTime"]
                if date is None:
                    if max_date is None:
                        exams.append(exam)
                else:
                    date = datetime.date(date.Year, date.Month, date.Day)
                    if max_date is None or date > max_date:
                        max_date = date
                        exams = [exam]

    # Prepare exams: fix imaging system name & add dates to exam names
    gated = []  # List of gated exams to include in new 4DCT group
    date = None
    avg = mip = non_gated = None
    for exam in exams:
        exam.EquipmentInfo.SetImagingSystemReference(ImagingSystemName="HOST-7307")  # Correct the imaging system

        dcm = dcm_data[exam.Name]
        desc = dcm["SeriesModule"]["SeriesDescription"] if dcm["SeriesModule"]["SeriesDescription"] is not None else dcm["StudyModule"]["StudyDescription"]  # Exam description is either series description (preferred) or study description
        date = dcm["StudyModule"]["StudyDateTime"].ToString("d")
        series_id = dcm["SeriesModule"]["SeriesInstanceUID"]
        other_exam_names = [exam_.Name for exam_ in case.Examinations if dcm_data[exam_.Name]["SeriesModule"]["SeriesInstanceUID"] != series_id]

        # Rename exam
        if "Non-Gated" in desc:
            non_gated = exam
            exam.Name = name_item("3D {}".format(date), other_exam_names)
        elif "AVG" in desc:
            avg = exam
            exam.Name = name_item("AVG (Tx Planning) {}".format(date), other_exam_names)
        elif "MIP" in desc:
            mip = exam
            exam.Name = name_item("MIP {}".format(date), other_exam_names)
        elif "Gated" in desc:
            pct = int(desc[(desc.index(",") + 2):-3])  # e.g., 10
            if pct == 0:
                exam.Name = name_item("Gated {}% (Max Inhale) {}".format(pct, date), other_exam_names)
            elif pct == 50:
                exam.Name = name_item("Gated {}% (Max Exhale) {}".format(pct, date), other_exam_names)
            else:
                exam.Name = name_item("Gated {}% {}".format(pct, date), other_exam_names)
            gated.append(exam.Name)  # Only include exam if it had to be renamed (is new)
        # Add plan or case name to exam name, if necessary
        else:
            names_to_chk = [plan_.Name for plan_ in case.TreatmentPlans if plan_.GetStructureSet().OnExamination.Name == exam.Name] + [case.CaseName]
            if not any(re.search(name, exam.Name) for name in names_to_chk):  # No plan or case name in exam name
                exam.Name = name_item("{} {}".format(names_to_chk[0], exam.Name), other_exam_names)
        # Add date to exam name, if necessary
        if not re.search("(\d{1,2}[/\-\. ]\d{1,2}[/\-\. ](\d{2}|\d{4}))|(\d{6}|\d{8})", exam.Name):  # E.g., 1/26/1999, 03/04/2020, 3/04/2020, 4/5/20, 7-8-21, 8-09-2021, 20200613, 210313
            exam.Name = name_item("{} {}".format(exam.Name, date), other_exam_names)

    # Create gated group from found gated exams
    if gated:
        grp_name = name_item("4D Phases {}".format(date), [grp.Name for grp in case.ExaminationGroups])
        case.CreateExaminationGroup(ExaminationGroupName=grp_name, ExaminationGroupType="Collection4dct", ExaminationNames=sorted(gated))  # Sort the exam names to ensure phases are in order
        if avg is None:  # Create AVG, if necessary
            avg_name = name_item("AVG (Tx Planning) {}".format(date), [exam_.Name for exam_ in case.Examinations])
            case.Create4DCTProjection(ExaminationName=avg_name, ExaminationGroupName=grp_name, ProjectionMethod="AverageIntensity")
            avg = case.Examinations[avg_name]
        if mip is None:  # Create MIP, if necessary
            mip_name = name_item("MIP {}".format(date), [exam_.Name for exam_ in case.Examinations])
            case.Create4DCTProjection(ExaminationName=mip_name, ExaminationGroupName=grp_name, ProjectionMethod="MaximumIntensity")

    # Deform from 3D to AVG
    if non_gated is not None and avg is not None:
        deform_name = name_item("{} to {}".format(non_gated.Name, avg.Name), [srg.Name for srg in case.PatientModel.StructureRegistrationGroups])
        case.PatientModel.CreateHybridDeformableRegistrationGroup(RegistrationGroupName=deform_name, ReferenceExaminationName=non_gated.Name, TargetExaminationNames=[avg.Name])
        poi_names = [poi.Name for poi in case.PatientModel.StructureSets[non_gated.Name].PoiGeometries]
        case.MapPoiGeometriesDeformably(PoiGeometryNames=poi_names, StructureRegistrationGroupNames=[deform_name], ReferenceExaminationNames=[non_gated.Name], TargetExaminationNames=[avg.Name])

    if avg is not None:
        return avg.Name
