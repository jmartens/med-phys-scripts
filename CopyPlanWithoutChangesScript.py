import sys

from connect import *  # Interact w/ RS


def name_item(item, l, max_len=sys.maxsize):
    copy_num = 0
    old_item = item
    while item in l:
        copy_num += 1
        copy_num_str = " ({})".format(copy_num)
        item = "{}{}".format(old_item[:(max_len - len(copy_num_str))].strip(), copy_num_str)
    return item[:max_len]


def copy_plan_without_changes(old_plan_name=None):
    """Copy the current plan, retaining beam set and isocenter names

    Return the name of the new plan
    
    The name of the copy is the name of the old plan with the copy number in parentheses
    The beam set and beam isocenter names match those of the old plan
    The beam numbers are unique across all cases in the current patient.
    The comments in the new plan are "Copy of ___"
    """

    # Get current variables
    try:
        patient = get_current("Patient")
    except:
        MessageBox.Show("There is no patient loaded. Click OK to abort the script.", "No Patient Loaded")
        sys.exit(1)  # Exit script with an error
    try:
        case = get_current("Case")
    except:
        MessageBox.Show("There is no case loaded. Click OK to abort the script.", "No Case Loaded")
        sys.exit(1)  # Exit script with an error
    if old_plan_name is None:
        try:
            old_plan = get_current("Plan")
        except:
            MessageBox.Show("There is no plan loaded. Click OK to abort the script.", "No Plan Loaded")
            sys.exit(1)
    else:
        old_plan = case.TreatmentPlans[old_plan_name]
    
    # Get iso names for each beam in each beam set in old plan
    iso_names = [[beam.Isocenter.Annotation.Name for beam in beam_set.Beams] for beam_set in old_plan.BeamSets]

    # Copy plan
    new_plan_name = name_item(old_plan.Name, [p.Name for p in case.TreatmentPlans], 16)  # Limit name length to 16 characters
    case.CopyPlan(PlanName=old_plan.Name, NewPlanName=new_plan_name)
    
    # Switch to copied plan
    patient.Save()  # Can't switch plans w/o saving
    case.TreatmentPlans[new_plan_name].SetCurrent()
    new_plan = get_current("Plan")

    new_plan.Comments = "Copy of {}".format(old_plan.Name)

    # Get old beam #'s
    beam_nums = []
    for c in patient.Cases:
        for p in c.TreatmentPlans:
            for bs in p.BeamSets:
                beam_nums.extend([b.Number for b in bs.Beams])
                beam_nums.extend([b.Number for b in bs.PatientSetup.SetupBeams])
    beam_num = max(beam_nums) + 1
    
    # Rename beam sets and isos
    with CompositeAction("Rename beam sets and isocenters in new plan"):
        for i, iso_names_beam in enumerate(iso_names):
            new_plan.BeamSets[i].DicomPlanLabel = old_plan.BeamSets[i].DicomPlanLabel
            for j, iso_name in enumerate(iso_names_beam):
                new_plan.BeamSets[i].Beams[j].Isocenter.EditIsocenter(Name=iso_name)
    
    # Renumber beams in new beam sets
    with CompositeAction("Renumber beams in new plan"):
        for bs in new_plan.BeamSets:
            for b in bs.Beams:
                b.Number = beam_num
                beam_num += 1
            for b in bs.PatientSetup.SetupBeams:
                b.Number = beam_num
                beam_num += 1

    return new_plan_name