from connect import *  # Interact w/ RS


def is_sbrt_exam(exam):
  # Helper function that returns True if exam description contains "AVG", "MIP", or "Gated", False otherwise.

  desc = exam.GetAcquisitionDataFromDicom()["SeriesModule"]["SeriesDescription"]
  if desc is None:
    return False
  return any(word in desc for word in ["AVG", "MIP", "Gated"])


def convert_virtual_jaw_to_mlc():
  """Convert virtual jaw fields from an external simulation (no MLCs) into jaw-and-MLC-defined fields

  Machine is set to "SBRT 6MV" if there are any AVG, MIP, or gated exams.
  Method is modified from SimPlanConvert_v4 from RS support
  """

  # Get current variables
  try:
    patient = get_current("Patient")
  except:
    MessageBox.Show("There is no patient loaded. Click OK to abort script.", "No Patient Loaded")
    sys.exit(1)
  try:
    case = get_current("Case")
  except:
    MessageBox.Show("There is no case loaded. Click OK to abort script.", "No Patient Loaded")
    sys.exit(1)
  try:
    plan = get_current("Plan")
  except:
    MessageBox.Show("There is no plan loaded. Click OK to abort script.", "No Plan Loaded")
    sys.exit(1)

  clinical_machine = "SBRT 6MV" if any(is_sbrt_exam(exam) for exam in case.Examinations) else "ELEKTA"

  # Delete any beam sets created by a previous run of this script
  # These beam set names end w/ a hyphen
  to_del = [bs.DicomPlanLabel for bs in plan.BeamSets if bs.DicomPlanLabel.endswith("-")]
  for name in to_del:
    plan.BeamSets[name].DeleteBeamSet()

  i = 0
  sim_beam_sets = [bs for bs in plan.BeamSets if bs.MachineReference.MachineName not in ["ELEKTA", "SBRT 6MV"]]  # If machine is correct, beam set is not a sim beam set
  while i < len(sim_beam_sets):
    sim_beam_set = sim_beam_sets[i]

    exam = sim_beam_set.GetPlanningExamination()
    position = sim_beam_set.PatientPosition

    # Create clinical beamset
    clin_beam_set = plan.AddNewBeamSet(Name=sim_beam_set.DicomPlanLabel[:(16 - i)] + "-" * (i + 1), ExaminationName=exam.Name,
      MachineName=clinical_machine, Modality="Photons",
      TreatmentTechnique="Conformal", PatientPosition=position,
      CreateSetupBeams=True, NumberOfFractions=1,
      UseLocalizationPointAsSetupIsocenter=True, Comment="")

    # Iterate over each beam in the simulation beamset 
    for b in sim_beam_set.Beams:
      iso_name = b.Isocenter.Annotation.Name
      iso_data = sim_beam_set.GetIsocenterData(Name=iso_name)
      iso_data["Name"] = iso_data["NameOfIsocenterToRef"] = "{} 1".format(iso_name)
      energy = 6 if b.MachineReference.Energy == 0 else b.MachineReference.Energy

      # Create the new beam
      new_b = clin_beam_set.CreatePhotonBeam(Energy=energy, Name=b.Name,
      GantryAngle=b.GantryAngle, CouchAngle=b.CouchAngle,
      CollimatorAngle=b.Segments[0].CollimatorAngle, IsocenterData=iso_data)

      # Create the field
      x1, x2, y1, y2 = b.Segments[0].JawPositions
      x_width = x2 - x1
      y_width = y2 - y1
      x_ctr = (x1 + x2) / 2
      y_ctr = (y1 + y2) / 2
      new_b.CreateRectangularField(Width=x_width, Height=y_width, CenterCoordinate={"x": x_ctr, "y": y_ctr},
      MoveMLC=True, MoveAllMLCLeaves=False, MoveJaw=True, JawMargins={"x": 0, "y": 0},
      DeleteWedge=False, PreventExtraLeafPairFromOpening=False)

    sim_beam_set.DeleteBeamSet()
    clin_beam_set.DicomPlanLabel = clin_beam_set.DicomPlanLabel[:(-i - 1)]
    i += 1

  plan.Comments = "Converted from sim"
  patient.Save()
  