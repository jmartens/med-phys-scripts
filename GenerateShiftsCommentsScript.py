from System.Windows.Forms import MessageBox
from connect import *
from os import system
from datetime import datetime
import sys
import clr
clr.AddReference("System.Windows.Forms")


def format_num(num, nearest_x):
    # Helper function that rounds the number to the nearest `nearest_x`
    # Return the string result with trailing zeroes and decimal point removed
    # `num`: The number to round
    # `nearest_x`: The interval to round to
    # E.g., format_num(1.61, 0.5) -> "1.5"

    num = round(num / nearest_x) * nearest_x
    # Don't count the "0." at the beginning
    num_places = len(str(nearest_x)) - 2
    num = format(num, ".{}f".format(num_places))
    # Remove trailing zeroes. If there was nothing but zeroes after the decinal point, remove it, too
    num = num.rstrip("0").rstrip(".")
    return num


def generate_shifts_comments():
    """Create and open a TXT file listing patient shifts (couch shifts, MOSAIQ Site Setup shifts) and CB/AP/PA setup SSD if applicable

    If there is no shift, write "No setup shifts" instead of the couch shift
    If there is no CBCT setup beam, use AP (for supine patients) or PA (for prone patients) setup beam.

    Assumptions
    -----------
    If a CB setup beam exists, only one CB setup beam exists.
    CB setup beam name or description contains "CB" (case insensitive).
    AP and PA setup beam name or description contain "AP" and "PA", respectively (case insensitive).
    """

    try:
        beam_set = get_current("BeamSet")  # Beam set
    except:
        MessageBox.Show(
            "No beam set is loaded. Click OK to abort the script.", "No Beam Set Loaded")
        sys.exit(1)
    if beam_set.Beams.Count == 0:
        MessageBox.Show(
            "There are no beams in the current beam set.", "No Beams")
        sys.exit(1)
    exam = beam_set.GetPlanningExamination()  # Planning exam
    case = get_current("Case")  # Case
    struct_set = case.PatientModel.StructureSets[exam.Name]  # Structure set

    # Localization geometry
    loc_pt = struct_set.LocalizationPoiGeometry
    # No loc point, or no loc point geometry on planning exam (infnite coordinates)
    if loc_pt is None or abs(loc_pt.Point.x) > 1000:
        MessageBox.Show(
            "There is no localization geometry on the planning exam. Click OK to abort the script.", "No Localization Geometry")
        sys.exit(1)
    loc_pt = loc_pt.Point  # The coordinates themselves

    # Beam set isocenter
    isos = [beam.Isocenter.Position for beam in beam_set.Beams]
    # Can't directly compare isocenter position b/c they are different objects: must compare coordinates themselves
    if len(set((iso.x, iso.y, iso.z) for iso in isos)) > 1:
        MessageBox.Show(
            "Beams in beam set have different isocenter coordinates. Click OK to abort script.", "Different Iso Coordinates")
        sys.exit(1)
    # All isos have same position, so doesn't matter which one we use
    iso = isos[0]

    # Calculate shifts
    x, y, z = (loc_pt[dim] - iso[dim]
               for dim in ["x", "y", "z"])  # Round to nearest tenth

    # "Left" or "Right"
    if x < 0:
        couch_r_l = "Right"
        msq_r_l = "Left"
    else:
        couch_r_l = "Left"
        msq_r_l = "Right"

    # "Superior" or "Inferior"
    if z < 0:
        couch_i_s = "Inferior"
        msq_i_s = "Superior"
    else:
        couch_i_s = "Superior"
        msq_i_s = "Inferior"

    # "Posterior" or "Anterior"
    if y < 0:
        couch_p_a = "Anterior"
        msq_p_a = "Posterior"
    else:
        couch_p_a = "Posterior"
        msq_p_a = "Anterior"

    # Round and format shifts
    x, y, z = (format_num(abs(shift), 0.1) for shift in (x, y, z))

    # Add couch shifts to comments
    comments = "1. Align to initial CT marks\n"
    if x == y == z == "0":
        comments += "No setup shifts"
    else:
        comments += "2. Shift couch so PATIENT is moved: {} {} cm (patient's right/left), {} {} cm, {} {} cm".format(
            couch_r_l, x, couch_i_s, z, couch_p_a, y)

    # CB setup SSD
    # "AP" for supine, "PA" for prone
    pos = "AP" if "Supine" in beam_set.PatientPosition else "PA"
    for string in ["CB", pos]:
        cb_ssds = [setup_beam.GetSSD() for setup_beam in beam_set.PatientSetup.SetupBeams if any(
            string in text.upper() for text in [setup_beam.Name, setup_beam.Description]) and abs(setup_beam.GetSSD() < 1000)]
        if cb_ssds:
            break
    if cb_ssds:
        # Assume only one CB setup beam, if any
        cb_ssd = format_num(cb_ssds[0], 0.25)
        comments += "\n{} setup SSD: {}".format(pos, cb_ssd)

    # Add MOSAIQ site setup shifts to comments
    #comments += "\n\nMOSAIQ Site Setup shifts:\n"
    #comments += "{} {} cm, {} {} cm, {} {} cm".format(msq_r_l, x, msq_i_s, z, msq_p_a, y)

    # Write comments to file
    dt = datetime.now().strftime("%m-%d-%y %H_%M_%S")
    filepath = r"\\vs20filesvr01\groups\CANCER\Physics\Scripts\Output Files\GenerateShiftsComments\{}.txt".format(
        dt)
    with open(filepath, "w") as f:
        f.write(comments)

    # Open comments file
    system(r'START /B notepad.exe "{}"'.format(filepath))
