t_path = r"\\vs20filesvr01\groups\CANCER\Physics"
z_path = r"\\vs19msqapp\MOSAIQ_APP\ESCAN"


import clr
clr.AddReference("System.Windows.Forms")
import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))
from os import system
from os.path import isfile
from re import split

from connect import *  # Interact w/ RS
from PyPDF2 import PdfFileMerger, PdfFileReader
from System.Windows.Forms import *

from SBRTLungAnalysisForm import sbrt_lung_analysis


def print_report():
    """Print a report for the current beam set, appending the SBRT Lung Analysis report if applicable.
    
    Report is saved to Z:\TreatmentPlans and opened for user to view
    If SBRT Lung Analysis report already exists, do not create a new one
    If an SBRT Lung Analysis report could not be created, display the error message from the SBRTLungAnalysis script and create the report without this file

    Assumptions
    -----------
    An SBRT lung case body site is "Thorax"
    """

    # Get current variables
    try:
        beam_set = get_current("BeamSet")
    except:
        MessageBox.Show("There are no beam sets in the current plan. Click OK to abort script.", "No Beam Sets")
        sys.exit(1)
    patient = get_current("Patient")
    plan = get_current("Plan")
    clinic_db = get_current("ClinicDB")

    # Exit with an error if beam set has no dose
    if beam_set.FractionDose.DoseValues is None:
        MessageBox.Show("Beam set has no dose. Click OK to abort script.", "No Dose")
        sys.exit(1)

    # Exit with an error if dose statistics need updating
    if any(geom.HasContours() and plan.TreatmentCourse.TotalDose.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution is None for geom in plan.GetStructureSet().RoiGeometries):
        MessageBox.Show("Dose statistics missing. Click OK to abort the script.", "Missing Dose Statistics")
        sys.exit(1)

    # Report template
    template = [t for t in clinic_db.GetSiteSettings().ReportTemplates if t.Name == "ReportTemplateV8B_050619"]
    if template:
        template = template[0]
    else:  # Display message if report template does not exist
        MessageBox.Show("The report template '{}' does not exist. Click OK to abort the script.".format(template.Name), "Template Does Not Exist")
        sys.exit(1)

    patient.Save()  # Must save patient before report creation

    # Report filename
    pt_name = ", ".join(split(r"\^+", patient.Name)[:2])
    filename = r"{}\TreatmentPlans\ {} {}.pdf".format(z_path, pt_name, plan.Name, beam_set.DicomPlanLabel)
    
    # Report
    try:
        beam_set.CreateReport(templateName=template.Name, filename=filename, ignoreWarnings=False)
    except Exception as e:
        res = MessageBox.Show("The script generated the following warnings:\n\n{}\nContinue?".format(str(e).split("at ScriptClient")[0]), "Warnings", MessageBoxButtons.YesNo)
        if res == DialogResult.No:
            sys.exit(1)
        beam_set.CreateReport(templateName=template.Name, filename=filename, ignoreWarnings=True)

    # Append SBRT Lung Analysis report if this is SBRT lung plan
    num_fx = beam_set.FractionationPattern.NumberOfFractions
    if num_fx == 5 and beam_set.Prescription.PrimaryDosePrescription.DoseValue >= 600 * num_fx and get_current("Case").BodySite == "Thorax":
        sbrt_report = True
        sbrt_filename = r"{}\Scripts\Output Files\SBRTLungAnalysis\{} SBRT Lung Analysis.pdf".format(t_path, pt_name)
        if not isfile(sbrt_filename):
            sbrt_report = False
            msg = sbrt_lung_analysis()
            if msg is None:
                sbrt_report = True
            else:
                MessageBox.Show("{}\nAn SBRT Lung Analysis report could not be created.".format(msg), "No SBRT Report")
        if sbrt_report:
            # Merge beam set report and SBRT report
            merger = PdfFileMerger()
            merger.append(PdfFileReader(open(filename, "rb")))
            merger.append(PdfFileReader(open(sbrt_filename, "rb")))
            merger.write(filename)
            merger.close()

    # Open report
    reader_paths = [r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe", r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"]  # Paths to Adobe Reader on RS servers
    for reader_path in reader_paths:
        try:
            system(r'START /B "{}" "{}"'.format(reader_path, filename))
            break
        except:
            continue
