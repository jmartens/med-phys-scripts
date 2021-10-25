from connect import *


case = get_current('Case')
examination = get_current('Examination')
mat = {'M11': 1, 'M12': 0, 'M13': 0, 'M14': 1,
    'M21': 0, 'M22': 1, 'M23': 0, 'M24': 0,
    'M31': 0, 'M32': 0, 'M33': 1, 'M34': 0, 
    'M41': 0, 'M42': 0, 'M43': 0, 'M44': 1}
case.PatientModel.RegionsOfInterest[0].TransformROI3D(Examination=examination, TransformationMatrix=mat)


'Conform MLC (3, Beam Set: Anonymized)'