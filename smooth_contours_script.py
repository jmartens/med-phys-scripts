import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import sys
sys.path.append(r"{}\Scripts\RayStation".format(t_path))

from connect import *


def smooth_contours():
    