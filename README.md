# Med Phys Scripts
These are scripts that I, the medical physics assistant at Cookeville Regional Medical Center, wrote for use in our radiation therapy treatment planning system. Please note that the scripts are **__works in progress__**. Most of them work, but there are a few bugs. You will need to make some changes to filepaths, etc. All scripts have been used with RayStation v8B only.

## Prerequisites
The scripts in this repository rely on a set off python packages to be available. They can easily be installed in your (virtual) environment using pip with the following command
```shell
pip install -r requirements.txt
```

## Usage
Create a new script in RayStation that calls the "main" function in the script you want to use. For example, create `AddClinicalGoals.py`:
```python
from AddClinicalGoalsForm import add_clinical goals
add_clinical_goals()
```

## Future Plans
My clinic is in the process of streamlining our development process. We currently do not have a "development environment." Once that is set up, the scripts in this repo, and others (some of which do not interact with RayStation) will be revamped into a Python package and published to PyPI.