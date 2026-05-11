# __init__.py in logics folder
import os
import glob
import importlib

# Automatically import all .py files in this folder except __init__.py
module_dir = os.path.dirname(__file__)
for filepath in glob.glob(os.path.join(module_dir, "*.py")):
    module_name = os.path.basename(filepath)[:-3]  # remove .py
    if module_name != "__init__":
        importlib.import_module(f"{__package__}.{module_name}")
