# logic_loader.py
import importlib.util
import os
import sys
from .utils import logic_registry


def load_all_logics():
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    logic_path = os.path.join(base_path, "logic")

    for filename in os.listdir(logic_path):
        if filename.endswith(".py") and filename != "__init__.py":
            module_name = filename[:-3]
            file_path = os.path.join(logic_path, filename)
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                print(f"✅ Imported {module_name}")
            except Exception as e:
                print(f"❌ Failed to import {module_name}: {e}")


def execute_all_logics(dataframes, mode="lite"):
    results = {}
    for name, meta in logic_registry.items():
        func = meta["function"]
        try:
            df = func(dataframes, mode=mode)
            results[name] = df
            print(f"✅ Executed {name}")
        except Exception as e:
            print(f"❌ {name} failed: {e}")
    return results
