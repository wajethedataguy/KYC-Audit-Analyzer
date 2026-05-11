# utils.py

import os
import pandas as pd


# 📁 Extracts filename from full path
def get_filename(file_path):
    return os.path.basename(file_path)


# 📅 Safely parses date strings to datetime
def parse_date(date_str):
    try:
        return pd.to_datetime(date_str, errors="coerce")
    except Exception:
        return pd.NaT


# 🧠 Logic registry to store all registered logic functions
logic_registry = {}


# 🧩 Decorator to register logic functions with metadata
def register_logic(name, description, category):
    def decorator(func):
        logic_registry[name] = {
            "function": func,
            "description": description,
            "category": category,
        }
        return func

    return decorator
