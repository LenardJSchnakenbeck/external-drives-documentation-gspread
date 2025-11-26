import json
import os


def load_json(file_path):
    if not os.path.exists(file_path):
        print("path does not exist: new file will be created")
        return {}
    else:
        with open(file_path, "r") as f:
            data = json.load(f)
            return data


def save_json(file_path, data):
    with open(file_path, "w") as f:
        json.dump(data, f, indent=2)
