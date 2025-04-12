import json
import os

class ReadJSON:
    def __init__(self, file="paraslides.json"):
        self.file = file
        self.json_path = os.path.join("text_output", self.file)
        os.makedirs("text_output", exist_ok=True)

        if os.path.exists(self.json_path):
            try:
                self.read_file()
            except Exception as e:
                print(f"Error reading JSON: {e}")
                self.data = {}
                self.write_file(self.data)
        else:
            self.data = {}
            self.write_file(self.data)

    def write_file(self, data):
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        print(f"Saved JSON file at {self.json_path}")

    def read_file(self):
        with open(self.json_path, "r", encoding="utf-8") as file:
            self.data = json.load(file)
