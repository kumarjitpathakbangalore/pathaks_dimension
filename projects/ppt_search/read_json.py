import json
import os

class JSONManager:
    """
    Manages reading from and writing to a JSON file.

    This class ensures that a JSON file exists in a 'text_output' directory,
    and handles its creation if it doesn't. It provides methods to read
    existing data and write new data to the file.
    """
    def __init__(self, filename="paraslides.json"):
        """
        Initializes the JSONManager with a specified filename.

        Args:
            filename (str): The name of the JSON file (default: "paraslides.json").
        """
        self.filename = filename
        # Construct the full path to the JSON file within the 'text_output' directory
        self.json_path = os.path.join("text_output", self.filename)

        # Ensure the 'text_output' directory exists
        os.makedirs("text_output", exist_ok=True)

        self.data = {}  # Initialize data as an empty dictionary

        # Check if the JSON file already exists
        if os.path.exists(self.json_path):
            try:
                # Attempt to read the file if it exists
                self._read_file()
            except json.JSONDecodeError as e:
                # Handle cases where the JSON file is malformed or empty
                print(f"Error reading JSON file '{self.filename}': {e}. Initializing with empty data.")
                self._write_file(self.data) # Overwrite with empty data to fix corrupt file
            except Exception as e:
                # Catch any other unexpected errors during file reading
                print(f"An unexpected error occurred while reading '{self.filename}': {e}. Initializing with empty data.")
                self._write_file(self.data) # Overwrite with empty data
        else:
            # If the file does not exist, create it with an empty JSON object
            print(f"File '{self.filename}' not found. Creating a new empty JSON file.")
            self._write_file(self.data)

    def _write_file(self, data):
        """
        Writes the given data to the JSON file.

        This is a private helper method, indicated by the leading underscore.

        Args:
            data (dict): The dictionary data to write to the JSON file.
        """
        try:
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            print(f"Successfully saved JSON file at {self.json_path}")
        except IOError as e:
            print(f"Error writing to JSON file '{self.filename}': {e}")
        except Exception as e:
            print(f"An unexpected error occurred while writing '{self.filename}': {e}")

    def _read_file(self):
        """
        Reads data from the JSON file and loads it into self.data.

        This is a private helper method, indicated by the leading underscore.
        Assumes the file exists and is valid JSON. Error handling for malformed
        JSON is done in the __init__ method.
        """
        with open(self.json_path, "r", encoding="utf-8") as file:
            self.data = json.load(file)
        print(f"Successfully loaded JSON data from {self.json_path}")

    def get_data(self):
        """
        Returns the currently loaded data from the JSON file.

        Returns:
            dict: The data loaded from the JSON file.
        """
        return self.data

    def update_data(self, new_data):
        """
        Updates the internal data and writes it back to the JSON file.

        Args:
            new_data (dict): The dictionary data to update and write.
        """
        self.data = new_data
        self._write_file(self.data)