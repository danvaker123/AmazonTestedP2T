import pandas  as pd
import json


def convert_excel_to_json(excel_file_path, json_file_path):
    """Convert Excel file to JSON format."""
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path)

        # Convert the DataFrame to JSON and save it to the specified file
        df.to_json(json_file_path, orient='records', lines=True)
        print(f"Successfully converted {excel_file_path} to {json_file_path}.")
    except Exception as e:
        print(f"Error converting Excel to JSON: {e}")

def load_user_input(json_file_path):
    """Load user input from a JSON file."""
    try:
        with open(json_file_path, 'r') as file:
            user_input = json.load(file)
        return user_input
    except Exception as e:
        print(f"Error loading user input from JSON: {e}")
        return []

def main():
    excel_file_path = ''  # Adjust path as necessary
    json_file_path = 'Input/input_data.json'    # Adjust path as necessary

    # Convert the Excel file to JSON
    convert_excel_to_json(excel_file_path, json_file_path)

    # Load user input from the JSON file
    user_input = load_user_input(json_file_path)

    # Example usage: Print user input data
    for entry in user_input:
        print(entry)

if __name__ == "__main__":
    main()
