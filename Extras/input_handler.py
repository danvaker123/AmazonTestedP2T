import pandas as pd

def functional_log_maker(input_file_path):
    file_path = input_file_path

    # Load Excel data to get task and configuration names
    df = pd.read_excel(file_path)
    task_names = df['Task Name'].to_list()
    config_names = df['Configuration Name'].to_list()

    # Define the log file path
    log_file_path = "../Logs/automation_log.txt"

    # Read the content of the log file
    with open(log_file_path, 'r') as file:
        content = file.readlines()

    # Create a list to store structured logs
    function_log = []
    logged_warnings = set()

    # Iterate over task and configuration pairs to process each section
    for idx, task in enumerate(task_names):
        section_header = f"\n--- Task: {task} | Configuration: {config_names[idx]} ---\n"
        function_log.append(section_header)
        errors = []
        warnings = []
        success = True

        for line in content:
            if "ERROR" in line and (f"Task '{task}'" in line or (config_names[idx]!= "None" and f"Configuration '{config_names[idx]}'" in line)):
                errors.append(line.strip())
                success = False
            elif "WARNING" in line and (f"Task '{task}'" in line or (config_names[idx]!= "None" and f"Configuration '{config_names[idx]}'" in line)):
                warnings.append(line.strip())

        if success and not warnings:
            function_log.append("    SUCCESS: Task executed successfully.\n")
        elif not success:
            function_log.append("    ERROR: Task encountered errors:\n")
            for error in errors:
                function_log.append(f"      - {error}\n")
        if warnings:
            function_log.append("    WARNING: Task encountered warnings:\n")
            for warning in warnings:
                function_log.append(f"      - {warning}\n")

    # Write the structured log to a new file
    output_file = "../Logs/functional_log.txt"
    with open(output_file, 'w') as file:
        for entry in function_log:
            file.write(entry)

functional_log_maker("../Input/input_data.xlsx")