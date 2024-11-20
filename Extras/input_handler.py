import pandas as pd

def functional_log_maker():
    file_path = "../Input/input_data.xlsx"

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
        # Get the configuration, and handle NaN or empty values gracefully
        config = config_names[idx] if idx < len(config_names) else ""
        if pd.isna(config) or config == "":
            config = "None"

        section_header = f"\n--- Task: {task} | Configuration: {config} ---\n"
        function_log.append(section_header)

        # Flags for tracking task state
        task_error_found = False
        task_warning_found = False
        task_success_found = False

        # Temporary storage for current task logs
        current_task_log = []

        # Iterate over the log content to process warnings, errors, and success messages
        for line in content:
            # Check if the line is relevant to the current task
            if f"Task '{task}'" in line or f"Subtask '{config}'" in line:
                current_task_log.append(line)

                # Capture warnings relevant to the task
                if "WARNING" in line and "Could not find element with locator" not in line:
                    if line.strip() not in logged_warnings:
                        function_log.append("    WARNING: " + line.strip() + "\n")
                        logged_warnings.add(line.strip())
                        task_warning_found = True

                # Capture any error for the task
                if "ERROR" in line:
                    function_log.append("    ERROR: " + line.strip() + "\n")
                    task_error_found = True  # Mark an error was found

                # Capture success message only if no prior errors for the task
                if f"Task '{task}' with Subtask '{config}' executed successfully" in line:
                    if not task_error_found:
                        function_log.append("    SUCCESS: " + line.strip() + "\n")
                        task_success_found = True

        # Finalize task status in the functional log
        if task_error_found:
            function_log.append(f"    FAILED: Task '{task}' did not execute successfully due to errors.\n")
        elif task_warning_found:
            function_log.append(f"    SUCCESS: Task '{task}' executed successfully with warnings.\n")
        elif task_success_found:
            function_log.append(f"    PASSED: Task '{task}' executed successfully with no warnings or errors.\n")
        else:
            function_log.append(f"    ERROR: Task '{task}' did not execute successfully.\n")

    # Write the structured log to a new file
    output_file = "../Logs/functional_log.txt"
    with open(output_file, 'w') as file:
        for entry in function_log:
            file.write(entry)

functional_log_maker()