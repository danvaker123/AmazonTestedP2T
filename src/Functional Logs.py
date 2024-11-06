import pandas as pd




def functional_log_maker():
    file_path = r"Input\input_data.xlsx"

    df = pd.read_excel(file_path)
    task_name = df['Task Name'].to_list()

    log_file_path = r'src\automation_log.txt'

    with open(log_file_path, 'r') as file:
        content = file.readlines()

    function_log = []
    function_log.append("Logs for Task Name : " + task_name[0] + "\n")

    for c in content:
        if "WARNING" in c:
            function_log.append(c)
        for task in task_name:
            if task + " executed successfully" in c:
                function_log.append(c)
                try:
                    function_log.append("Logs for Task Name : " + task_name[task_name.index(task) + 1] + "\n")
                except IndexError:
                    pass

    output_file = r"src\functional_log.txt"

    with open(output_file, 'a') as file:
        for f in function_log:
            file.write(f)

