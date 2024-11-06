import logging
from Extras.browser_automation import initialize_browser, execute_action, close_browser
from Extras.config_handler import load_config
from Extras.input_handler import load_user_input

# Set up logging
logging.basicConfig(filename='logs/task_log.txt', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')
error_logger = logging.getLogger('error_logger')
error_logger.addHandler(logging.FileHandler('logs/error_log.txt'))


def execute_subtasks(config_file, input_file):
    # Load configuration and user input
    config_data = load_config(config_file)  # Load YAML config
    user_input_data = load_user_input(input_file)  # Load user input

    for task in user_input_data['tasks']:
        task_name = task['Task Name']
        subtask_id = task['Subtask ID']

        # Get relevant actions from the config based on Task Name and Subtask ID
        relevant_actions = [
            action for action in config_data['tasks']
            if action['Task Name'] == task_name and action['Subtask ID'] == subtask_id
        ]

        logging.info(f"Executing {task_name} - Subtask ID: {subtask_id}")

        # Open the browser for this subtask
        browser = initialize_browser()

        try:
            for action in relevant_actions:
                action_type = action['Action']
                locator_type = action['Locator']
                locator_value = action['Locator Value']
                input_value = action.get('Value')  # Optional for input actions

                logging.info(f"Performing {action_type} on {locator_type}: {locator_value}")
                execute_action(browser, action_type, locator_type, locator_value, input_value)

                # Special handling for ManageTransmission flow
                if task_name == "ManageTransmission" and action_type == "Send Keys":
                    logging.info("Handling ManageTransmission flow dynamically.")
                    # Additional logic can be added here as needed

        except Exception as e:
            error_logger.error(f"Error during {task_name} - Subtask ID: {subtask_id}: {e}")

        finally:
            # Close the browser after each subtask
            close_browser(browser)
            logging.info(f"Browser closed for {task_name} - Subtask ID: {subtask_id}.")


if __name__ == "__main__":
    config_file = "config/config.yaml"
    user_input_file = "input/input_data.json"

    execute_subtasks(config_file, user_input_file)
