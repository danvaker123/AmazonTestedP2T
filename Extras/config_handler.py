import yaml

def load_config(file_path):
    """Load configuration from a YAML file."""
    with open(file_path, 'r') as file:
        config = yaml.safe_load(file)
    return config

def get_task_actions(config, task_name):
    """Retrieve actions for a specific task from the config."""
    task_actions = []
    for task in config.get('tasks', []):
        if task['task_name'] == task_name:
            task_actions = task.get('actions', [])
            break
    return task_actions

def get_action_details(action):
    """Extract details for each action."""
    action_details = {
        'type': action.get('type'),
        'locator_type': action.get('locator_type'),
        'locator_value': action.get('locator_value'),
        'input_value': action.get('input_value')
    }
    return action_details

def main():
    config_path = 'config/config.yaml'  # Adjust path as necessary
    config = load_config(config_path)

    # Example usage: Get actions for a specific task
    task_name = "ManageTransmission"  # Change this as needed
    actions = get_task_actions(config, task_name)

    for action in actions:
        details = get_action_details(action)
        print(f"Action Type: {details['type']}, Locator Type: {details['locator_type']}, "
              f"Locator Value: {details['locator_value']}, Input Value: {details['input_value']}")

if __name__ == "__main__":
    main()
