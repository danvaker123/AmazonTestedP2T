from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time


def initialize_browser():
    options = Options()
    options.add_argument("--start-maximized")  # Start browser maximized
    service = Service('path/to/chromedriver')  # Specify the path to your ChromeDriver
    browser = webdriver.Chrome(service=service, options=options)

    return browser


def execute_action(browser, action_type, locator_type, locator_value, input_value=None):
    try:
        if action_type in ["Click", "JavaScript Click", "Check Box"]:
            element = locate_element(browser, locator_type, locator_value)

            if action_type == "Click":
                element.click()
            elif action_type == "JavaScript Click":
                browser.execute_script("arguments[0].click();", element)
            elif action_type == "Check Box":
                if not element.is_selected():
                    element.click()

        elif action_type == "Send Keys":
            element = locate_element(browser, locator_type, locator_value)
            element.clear()  # Clear existing text before sending keys
            element.send_keys(input_value)

        elif action_type == "Dynamic":
            # Handle dynamic actions (to be defined based on your needs)
            pass

        elif action_type == "Retrieve Value":
            element = locate_element(browser, locator_type, locator_value)
            return element.text  # Or another attribute if needed

        elif action_type == "Drop Down":
            # Implement dropdown selection logic here
            pass

        elif action_type == "Switch Window":
            browser.switch_to.window(browser.window_handles[-1])

        elif action_type == "Close Window":
            browser.close()
            browser.switch_to.window(browser.window_handles[0])  # Switch back to the original window

        elif action_type == "Time Sleep":
            time.sleep(float(input_value))  # Input value should be the number of seconds to sleep

        else:
            print(f"Unsupported action type: {action_type}")

    except Exception as e:
        print(f"Error executing action '{action_type}' with locator '{locator_value}': {e}")


def locate_element(browser, locator_type, locator_value):
    """Helper function to locate elements based on the locator type."""
    if locator_type == "ID":
        return browser.find_element(By.ID, locator_value)
    elif locator_type == "XPATH":
        return browser.find_element(By.XPATH, locator_value)
    elif locator_type == "NAME":
        return browser.find_element(By.NAME, locator_value)
    elif locator_type == "CSS_SELECTOR":
        return browser.find_element(By.CSS_SELECTOR, locator_value)
    elif locator_type == "LINK_TEXT":
        return browser.find_element(By.LINK_TEXT, locator_value)
    elif locator_type == "PARTIAL_LINK_TEXT":
        return browser.find_element(By.PARTIAL_LINK_TEXT, locator_value)
    else:
        raise ValueError("Unsupported locator type")


def close_browser(browser):
    browser.quit()
