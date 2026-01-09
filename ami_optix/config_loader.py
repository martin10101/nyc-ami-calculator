import yaml
import os

# Construct a robust path to the config file, assuming it's in the project root.
# This makes the loader independent of where the main script is run from.
CONFIG_FILE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'rules_config.yml'))

def load_config():
    """
    Loads the rules and preferences from the main rules_config.yml file.

    Returns:
        dict: A dictionary containing the configuration.

    Raises:
        FileNotFoundError: If the rules_config.yml file cannot be found.
        IOError: If there is an error reading or parsing the file.
    """
    if not os.path.exists(CONFIG_FILE_PATH):
        raise FileNotFoundError(f"Configuration file not found. Expected at: {CONFIG_FILE_PATH}")

    try:
        with open(CONFIG_FILE_PATH, 'r') as f:
            config = yaml.safe_load(f)
            if config is None:
                raise ValueError("Configuration file is empty or invalid.")
            return config
    except (yaml.YAMLError, ValueError) as e:
        raise IOError(f"Error parsing the configuration file '{CONFIG_FILE_PATH}': {e}")
    except Exception as e:
        raise IOError(f"An unexpected error occurred while reading the config file: {e}")
