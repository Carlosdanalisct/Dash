from app import CONFIG, CONFIG_PATH, DEFAULT_CONFIG, load_config


def get_config():
    return load_config()


def config_path():
    return CONFIG_PATH

