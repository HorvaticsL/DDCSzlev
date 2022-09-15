import configparser

# read config file settings
def read_config(iniFile):
    config = configparser.ConfigParser()
    config.read(iniFile)
    return config