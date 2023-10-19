# Module to check and install all dependencies for running para_loss.py.
#Does not install OSGeo4W. 

import importlib
import subprocess

# List of required modules+versions
required_modules = {
    'Selenium': '4.11',
    'webdriver_manager': '4.0',
    'requests': '2.31',
    'openpyxl': '3.1',
    'csv': '1.0'
}

# Check if a module is installed
def is_module_installed(module_name, module_version=None):
    try:
        module = importlib.import_module(module_name)
        if module_version is not None:
            return module.__version__ == module_version
        return True
    except ImportError:
        return False

# Install module using pip
def install_module(module_name, module_version=None):
    if module_version is None:
        subprocess.call(['pip', 'install', module_name])
    else:
        subprocess.call(['pip', 'install', f'{module_name}=={module_version}'])

# Check to install required modules
for module_name, module_version in required_modules.items():
    if not is_module_installed(module_name, module_version):
        print(f"{module_name} is not installed. Installing...")
        install_module(module_name, module_version)
        print(f"{module_name} has been successfully installed.")

print("All required modules are installed.")