from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import shutil
import platform

# PATH where IEDriverServer needs to be copied to, C:\Users\<user>\AppData\Local\Microsoft\WindowsApps
# get current username
username = os.getlogin()

def get_arch():

    # find the architecture and save the correct file in _arch
    try:
        _arch = ""
        # check if IEDriverServer.exe exists or not
        if 'IEDriverServer.exe' not in os.listdir(f"C:\\Users\\{username}\\AppData\\Local\\Microsoft\\WindowsApps"):

            # IEDriverServer.exe doesn't exist, check the architecture and copy the right file over from \\source_to_file
            # on my server, I have the files saved as _x64\x86.exe

            if '64bit' in platform.architecture():
                _arch = "IEDriverServer_x64.exe"

            else:
                _arch = "IEDriverServer_x86.exe"
            # the actual name of the file needs to be exactly "IEDriverServer.exe"
            shutil.copyfile(f"\\\\source_to_file\\{_arch}",\
                            f"C:\\Users\\{username}\\AppData\\Local\\Microsoft\\WindowsApps\\IEDriverServer.exe")
    except:
        print("Error occured!")
        

def start_instance():
    try:
        ie_options = webdriver.IeOptions()
        ie_options.attach_to_edge_chrome = True
        
        # Path to Microsoft Edge
        ie_options.edge_executable_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
        
        # Set zoom level to 100% or ignore it, otherwise Edge will run in IE compat mode, but will complain about the zoom level and driver.get() won't use the provided url
        ie_options.ignore_zoom_level = True
       
        driver = webdriver.Ie(options=ie_options)
       
        # Maximize the window, just QoL
        driver.maximize_window()

        driver.get("url_to_be_opened")

    except:
        print("Error occured!")

if __name__ == "__main__":
    get_arch()
    start_instance()
