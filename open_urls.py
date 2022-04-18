# Quick python script to open a list of urls using Chrome, maximize the window and switch to the 1st opened tab for easier management
# Got quite a few network printers at work and decided to automate the opening of the web interface for each printer

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time

URLS = [
    "ip_1", # printer_1,
    "ip_2", # printer_2,
    "ip_3", # printer_3
]

def open_urls():
    
    # set the path to chromedriver
    service = Service("C:\\Users\\<user>\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Python 3.10\\chromedriver")
    service.start()

    driver = webdriver.Remote(service.service_url)
    
    # maximize the window, use driver.get() to load the 1st url in URLS[]
    driver.maximize_window()
    driver.get(URLS[0])

    try:
        # continue with opening each url in URLS[] starting at index 1 (index 0 already opened)
        for url in URLS[1:]:
            # open new tab and go to {url}
            driver.execute_script(f"window.open(\"{url}\");")
            time.sleep(0.5)
            
    except Exception as e:
        print("Error occured! " + e)
        
    # move to the 1st tab, just QoL    
    driver.switch_to.window("firsttab")

if __name__ == "__main__":
    open_urls()
