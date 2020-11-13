from selenium import webdriver
from selenium.webdriver.common.keys import Keys  # Allows access to non character keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import openpyxl
import re

def locate_by_name(web_driver, name):
    """Clicks on something by name."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.NAME, name))).click()
    # Returns nothing

def get_by_name(web_driver, name):
    """Returns something by name."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.NAME, name)))

def locate_by_id(web_driver, id):
    """Clicks on something by id."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.ID, id))).click()
    # Returns nothing

def get_by_id(web_driver, id):
    """Returns something by id."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.ID, id)))

def locate_by_class(web_driver, class_name):
    """Clicks on something by class name."""
    WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, class_name))).click()
    # Returns nothing

def get_by_class(web_driver, class_name):
    """Returns something by class name."""
    return WebDriverWait(web_driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, class_name)))

def select_drop(web_driver, id, value):
    """Chooses an option by value in a dropdown box by id."""
    select = Select(web_driver.find_element_by_id(id))
    select.select_by_value(value)
    # Returns nothing

def IT(driver):
    """Create an excel sheet detailing the use of components."""
    parts = {}
    containers = []
    def check_container():
        part = get_by_name(driver, "txtPartNo").get_attribute("value")
        if part not in parts:
            parts[part] = [[]]
        else:
            parts[part].append([])

        parts[part][-1].append(get_by_id(driver, "lblFormTitle"))
        check_history()

    def check_history():
        locate_by_id(driver, "lnkContainerHistory")
        

    # Navigate to inventory tracking menu
    menuNodes = ["tableMenuNode1", "tableMenuNode4", "tableMenuNode3", "tableMenuNode1"]
    for node in menuNodes:
        locate_by_id(driver, node)
        time.sleep(0.5)

    # Fill out the search criteria
    time.sleep(2)
    select_drop(driver, "Layout1_el_385623", "Last_365_Days")
    locs = ["SR00", "SR01", "SR02", "SR03", "SR04", "SR05", "SR06", "SR07", "SR08", "SR09",
            "C00", "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09"]
    for loc in locs:  # Cycle through locations
        input_box = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "Layout1_el_6102")))
        input_box.clear()
        input_box.send_keys(loc)
        locate_by_id(driver, "Layout1_el_56")
        time.sleep(1)
        # Find length of results
        links = driver.find_elements_by_xpath("//a[@href]")  # Get every link
        loc_links = 0  # This variable stores the amount of containers in one location
        for link in links:
            if re.search("ContainerForm", link.get_attribute("href")):
                loc_links += 1

        for index in range(loc_links):  # Cycle through result entries in one location
            links = driver.find_elements_by_xpath("//a[@href]")  # Refresh the links
            encountered = 0
            for link in links:  # Cycle through all links in one location
                if re.search("ContainerForm", link.get_attribute("href")):
                    if encountered == index:
                        link.click()
                        time.sleep(1)
                        locate_by_id(driver, "btnBack_Label")
                        print(parts)
                        raise(Exception)
                        break
                    encountered += 1

    input("Program Pause")


try:
    # Getting into Plex
    driver = webdriver.Chrome("chromedriver.exe")
    driver.get("https://www.plexonline.com/modules/systemadministration/login/index.aspx?")
    # Will need to change the credentials later
    driver.find_element_by_name("txtUserID").send_keys("w.Andre.Le")
    driver.find_element_by_name("txtPassword").send_keys("ThisExpires7")
    driver.find_element_by_name("txtCompanyCode").send_keys("wanco")
    locate_by_id(driver, "btnLogin")
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(3)
    IT(driver)
except Exception as e:
    print("An error was encountered:")
    print(e)
finally:
    driver.quit()