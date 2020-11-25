from selenium import webdriver
from selenium.webdriver.common.keys import Keys  # Allows access to non character keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
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
    inactive_containers = {}  # A dictionary
    def check_container():
        # Get the container number, part number, and location
        temp_container = ((get_by_id(driver, "lblFormTitle").text).split())[2]
        temp_part = get_by_name(driver, "txtPartNo").get_attribute("value")
        temp_location = get_by_id(driver, "pikLocation").text
        # For now, add these to the dictionary
        inactive_containers[temp_container] = [temp_part, temp_location]

        # Check container history
        locate_by_id(driver, "lnkContainerHistory")
        time.sleep(1)
        no_wraps = driver.find_elements_by_class_name("NoWrap")
        # 22 columns, row indices are multiples of 22 (row one starts at index 0)
        row_nums = int(len(no_wraps) / 22)
        # 18 is the first index in the "Last Action" column
        # 14 is the first index in the "Location" column
        act = inact = 0
        for row in range(row_nums):
            if no_wraps[18+(row*22)].text == "Split Container":
                if no_wraps[14+(row*22)].text == "SR RECEIV":
                    continue
                else:
                    act += 1
            elif no_wraps[18+(row*22)].text == "Cycle Complete" or no_wraps[18+(row*22)].text == "Container Move":
                inact += 1
        # Deem the container active or inactive
        if inact >= act:  # Inactive
            if inact == 0:  # A zero ratio means that inact and act are zero
                inactive_containers[temp_container].append(0)
            elif act > 0:
                inactive_containers[temp_container].append(inact/act)
            else:  # If act is zero, but inact is more than zero, the ratio is infinity
                inactive_containers[temp_container].append('inf')
        else:  # Active
            inactive_containers.popitem()
        # Go back to the results page
        locate_by_class(driver, "left-arrow-purple")
        locate_by_id(driver, "btnBack_Label")

    # Navigate to inventory tracking menu
    action = ActionChains(driver)
    action.key_down(Keys.CONTROL).send_keys('M').key_up(Keys.CONTROL).perform()
    action = 404
    time.sleep(1)
    action = ActionChains(driver)
    action.send_keys("Inventory Tracking").send_keys(Keys.RETURN).perform()
    action = 404
    time.sleep(1)
    action = ActionChains(driver)
    action.send_keys(Keys.RETURN).perform()

    # Fill out the search criteria
    time.sleep(2)
    input_box = get_by_name(driver, "Layout1$el_385621")
    input_box.clear()
    input_box.send_keys("1/1/2000")
    input_box = get_by_name(driver, "Layout1$el_385622")
    input_box.clear()
    input_box.send_keys("11/25/2019")
    #locs = ["SR03", "SR04", "SR05", "SR06", "SR07", "SR08", "SR09", "SR10", "SR11", "SR12", "SR13", "SR14", "SR15",
    #        "C00", "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11"]
    locs = ["C06", "C07", "C08", "C09", "C10", "C11"]
    for loc in locs:  # Cycle through locations
        input_box = get_by_id(driver, "Layout1_el_6102")
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
        if loc_links == 0:  # If so results come up
            continue  # Go to the next location

        for index in range(loc_links):  # Cycle through containers in one location
            links = driver.find_elements_by_xpath("//a[@href]")  # Refresh the links
            encountered = 0
            for link in links:  # Cycle through all container links in one location
                if re.search("ContainerForm", link.get_attribute("href")):
                    if encountered == index:
                        link.click()  # Enter container
                        check_container()
                        break
                    encountered += 1

    # Write data to new excel sheet
    wb_obj = openpyxl.Workbook()
    sheet_obj = wb_obj.active
    headers = ["Container Number", "Part Number", "Location", "Activity Ratio (inactive/active)"]
    for i in range(1, 5):  # Write in headers
        sheet_obj.cell(row=1, column=i).value = headers[i-1]
    for index, key in enumerate(inactive_containers):
        sheet_obj.cell(row=index+2, column=1).value = key
        sheet_obj.cell(row=index+2, column=2).value = inactive_containers[key][0]
        sheet_obj.cell(row=index+2, column=3).value = inactive_containers[key][1]
        sheet_obj.cell(row=index+2, column=4).value = inactive_containers[key][2]
    # This creates a new excel workbook in the program's directory
    wb_obj.save("Inactive Containers List.xlsx")

try:
    # Getting into Plex
    driver = webdriver.Chrome("chromedriver.exe")
    # Link to new Plex site
    driver.get("https://accounts.plex.com/interaction/fea73869-0eda-4f67-b381-c167be521da6#ilp=woW7Rk4HS5ijknMk0L8Jjl8&ie=1606149525001")

    parent = "//form[@class='form-horizontal']//div[@class='plex-idp-wrapper']"  # Allows access to input fields, which are hidden
    # Enter in company code
    form = driver.find_element_by_xpath(parent + "//div[@id='companyCodeInput']//div[@class='col-sm-12']//input[@id='inputCompanyCode3']")
    form.send_keys("wanco")
    action = ActionChains(driver)
    action.send_keys(Keys.RETURN).perform()
    time.sleep(.5)
    # Enter in username
    form = driver.find_element_by_xpath(parent + "//div[@id='usernameInput']//div[@class='col-sm-12']//input[@id='inputUsername3']")
    form.send_keys("w.mc.tester")
    action.perform()
    time.sleep(.5)
    # Enter in password
    form = driver.find_element_by_xpath(parent + "//div[@id='passwordInput']//div[@class='col-sm-12']//input[@id='inputPassword3']")
    form.send_keys("test1wanco")
    action.perform()
    time.sleep(.5)
    
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(3)
    IT(driver)
except Exception as e:
    print("An error was encountered:")
    print(e)
finally:
    driver.quit()