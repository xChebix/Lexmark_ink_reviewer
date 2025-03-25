import pandas as pd

import requests
from bs4 import BeautifulSoup
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException

from report_maker import generate_printer_report 
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Array values 0= printer IP, 1=black ink, 2= cyan, 3= yellow, 4 = magenta, 5 = maintenance kit,
# 6 = imaging unit, 7 = status,

# Access printers_inventory excel and extracts the third column ips
def Extract_Ips():
    data = pd.read_excel("./printers_Inventory.xlsx", "Hoja1")
    ip_addresses = data.iloc[:,2].tolist()
    return ip_addresses


def Extract_Printers_Data():
    ip_list = Extract_Ips()
    printers_data = []
    for ip in ip_list:
        url = f"http://{ip}/"

        #Using try and catch to detect the type of UI by error
        # Wait to the page to respond, if not continues with the rest of IP
        try:
            driver.get(url)
        except WebDriverException as e:
            print(f"Failed to connect to: {url}: {e}")
            continue
        try:
            # Extract the data from the web page and appends
            data = Onlyblack_UI_printer(driver)
            printers_data.append({
                "ip" : ip,
                **data
            })
            print(f"Only black UI detected, with IP: {ip}")
        except Exception as e:
            try:
                data = Allinks_UI_printer(driver)
                printers_data.append({
                    "ip" : ip,
                    **data
                })

                print(f"All inks UI detected, with IP: {ip}")

            except Exception as e:
                try:
                    data = Old_UI_printer(url)
                    printers_data.append({
                        "ip" : ip,
                        **data
                    })
                    print(f"Old UI detected, with IP: {ip}")
                except TimeoutError as e:
                    print(f"Failed to connect to: {url}: {e}")
                    continue
    return printers_data

def Detect_Low_levels(printers_data):
    low_level = []
    keys_to_check = ["black_ink", "cyan_ink", "yellow_ink", "magenta_ink", "maintenance_kit", "imaging_unit", "status"]
    
    for printer in printers_data:
        # Flag for detecting low level
        has_low_level = False
        for key in keys_to_check:
            if key == "status" and printer[key] == "Nearly Full" or key == "status" and printer[key] == "Full":
                has_low_level = True
                break
            
            valuestr = printer[key]
            if valuestr == None:
                continue
            if valuestr == "OK":
                continue
            else:
                value = int(valuestr.strip("%"))
            # Check if value is below 30 to set true the flag    
            if value < 30:
                has_low_level = True
                break
        if has_low_level:
            low_level.append(printer)

    return low_level

def Allinks_UI_printer(driver): 
    wait = WebDriverWait(driver, 3)  # Wait 3 segs to load the page

    div_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "supplyStatusContainer")))
    maintnance_element = wait.until(EC.presence_of_element_located((By.ID, "FuserSuppliesStatus")))
    imaging_element = wait.until(EC.presence_of_element_located((By.ID, "OtherSuppliesStatus")))
    status_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "supplyStatusIndicator")))

    # Find ink levels
    div_html = div_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html,"html.parser")
    black_ink_level = soup.find("div", class_="progress-inner BlackGauge").text
    magenta_ink_level = soup.find("div", class_="progress-inner MagentaGauge").text
    cyan_ink_level = soup.find("div", class_="progress-inner CyanGauge").text
    yellow_ink_level = soup.find("div", class_="progress-inner YellowGauge").text

    #find maintnance kit 
    div_html = maintnance_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html,"html.parser")
    maintnance_level = soup.find("div", class_="progress-inner BlackGauge").text

    #find imaging unit
    div_html = imaging_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html,"html.parser")
    imaging_level = soup.find("div", class_="progress-inner BlackGauge").text

    # Detect waste toner status
    div_html = status_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html,"html.parser")
    try:
        status_level = soup.find("div", class_="statusIndicator ok selected").text
    except Exception as e:
        print()
        try:
            status_level = soup.find("div", class_="statusIndicator warning selected").text
        except Exception as e:
            status_level = soup.find("div", class_="statusIndicator error selected").text

    # Print all statuses
    print(f"black ink: {black_ink_level.strip()}%")  # strip para escribir todo el texto en la misma linea
    print(f"magenta ink: {magenta_ink_level.strip()}%")
    print(f"yellow ink: {yellow_ink_level.strip()}%")
    print(f"cyan ink: {cyan_ink_level.strip()}")
    print(f"maintnance kit level: {maintnance_level.strip()}%")
    print(f"imaging unit level: {imaging_level.strip()}%")
    print(f"status level: {status_level.strip()}")

    return {
        "black_ink": black_ink_level.strip(),
        "cyan_ink": cyan_ink_level.strip(),
        "yellow_ink": yellow_ink_level.strip(),
        "magenta_ink": magenta_ink_level.strip(),
        "maintenance_kit": maintnance_level.strip(),
        "imaging_unit": imaging_level.strip(),
        "status": status_level.strip()
    }
def Onlyblack_UI_printer(driver):
    wait = WebDriverWait(driver, 3)
    div_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "supplyStatusContainer")))
    drum_element = wait.until(EC.presence_of_element_located((By.ID, "PCDrumStatus")))
    maintnance_element = wait.until(EC.presence_of_element_located((By.ID, "FuserSuppliesStatus")))
    status_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "binStatusIndicator")))

    # search black ink level
    div_html = div_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html, "html.parser")
    ink_level = soup.find("div", class_="progress-inner BlackGauge").text

    # search imaging unit
    div_html = drum_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html, "html.parser")
    imaging_level = soup.find("div", class_="progress-inner BlackGauge").text

    # search maintnance unit
    div_html = maintnance_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html, "html.parser")
    maintnance_level = soup.find("div", class_="progress-inner BlackGauge").text

    # search waste toner status
    div_html = status_element.get_attribute("outerHTML")
    soup = BeautifulSoup(div_html,"html.parser")
    try:
        status_level = soup.find("div", class_="statusIndicator ok selected").text
    except Exception as e:
        print()
        try:
            status_level = soup.find("div", class_="statusIndicator warning selected").text
        except Exception as e:
            status_level = soup.find("div", class_="statusIndicator error selected").text
    
    # Print levels
    print(f"Black ink level: {ink_level.strip()}%")
    print(f"imaging unit level: {imaging_level.strip()}%")
    print(f"maintnance level: {maintnance_level.strip()}%")
    print(f"waste level: {status_level.strip()}")

    # Array values 0= printer IP, 1=black ink, 2= cyan, 3= yellow, 4 = magenta, 5 = maintenance kit,
    # 6 = imaging unit, 7 = status
    return {
        "black_ink": ink_level.strip(),
        "cyan_ink": None,
        "yellow_ink": None,
        "magenta_ink": None,
        "maintenance_kit": maintnance_level.strip(),
        "imaging_unit": imaging_level.strip(),
        "status": status_level.strip()
    }
#Return values blackink, maintenance and imaging unit levels from old UI
def Old_UI_printer(url):
    # Retrieve HTML
    response = requests.get(url + "cgi-bin/dynamic/printer/PrinterStatus.html")
    soup = BeautifulSoup(response.text, "html.parser")

    #Find ink level
    td_element = soup.find("td", {"bgcolor": "#000000"})
    width = td_element.get("width") 
    blackink_level = width.rstrip("%")

    #Find maintnance kit and imaging unit
    td_element = soup.find_all("td")
    maintenance_kit_level = None
    imaging_unit_level = None
    bin_status = None
    # search on all td tags where does the following text is, if there is a match it goes to his sibling
    for td in td_element:
        if "Maintenance Kit Life Remaining:" in td.text:
            maintenance_kit_level = td.find_next_sibling("td").text
        elif "Imaging Unit Life Remaining:" in td.text:
            imaging_unit_level = td.find_next_sibling("td").text
    
    print(f"Black ink level: {blackink_level}")
    print("Maintenance Kit Level:", maintenance_kit_level.rstrip("%"))
    print("Imaging Unit Level:", imaging_unit_level.rstrip("%"))

    return {
        "black_ink": blackink_level.strip(),
        "cyan_ink": None,
        "yellow_ink": None,
        "magenta_ink": None,
        "maintenance_kit": maintenance_kit_level.strip("%"),
        "imaging_unit": imaging_unit_level.strip("%"),
        "status": None
    }


def main():
    image_path = "./encabezado.png"
    printers_data = Extract_Printers_Data()
    low_level_printers = Detect_Low_levels(printers_data)
    date_str = datetime.now().strftime("%d-%m-%Y")
    filename = fr"C:\Users\setchyberre\Documents\Informes_Tintas\printer_report {date_str}.pdf"
    generate_printer_report(filename,low_level_printers,printers_data,image_path)
    #print(low_level_printers)

if __name__ == "__main__":
    main()