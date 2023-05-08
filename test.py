# File created by: NOLAN AGAH

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# drug1 = input("First drug: ")
# drug2 = input("Second drug: ")
drug1 = 'warfarin'
drug2 = 'aspir 81'

def find_drug_interactions():
    driver_path = "C:/path/to/your/edgedriver.exe"
    service = Service(executable_path=driver_path)
    driver = webdriver.Edge(service=service)
    driver.get("https://www.drugs.com/interaction/list/?drug_list=")
    driver.maximize_window()
    elem1 = driver.find_element(By.ID, "livesearch-interaction")
    elem1.clear()
    elem1.send_keys(drug1)
    elem1.send_keys(Keys.ARROW_DOWN)
    elem1.send_keys(Keys.RETURN)
    elem2 = driver.find_element(By.ID, "livesearch-interaction")
    elem2.clear()
    elem2.send_keys(drug2)
    elem2.send_keys(Keys.ARROW_DOWN)
    elem2.send_keys(Keys.RETURN)

    # wait for the "Check Interactions" button to become clickable
    wait = WebDriverWait(driver, 10)
    check_interactions_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@class="ddc-btn" and contains(@href, "interactions-check.php")]')))

    # click the "Check Interactions" button
    check_interactions_button.click()

    parent_element = driver.find_element(By.XPATH, "//*[contains(@class, 'interactions-reference')]")

    # Find the second child <p> element using XPath
    child_elements = parent_element.find_elements(By.XPATH, ".//p")
    second_child_element = child_elements[1]  # get the second element from the list

    element_text = second_child_element.text

    print(element_text)

    # close the browser after you press Enter
    driver.quit()

# def print_drug_interactions():
    


find_drug_interactions()