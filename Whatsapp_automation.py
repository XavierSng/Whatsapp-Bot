""" Program to send bulk customized message through WhatsApp web application"""

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import pandas
import time
import random

# Load the chrome driver
driver = webdriver.Chrome()
count = 0

# Open WhatsApp URL in Chrome browser
driver.get("https://web.whatsapp.com/")
input('Scanning of QR Code, press enter when done.')
wait = WebDriverWait(driver, 10)

# Read data from excel
excel_data = pandas.read_excel(r"C:\Users\sngxa\PycharmProjects\pythonProject\Whatsapp automation\Customer bulk email data.xlsx", sheet_name='Customers')

# Iterate excel rows till to finish
for column in excel_data['Contact'].tolist():

    # Assign customized message
    message = """
*BCA Sandbox Program (Tiler & Plasterer)*

Dear Sir/ Madam,

BCA has just launched a Sandbox Program to bring in *trained Myanmar tilers and Plasterers*. We are one of the appointed agency to supply these workers.

Besides *tilers/ plasterers* from this Sandbox program, we also supply other *SECK (R1)/ Low levy construction workers* trained in:

- Timber/ System Formwork
- Plumbing & Pipe Fitting
- Ceiling Partition
- Aircon Servicing & Installation
- Welding

Kindly contact us at *8600 2330/ 8600 8840* for more info. No agency fees to your company

Thanks & regards,
Eileen Siah 
Times Recruitment Pte Ltd
MOM License 16C7960
Email : eileen.siah@timesrecruitment.sg
"""

    # Locate search box through x_path
    searchbar_xpath = "//div[@class='_2vDPL']//p[@class='selectable-text copyable-text iq0m558w']"
    person_title = wait.until(lambda driver: driver.find_element(By.XPATH, searchbar_xpath))

    # Clear search box if any contact number is written in it
    person_title.click()
    time.sleep(random.uniform(1,2))
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(
                Keys.DELETE).perform()

    # Send contact number in search box
    time.sleep(random.uniform(1,2))
    person_title = wait.until(lambda driver: driver.find_element(By.XPATH, searchbar_xpath))

    person_title.send_keys(str(excel_data['Contact'][count]))
    count += 1

    # Wait for 3-5 seconds to search contact number, cannot shorten
    time.sleep(random.uniform(3,5))

    try:
        # Load error message in case unavailability of contact number
        check_contact_xpath = '//*[text()="No chats, contacts or messages found"]'
        check_contact = driver.find_element(By.XPATH, check_contact_xpath)
        print(str(excel_data['Contact'][count]) + " contact not found.")

    except NoSuchElementException:
        
        # Format the message
        person_title.send_keys(Keys.ENTER)
        
        # Time to open chat, might be unnecessary(?)
        time.sleep(random.uniform(1,2))

        attach_xpath = "//span[@data-icon ='clip']"
        attach = wait.until(lambda driver: driver.find_element(By.XPATH, attach_xpath))
        attach.click()

        # Time to open gallery, might be unnecessary(?)
        time.sleep(random.uniform(0,1))
        img_path = r"C:\Users\sngxa\PycharmProjects\pythonProject\Whatsapp automation\tiler_img.jpeg"

        gallery_xpath = "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']"
        gallery = wait.until(lambda driver: driver.find_element(By.XPATH, gallery_xpath))
        gallery.send_keys(img_path)

        # Time to open image send portion, necessary **CANNOT** shorten
        time.sleep(random.uniform(0,1))

        emoji_xpath = "//span[@data-icon= 'media-editor-emoji']"
        emoji_check = wait.until(lambda driver: driver.find_element(By.XPATH, emoji_xpath))

        actions = ActionChains(driver)
        actions.pause(random.uniform(1,3))
        actions.send_keys(Keys.ENTER)
        actions.pause(random.uniform(1,3))
        actions.send_keys(Keys.ENTER)
        actions.perform()
        # CANNOT REMOVE
        time.sleep(random.uniform(3,5))

        for part in message.split('\n'):
            # Chatbox locator
            chatbox_xpath = "//div[@title='Type a message']//p[@class='selectable-text copyable-text iq0m558w']"
            chatbox = wait.until(lambda driver: driver.find_element(By.XPATH, chatbox_xpath))

            chatbox.send_keys(part)
            ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                Keys.ENTER).perform()
            
        time.sleep(random.uniform(1,3))
        chatbox.send_keys(Keys.ENTER)

# Loading of last image sent, **CANNOT** shorten
time.sleep(3)
