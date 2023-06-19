# Program to send bulk customized message through WhatsApp web application

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time
import os

# Load the chrome driver
driver = webdriver.Chrome()
count = 0

# Open WhatsApp URL in Chrome browser
driver.get("https://web.whatsapp.com/")
input('Scanning of QR Code')
wait = WebDriverWait(driver, 10)

# Read data from excel
excel_data = pandas.read_excel('Customer bulk email data.xlsx', sheet_name='Customers')

# Iterate excel rows till to finish
for column in excel_data['Name'].tolist():
    # Assign customized message
    message = """
Low Levy Myanmar Construction Workers (R1)

Dear Sir Sir/ Madam,

We supply low levy fresh Myanmar Construction workers with BCA Certificate (SECK - R1) trained in formwork and pipe fitting

- Workers with SEC(K) - R1 Cert
- Fast arrival (Less than 2 weeks)
- Low levy
- No agency fees

Pls contact me if you are keen to employ them.

Thanks & regards,
Eileen Siah 
Times Recruitment Pte Ltd
MOM License 16C7960
"""

    # Locate search box through x_path
    searchbar_xpath = "//div[@class='_2vDPL']//p[@class='selectable-text copyable-text iq0m558w']"
    person_title = wait.until(lambda driver: driver.find_element(By.XPATH, searchbar_xpath))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(str(excel_data['Contact'][count]))
    count = count + 1

    # Wait for 5 seconds to search contact number
    time.sleep(5)

    try:
        # Load error message in case unavailability of contact number
        check_contact_xpath = '//*[text()="No chats, contacts or messages found"]'
        check_contact = driver.find_element(By.XPATH, check_contact_xpath)
    except NoSuchElementException:

        # Format the message
        person_title.send_keys(Keys.ENTER)

        # Chatbox locator
        chatbox_xpath = "//div[@title='Type a message']//p[@class='selectable-text copyable-text iq0m558w']"
        chatbox = driver.find_element(By.XPATH, chatbox_xpath)

        for part in message.split('\n'):
            chatbox.send_keys(part)
            ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                Keys.ENTER).perform()
        chatbox.send_keys(Keys.ENTER)

        attach = driver.find_element(By.XPATH, "//span[@data-icon ='clip']")
        attach.click()

        time.sleep(3)
        gallery = wait.until(lambda driver: driver.find_element(By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']"))
        gallery.send_keys(r"C:\Users\sngxa\PycharmProjects\pythonProject\Construction.jpg")

        time.sleep(3)
        actions = ActionChains(driver)
        actions.send_keys(Keys.ENTER)
        actions.send_keys(Keys.ENTER)
        actions.perform()
        time.sleep(3)

