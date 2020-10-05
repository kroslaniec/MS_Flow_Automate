import sys
import keyboard
import pyautogui
import time
import pyperclip

from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime

driver = webdriver.Chrome("C:\\chromedriver\\chromedriver.exe")


# Used to log in into website, at the end check if site is fully loaded by searching for XPATH
def log_into_dxc():

    username = 'XYZ'
    password = 'XYZ'

    project_url = 'XYZ'

    driver.get(project_url)

    wait_for_page("//*[@id='i0116']")

    email_textbox = driver.find_element_by_xpath("//*[@id='i0116']")
    email_textbox.send_keys(username)
    driver.find_element_by_xpath("//*[@id='idSIButton9']").send_keys(Keys.ENTER)

    wait_for_page("//*[@id='passwordInput']")

    password_textbox = driver.find_element_by_xpath("//*[@id='passwordInput']")
    password_textbox.send_keys(password)
    driver.find_element_by_xpath("//*[@id='submitButton']").send_keys(Keys.ENTER)

    wait_for_page('//*[@id="content-container"]/section/landing/div/react-app/div/section/section/section/main/'
                  'div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div/div[2]/div[1]/div/div[1]/div/button')

    if driver.current_url == project_url:
        print("Log in successful")

    else:
        sys.exit("You failed to log-in - Try one more time")


# Generate link to the PDF file and copying and formatting the number from DOC by using proper function
def open_file(id_number, file_name):

    link_to_file = 'XYZ/' + str(id_number) + '/' + file_name

    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(link_to_file)

    des_number = copy_and_format()

    try:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except NoSuchWindowException:
        pass
    return des_number


# Copy to clipboard, wait for enter and reformat the number to fit the project requirements
def copy_and_format():
    pyperclip.copy('')
    keyboard.wait('enter')
    des_number = pyperclip.paste()
    new_des_number = ''

    for i in range(0, len(des_number)):
        if des_number[i] != " ":
            new_des_number = new_des_number + des_number[i]

    des_number = new_des_number

    if len(des_number) == 8:
        pass
    elif len(des_number) < 8:
        while len(des_number) < 8:
            des_number = '0' + des_number
    elif len(des_number) > 8:
        des_number = des_number[-8:]

    return des_number


# Function to find all attributes from Microsoft Flow - based on XPATH
# One of XPATH element changes due to line's number, so I had to make some preparation for that
def find_line_attributes(counter):

    id_number_find_line_attributes = 0

    next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/section/' \
                                     'section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div/d' \
                                     'iv[2]/div[1]/div/div[1]/div/button'

    if counter < 11:

        next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/' \
                                         'section/section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/' \
                                         'div[' + str(counter) + ']/div/div[2]/div[1]/div/div[1]/div/button'

    elif 11 <= counter < 21:

        counter = counter - 10

        next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/' \
                                         'section/section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[2]/' \
                                         'div[' + str(counter) + ']/div/div[2]/div[1]/div/div[1]/div/button'

    elif 21 <= counter < 31:

        counter = counter - 20

        next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/' \
                                         'section/section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[3]/' \
                                         'div[' + str(counter) + ']/div/div[2]/div[1]/div/div[1]/div/button'

    elif 31 <= counter < 41:

        counter = counter - 30

        next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/' \
                                         'section/section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[4]/' \
                                         'div[' + str(counter) + ']/div/div[2]/div[1]/div/div[1]/div/button'

    elif 41 <= counter < 51:

        counter = counter - 40

        next_line_find_line_attributes = '//*[@id="content-container"]/section/landing/div/react-app/div/section/' \
                                         'section/section/main/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[5]/' \
                                         'div[' + str(counter) + ']/div/div[2]/div[1]/div/div[1]/div/button'

    temp_find_line_attributes = driver.find_element_by_xpath(next_line_find_line_attributes).text
    file_name_splitted = temp_find_line_attributes.split()

    if temp_find_line_attributes[-1] == ',':

        file_name = file_name_splitted[-1][:-1]

    else:

        file_name = file_name_splitted[-1]

    if "ID: " in temp_find_line_attributes:

        id_number_find_line_attributes = temp_find_line_attributes.split('ID: ', maxsplit=1)[-1].split(maxsplit=1)[0][:-1]

    else:

        for i in temp_find_line_attributes.split():

            if i.startswith("ID:"):

                id_number_find_line_attributes = i[3:]

    return file_name, temp_find_line_attributes, next_line_find_line_attributes, id_number_find_line_attributes


# Main function, using some other to paste DES number into Flow's form
def find_des_number(counter):

    file_name_tif, temp, next_line, id_number = find_line_attributes(counter)
    file_name_pdf = file_name_tif.replace(".tif", ".pdf")
    des_number = open_file(id_number, file_name_pdf)
    driver.find_element_by_xpath(next_line).click()

    print("ID: " + id_number + ", DES number: ", des_number)

    warunek = True

    while warunek:

        if pyautogui.locateOnScreen('deny.png', region=(1000, 1300, 200, 150)) is None:

            time.sleep(1)

        else:

            warunek = False

    pyautogui.click(1100, 680)
    pyautogui.click(1100, 720)
    pyautogui.click(1100, 820)
    keyboard.write(des_number)
    pyautogui.click(990, 1360)

    now = datetime.now()

    current_time = now.strftime("%H:%M:%S")
    save_to_report_file(id_number, des_number, current_time)

    time.sleep(5)
    pyautogui.click(1245, 155)


# When there's no any line to process - refresh the page. Check if it's loaded by searching for XPATH
def refresh_page():

    time.sleep(5)
    driver.switch_to.window(driver.window_handles[0])
    driver.refresh()
    wait_for_page('//*[@id="content-container"]/section/landing/div/react-app/div/section/section/section/main/'
                  'div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div/div[2]/div[1]/div/div[1]/div/button')


# Used to check if the page is fully loaded - if it's not, refresh it
def wait_for_page(wanted_xpath):

    try:
        wait = WebDriverWait(driver, 20)
        men_menu = wait.until(ec.visibility_of_element_located((By.XPATH, wanted_xpath)))
        ActionChains(driver).move_to_element(men_menu).perform()

    except TimeoutException:
        refresh_page()


# Additional module to save work progress into a TXT file
def save_to_report_file(id_number, des_number, actual_time):

    with open('report.txt', 'a') as file_object:
        file_object.write("ID: " + id_number + "; DES Number: " + des_number + "; Time: " + actual_time)
        file_object.write("\n")
