from selenium import webdriver 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import random
import time
import requests
import os
from openpyxl import Workbook

folder_name = "part5"
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

folder_test = folder_name + "/IncompleteSentences_01"
if not os.path.exists(folder_test):
    os.makedirs(folder_test)

def login_and_redirect(email, password, login_url):
    path = r"D:\crawl-toeic\chromedriver.exe"
    ser = Service(path)
    
    driver = webdriver.Chrome(service=ser)
    
    try:
        driver.get(login_url)

        login_button_in_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div/div/div/div/div[1]/div[1]/span')))
        login_button_in_dropdown.click()
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="email"]')))
        
        driver.find_element(By.XPATH, '//*[@id="email"]').send_keys(email)
        driver.find_element(By.XPATH, '//*[@id="pass"]').send_keys(password)
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginbutton"]'))).click()
        driver.find_element(By.XPATH, '//*[@id="navbar-collapse"]/ul/li[3]/a').click()
        driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[1]/div/div/div[2]/div/ul/li[4]/a').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[3]/div/div[1]/div[1]/div/div[3]/div/div/a[2]').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="test-info"]/div[3]/div/table/tbody/tr[10]/td[4]/a').click()
        time.sleep(3)
        driver.find_element(By.PARTIAL_LINK_TEXT, "Xem chi tiết đáp án").click()
        time.sleep(3)

        data_to_save = []
        
        quiz_items = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH, '//*[@id="partcontent-9784"]/div[2]')))
        for quiz_item in quiz_items:
            question_wrappers = quiz_item.find_elements(By.CLASS_NAME, "question-wrapper")
            question_explain_wrappers = quiz_item.find_elements(By.CLASS_NAME, "question-explanation-wrapper")

            for (question_wrapper, question_explain_wrapper) in zip(question_wrappers, question_explain_wrappers):                
                current_time = int(time.time())
                random_number = random.randint(0, 9999)
                question_id = f"{current_time}{random_number:04}" 
                time.sleep(3)

                question_number_element = question_wrapper.find_element(By.CLASS_NAME, "question-number")
                question_number = question_number_element.text
                question_text = question_wrapper.find_element(By.CLASS_NAME, "question-text").text
                answer_elements = question_wrapper.find_elements(By.CLASS_NAME, "form-check-label")
                answer_texts = [answer_element.text for answer_element in answer_elements]
                
                correct_answer = question_wrapper.find_element(By.CLASS_NAME, "text-success").text.split(":")[1].strip()
                time.sleep(3)
                explanation = question_explain_wrapper.find_element(By.CLASS_NAME, "collapse").text
                    
                data_to_save.append([question_id, question_number, question_text, *answer_texts, correct_answer, explanation])

        save_to_excel(data_to_save)

    finally:
        driver.quit()

email = ""
password = ""

def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.append(["code_part5", "question_number", "title_question", "answer1", "answer2", "answer3", "answer4", "correct_answer", "explanation"])
    for row in data:
        ws.append(row)

    filename = folder_test + "/quiz_data_p5.xlsx"
    wb.save(filename)
    print("Data saved to", filename)

login_url = "https://study4.com/login/"
login_and_redirect(email, password, login_url)
