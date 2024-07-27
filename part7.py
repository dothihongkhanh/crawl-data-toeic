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

folder_name = "part7"
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

folder_test = folder_name + "/ReadingComprehension_01"
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
        driver.find_element(By.XPATH, '//*[@id="test-info"]/div[3]/div/table/tbody/tr[5]/td[4]/a').click()
        time.sleep(3)
        driver.find_element(By.PARTIAL_LINK_TEXT, "Xem chi tiết đáp án").click()
        time.sleep(3)

        data_to_save = []

        image_folder = folder_test +"/images_p7"
        os.makedirs(image_folder, exist_ok=True)

        quiz_items = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "question-group-wrapper")))
        for idx, item in enumerate(quiz_items, start=1):
            current_time = int(time.time())
            random_number = random.randint(0, 9999)
            question_id = f"{current_time}{random_number:04}" 

            image_items = item.find_elements(By.CLASS_NAME, "context-content.text-highlightable")
            image_names = []  # Tạo một danh sách để lưu tên của tất cả các hình ảnh trong mỗi phần tử
            for image_item in image_items:
                image_elements = image_item.find_elements(By.TAG_NAME, "img")
                for img_idx, image_element in enumerate(image_elements, start=1):
                    image_url = image_element.get_attribute("src")
                    image_name = f"{question_id}_image_{idx}_{img_idx}.png"
                    image_path = os.path.join(image_folder, image_name)
                    with open(image_path, 'wb') as image_file:
                        image_file.write(requests.get(image_url).content)
                    image_names.append(image_name)

            transcript_element = item.find_element(By.CLASS_NAME, "context-content.context-transcript.text-highlightable")
            time.sleep(5)
            transcript_text = transcript_element.find_element(By.CLASS_NAME, "collapse").text

            question_group_image_elements = item.find_elements(By.CLASS_NAME, "question-twocols-right")
            for question_group_image_element in question_group_image_elements:
                explain_texts = []
                explain_elements = question_group_image_element.find_elements(By.CLASS_NAME, "question-explanation-wrapper")
                for explain_element in explain_elements:
                    time.sleep(5)
                    explain_texts.append(explain_element.find_element(By.CLASS_NAME, "collapse").text)

            question_elements = item.find_elements(By.CLASS_NAME, "question-wrapper")
            for index, question_element in enumerate(question_elements):
                question_number_element = question_element.find_element(By.CLASS_NAME, "question-number")
                question_number = question_number_element.text
                question_text = question_element.find_element(By.CLASS_NAME, "question-text").text                
                answer_elements = question_element.find_elements(By.CLASS_NAME, "form-check-label")
                answer_texts = [answer_element.text for answer_element in answer_elements]
                
                answer_correct = question_element.find_element(By.CLASS_NAME, "text-success").text.split(":")[1].strip()
                
                data_to_save.append([question_id, transcript_text, question_number, question_text, *answer_texts, answer_correct, explain_texts[index], *image_names])

        save_to_excel(data_to_save)

    finally:
        driver.quit()

email = ""
password = ""

def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.append(["question_id", "transcript", "question_number", "title_question", "answer1", "answer2", "answer3", "answer4", "correct_answer", "explanation", "image1", "image2", "image3"])
    for row in data:
        ws.append(row)

    filename = folder_test + "/quiz_data_p7.xlsx"
    wb.save(filename)
    print("Data saved to", filename)


login_url = "https://study4.com/login/"

login_and_redirect(email, password, login_url)
