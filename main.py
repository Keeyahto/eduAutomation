import argparse
import openpyxl
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import datetime

import config

def parse_args():
    parser = argparse.ArgumentParser(description='Load Excel data into form using Selenium')
    parser.add_argument('-f', '--excel_file', type=str, help='path to the csv file', default='students.xlsx')
    parser.add_argument('--login', type=str, help='username for the form', default=config.login)
    parser.add_argument('--password', type=str, help='password for the form', default=config.password)

    return parser.parse_args()

def get_students(args):
    wb = openpyxl.load_workbook(args.excel_file)
    ws = wb.active
    students = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(row):
            students.append(row)
        else:
            print(f"Студент номер {row[0]} не будет загружен! Заполните все данные")
    return students

def login_edu(args, driver):
    log_el = driver.find_element(By.NAME, 'main_login2')
    log_el.send_keys(args.login)
    pwd_el = driver.find_element(By.NAME, 'main_password2')
    pwd_el.send_keys(args.password)
    login_btn = driver.find_element(By.XPATH, '//*[contains(concat( " ", @class, " " ), concat( " ", "button--green", " " ))]')
    login_btn.click()

def find_el_by(by, locator, waiter):
    while True:
        try:
            el = waiter.until(EC.element_to_be_clickable((by, locator)))
            el.click()
        except:
            continue
        else:
            return el


def upload_student(driver, student):
    waiter = WebDriverWait(driver, 10)
    while True:
        try:
            create_new_el = waiter.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.add-element')))
            create_new_el.click()
        except:
            continue
        else:
            break



    last_name_el = find_el_by(By.ID, 'applicant-last-name', waiter)
    first_name_el = find_el_by(By.ID, 'applicant-first-name', waiter)
    middle_name_el = find_el_by(By.ID, 'applicant-middle-name', waiter)
    birth_date_el = find_el_by(By.ID, 'birth-date-a', waiter)
    gender_el = Select(find_el_by(By.ID, 'gender-select', waiter))
    snils_el = find_el_by(By.ID, 'snils-a', waiter)
    phone_el = find_el_by(By.ID, 'applicant-phone', waiter)
    pasp_series_el = find_el_by(By.ID, 'doc-series-a', waiter)
    pasp_num_el = find_el_by(By.ID, 'doc-number-a', waiter)
    pasp_date = find_el_by(By.ID, 'passport-date', waiter)
    issuer_el = find_el_by(By.ID, 'issuer', waiter)
    doc_type_el = Select(driver.find_element(By.ID, 'spo-doc-type'))
    doc_series_el = find_el_by(By.ID, 'spo-doc-series', waiter)
    doc_num_el = find_el_by(By.ID, 'spo-doc-num', waiter)
    doc_year_el = find_el_by(By.ID, 'spo-doc-year', waiter)
    doc_mean_el = find_el_by(By.ID, 'spo-doc-gpa', waiter)
    course_el = find_el_by(By.ID, 'spo-course-num', waiter)
    is_payment_el = Select(find_el_by(By.ID, 'spo-payment', waiter))

    course_el.send_keys(student[1])
    last_name_el.send_keys(student[3])
    first_name_el.send_keys(student[4])
    middle_name_el.send_keys(student[5])
    birth_date_el.send_keys(student[6].strftime('%d.%m.%Y'))
    snils_el.send_keys(student[7])
    phone_el.send_keys(student[8])
    pasp_series_el.send_keys(student[9])
    pasp_num_el.send_keys(student[10])
    pasp_date.send_keys(student[11].strftime('%d.%m.%Y'))
    issuer_el.send_keys(student[12])
    if student[13].lower() == 'аттестат':
        doc_type_el.select_by_value('attestat')
    elif student[13].lower() == 'диплом':
        doc_type_el.select_by_value('diplom')
    elif student[13].lower() == 'cправка':
        doc_type_el.select_by_value('spravka')
    elif student[13].lower() == 'cвидетельство об обучении':
        doc_type_el.select_by_value('svidetelstvo')
    doc_series_el.send_keys(student[14])
    doc_num_el.send_keys(student[15])
    doc_year_el.send_keys(student[16])
    doc_mean_el.send_keys(student[17])
    if student[18].lower() == 'бюджет':
        is_payment_el.select_by_value('byudzhet')
    else:
        is_payment_el.select_by_value('kommercheskaya')

    find_el_by(By.CLASS_NAME, 'save', waiter)




def edu_automation(args, students):
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.implicitly_wait(10)
    driver.get('https://edu.tatar.ru/login')
    login_edu(args, driver)
    driver.get('https://edu.tatar.ru/ased/#!/documents/DocumentClaimEnrollmentSPO')
    for student in students:
        print('Происходит загрузка:', student)
        upload_student(driver, student)



def main():
    args = parse_args()
    students = get_students(args)
    edu_automation(args, students)




if __name__ == "__main__":
    main()