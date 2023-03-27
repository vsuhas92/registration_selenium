import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.support.ui import Select

import pandas

import random


def esic_registration(username, password, xlsx_filename):
    start_time = time.time()

    data = pandas.read_excel(xlsx_filename, sheet_name='Sheet1')

    # insert ESIC column into the dataframe
    data.insert(1, 'New ESIC number',
                '[ERROR]_missing')  # Default value will get replaced if the registration is successful

    no_rows = data.shape[0]
    # print(data.dtypes)

    # Constant Variable
    # phone_number = '9480574070'
    # name = 'testname'
    # dependent_rel = 'Father'
    # dependent_name = 'testfathername'
    # dob = '15/03/2002'
    # marital_status = "Married"
    # gender = 'Male'
    # address = 'testaddress'
    # state = 'Karnataka'
    # district = 'Bangalore'
    # doa = '02/01/2023'
    #
    # nominee_name = "test_name"
    # relationship = "Spouse"
    #
    #
    # ifsc_code = "UTIB0000514"
    # account_number = "123456789101112"
    # account_type = "Savings"
    #
    # preferred_lan = "English / Indian English"

    # esic website
    website = "https://www.esic.in/EmployerPortal/ESICInsurancePortal/Portal_Loginnew.aspx"
    # username = "49000217450000000"
    # password = "esic"

    # Google Chrome
    # if the Google Chrome gets updated then goto- "https://chromedriver.chromium.org/downloads"
    # download "chromedriver_win32.zip"
    # paste "chromedriver.exe" to the path below
    path = "C:\Program Files (x86)\chromedriver.exe"
    s = Service(path)
    driver = webdriver.Chrome(service=s)
    driver.get(website)

    # ESIC website initialization
    try:
        # login
        # Enter UserName
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtUserName"]')))
        element.send_keys(username)

        # Enter password
        driver.find_element(By.XPATH, '//*[@id="txtPassword"]').send_keys(password)

        input('Press enter after entering the captcha')
        driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()

        # 4 clicks
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="div1_close"]')))
        element.click()

        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="btnpnlcolse"]')))
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="btnpnlPwdcolse"]')))
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="btnpnlSanjeevanicolse"]')))
        element.click()

    except:
        print("Failed to do new registration initialization- login, esic website initialization")
        driver.quit()
        quit()

    success = fail = 0
    for row in range(no_rows):
        fail_pt = "None"

        # pandas_ read information from Excel file
        prev_esic_num = str(data.at[row, "Prev ESIC number"])
        phone_number = str(data.at[row, "Contact Mobile No."])
        name = data.at[row, 'Name of the Member (As per documents-In the block capitals )']
        dependent_rel = data.at[row, "Relationship with the employee"]
        dependent_name = data.at[row, "Father's Name (or husband's name in case of married women)"]

        # formatting date to dd/mm/yyyy
        dob_temp = str(data.at[row, "Date of birth (dd/mm/yyyy)"])[0:10]
        dob = dob_temp[8:10] + '/' + dob_temp[5:7] + '/' + dob_temp[0:4]

        marital_status = data.at[row, "Marital status"]
        gender = data.at[row, "Sex"]
        address1 = data.at[row, "Address Line 1"]
        address2 = data.at[row, "Address Line 2"]
        address3 = data.at[row, "Address Line 3"]
        state = data.at[row, "State"]
        district = data.at[row, "District"]

        disp_state = data.at[row, "Dispensary State"]
        disp_district = data.at[row, "Dispensary District"]
        dispensary = data.at[row, "Dispensary"]

        # copy information to family if check box is checked
        same = data.at[row, "Family same as Employee"]
        if same:
            family_disp_state = disp_state
            family_district = disp_district
            family_dispensary = dispensary
        else:
            family_disp_state = data.at[row, "Family Dispensary State"]
            family_district = data.at[row, "Family Dispensary District"]
            family_dispensary = data.at[row, "Family Dispensary"]

        # formatting date to dd/mm/yyyy
        doa_temp = str(data.at[row, "Date of appointment"])[0:10]
        doa = doa_temp[8:10] + '/' + doa_temp[5:7] + '/' + doa_temp[0:4]

        nominee_name = data.at[row, "Nominee Name"]
        relationship = data.at[row, "Relationship with IP"]

        ifsc_code = data.at[row, "IFSC"]
        account_number = str(data.at[row, "Bank Account Number"])
        account_type = data.at[row, "Bank Account Type"]

        preferred_lan = data.at[row, "Preferred Language"]

        try:
            # click on New Registration
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="lnkRegisterNewIP"]')))
            element.click()

            # Switch the driver to pop-up form
            driver.switch_to.window(driver.window_handles[1])

            # new esic number, no previous insurance number
            if prev_esic_num == '-':
                # Select 'NO'
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="ctl00_HomePageContent_rbtnlistIsregistered_1"]')))
                element.click()

                # Handle chrome pop-ups
                time.sleep(2)
                driver.switch_to.alert.accept()

                # Phone number

                # handling phone number if it already exists in the esic database
                # Changing the last digit to try again until a valid phone number is found
                while True:
                    fail_pt = "Phone number"
                    # Enter Phone number
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '// *[ @ id = "ctl00_HomePageContent_txtmobilenumber"]')))
                    element.send_keys(phone_number)

                    fail_pt = "Validating Phone number"
                    # Validate Phone number
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_lnkmobilecheck"]').click()

                    # trying to find the continue button, if not found this means phone number wasn't validated
                    # successfully
                    try:
                        # Click Continue
                        element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HomePageContent_btnContinue"]')))

                        element.click()
                        break

                    except:
                        # CLick back button
                        driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_btncancel"]').click()

                        # new phone number is created by randomizing the last digit of the phone number
                        random_number = random.randrange(0, 9, 1)
                        phone_number = phone_number[0:9] + str(random_number)

                        # important_ this is required
                        time.sleep(10)

                # Handle chrome pop-ups
                driver.switch_to.alert.accept()

                # ----------------------------------------------------------------------------------------------------------------------------------------
                # Fill Employee Registration form

                # name
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextEmpName"]')))
                element.send_keys(name)

                # Dependent Radio Father/ husband
                if dependent_rel == "Father":  # Father
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlFatherOrHus_0"]').click()
                else:  # Husband
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlFatherOrHus_0"]').click()

                # Dependent Name
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextFatherHusName"]').send_keys(
                    dependent_name)

                # dob
                fail_pt = "date of birth"
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTxtIpDate"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_title" and '
                                    '@class="ajax__calendar_title"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_title" and '
                                    '@class="ajax__calendar_title"]').click()

                time.sleep(2)
                for i in range(10):  # 100 years check
                    title_calender = driver.find_element(By.XPATH,
                                                         '//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_title" and @class="ajax__calendar_title"]').text

                    # 2020-2029
                    y1 = int(title_calender[0:4])
                    y2 = int(title_calender[5:9])
                    dob_year_yyyy = int(dob[6:10])

                    if y1 <= dob_year_yyyy <= y2:
                        dob_year_last = dob_year_yyyy % y1
                        dob_year_y = str(dob_year_last)
                        # year
                        # xxxx xxx0 xxx1 xxx2
                        # xxx3 xxx4 xxx5 xxx6
                        # xxx7 xxx8 xxx9 xxxx
                        #
                        year_dic = {"0": "0_1", "1": "0_2", "2": "0_3",
                                    "3": "1_0", "4": "1_1", "5": "1_2", "6": "1_3",
                                    "7": "2_0", "8": "2_1", "9": "2_2"}

                        y = year_dic[dob_year_y]
                        year_xpath = f'//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_year_{y}"]'
                        driver.find_element(By.XPATH, year_xpath).click()
                        time.sleep(2)
                        break
                    else:
                        driver.find_element(By.XPATH,
                                            '//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_prevArrow"]').click()
                        time.sleep(2)

                # month
                # 3 * 4 array
                # Jan Feb Mar Apr
                # May Jun Jul Aug
                # Sep Oct Nov Dec
                month_dic = {"01": "0_0", "02": "0_1", "03": "0_2", "04": "0_3",
                             "05": "1_0", "06": "1_1", "07": "1_2", "08": "1_3",
                             "09": "2_0", "10": "2_1", "11": "2_2", "12": "2_3"}

                dob_month = dob[3:5]
                m = month_dic[dob_month]
                month_xpath = f'//*[@id="ctl00_HomePageContent_CalendarExtenderCtrlTxtEndDate_month_{m}"]'
                driver.find_element(By.XPATH, month_xpath).click()
                time.sleep(2)

                # date
                dob_date_dd = dob[0:2]

                # Checking if the date is "0x", then convert it to 'x'
                if dob_date_dd[0] == '0':
                    dob_date = dob_date_dd[1]
                else:
                    dob_date = dob_date_dd

                dates = driver.find_elements(By.XPATH, '//*[@class="ajax__calendar_day"]')

                # finding 1 thereby removing previous months dates
                i = 0
                for date in dates:
                    if date.text == '1':
                        break
                    i += 1
                dates = dates[i:-1]

                # find the date matching dob
                for date in dates:
                    if date.text == dob_date:
                        date.click()
                        break

                # Marital Status-drop down
                fail_pt = "Martial Status"
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlRDMarried"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(marital_status)

                # Gender
                if gender == "Male":  # Male
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlRDMale_0"]').click()
                elif gender == "Female":  # Female
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlRDMale_1"]').click()
                else:  # TG
                    driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlRDMale_2"]').click()
                time.sleep(2)
                driver.switch_to.alert.accept()

                # Address
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextPresentAddress1"]').send_keys(
                    address1)

                # State drop down
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTxtPresentState"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(state)

                # District drop down
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextPresentDistrict"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(district)

                # Copy Present Address to Permanent Address
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_chkboxCopyPresentAddress"]').click()

                # Dispensary logic

                # Dispensary State selection
                fail_pt = "Dispensary State"
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddlDispensaryState"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(disp_state)

                # Dispensary District
                fail_pt = "Dispensary District"
                try_again = True
                while try_again is True:
                    x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddlDispensaryDistrict"]')
                    drop_down = Select(x)
                    if len(drop_down.options) > 1:
                        # for x in drop_down.options:
                        #     print(x.text)
                        for x in drop_down.options:
                            if x.text == disp_district:
                                drop_down.select_by_visible_text(disp_district)
                                try_again = False
                                break
                    time.sleep(2)

                # Dispensary
                fail_pt = "Dispensary"
                try_again = True
                while try_again is True:
                    x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextDispensary"]')
                    drop_down = Select(x)
                    if len(drop_down.options) > 1:
                        # for x in drop_down.options:
                        #     print(x.text)
                        for x in drop_down.options:
                            if x.text == dispensary:
                                drop_down.select_by_visible_text(dispensary)
                                try_again = False
                                break
                    time.sleep(2)

                fail_pt = "Family Dispensary State"
                # Family Dispensary State selection
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddldependantDispensaryState"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(family_disp_state)

                fail_pt = "Family Dispensary District"
                # Family Dispensary District
                try_again = True
                while try_again is True:
                    x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddldependantDispensaryDistrict"]')
                    drop_down = Select(x)
                    if len(drop_down.options) > 1:
                        for x in drop_down.options:
                            if x.text == disp_district:
                                drop_down.select_by_visible_text(family_district)
                                try_again = False
                                break
                    time.sleep(2)

                # Family Dispensary
                fail_pt = "Family Dispensary"
                try_again = True
                while try_again is True:
                    x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddldependantdispensary"]')
                    drop_down = Select(x)
                    if len(drop_down.options) > 1:
                        for x in drop_down.options:
                            if x.text == dispensary:
                                drop_down.select_by_visible_text(family_dispensary)
                                try_again = False
                                break
                    time.sleep(2)

                # checking for address autofill
                try_again = True
                while try_again is True:
                    if driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_txtdependantdispaddress"]').text:
                        try_again = False
                    time.sleep(2)

                # doa
                fail_pt = "DOA- " + doa
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlDIDateOfAppointmentDy"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '//*[@id="ctl00_HomePageContent_cEDOA_title"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '//*[@id="ctl00_HomePageContent_cEDOA_title"]').click()

                time.sleep(2)
                for i in range(10):  # 100 years check
                    title_calender = driver.find_element(By.XPATH,
                                                         '//*[@id="ctl00_HomePageContent_cEDOA_title"]').text

                    # 2020-2029
                    y1 = int(title_calender[0:4])
                    y2 = int(title_calender[5:9])
                    doa_year_yyyy = int(doa[6:10])

                    if y1 <= doa_year_yyyy <= y2:
                        doa_year_last = doa_year_yyyy % y1
                        doa_year_y = str(doa_year_last)
                        # year
                        # xxxx xxx0 xxx1 xxx2
                        # xxx3 xxx4 xxx5 xxx6
                        # xxx7 xxx8 xxx9 xxxx
                        #
                        year_dic = {"0": "0_1", "1": "0_2", "2": "0_3",
                                    "3": "1_0", "4": "1_1", "5": "1_2", "6": "1_3",
                                    "7": "2_0", "8": "2_1", "9": "2_2"}

                        y = year_dic[doa_year_y]
                        year_xpath = f'//*[@id="ctl00_HomePageContent_cEDOA_year_{y}"]'
                        driver.find_element(By.XPATH, year_xpath).click()
                        time.sleep(2)
                        break
                    else:
                        driver.find_element(By.XPATH,
                                            '//*[@id="ctl00_HomePageContent_cEDOA_prevArrow"]').click()
                        time.sleep(2)

                # month
                # 3 * 4 array
                # Jan Feb Mar Apr
                # May Jun Jul Aug
                # Sep Oct Nov Dec
                month_dic = {"01": "0_0", "02": "0_1", "03": "0_2", "04": "0_3",
                             "05": "1_0", "06": "1_1", "07": "1_2", "08": "1_3",
                             "09": "2_0", "10": "2_1", "11": "2_2", "12": "2_3"}

                doa_month = doa[3:5]
                m = month_dic[doa_month]
                month_xpath = f'//*[@id="ctl00_HomePageContent_cEDOA_month_{m}"]'
                driver.find_element(By.XPATH, month_xpath).click()
                time.sleep(5)

                # date
                doa_date_dd = doa[0:2]

                # Checking if the date is "0x", then convert it to 'x'
                if doa_date_dd[0] == '0':
                    doa_date = doa_date_dd[1]
                else:
                    doa_date = doa_date_dd

                dates = driver.find_elements(By.XPATH, '//*[@class="ajax__calendar_day"]')

                # finding 1 thereby all the removing previous months dates
                i = 0
                for date in dates:
                    if date.text == '1':
                        break
                    i += 1
                dates = dates[i:-1]

                # find the date matching doa
                for date in dates:
                    if date.text == doa_date:
                        date.click()
                        break

                # click to enter nominee details
                fail_pt = "Nominee"
                driver.find_element(By.XPATH, '//tr[@id="Tr11"]//a[contains(text(),"Enter Details Here")]').click()

                # switching to new window_ nominee
                driver.switch_to.window(driver.window_handles[2])

                # nominee name
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlTextUserName"]')))
                element.send_keys(nominee_name)

                # relationship
                x = driver.find_element(By.XPATH, '//select[@id="ctl00_HomePageContent_RelationShipWithIp"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(relationship)

                # Address
                driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_ctrlTextAddress1"]').send_keys(
                    address1)

                # State drop down
                x = driver.find_element(By.XPATH, '//select[@id="ctl00_HomePageContent_States"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(state)

                # District drop down
                x = driver.find_element(By.XPATH, '//select[@id="ctl00_HomePageContent_Districts"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(district)

                # click save
                driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_Save"]').click()

                # click close
                driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_btnClose"]').click()

                # switching back to new registration page
                driver.switch_to.window(driver.window_handles[1])

                # # click to enter family particular details
                # driver.find_element(By.XPATH, '//tr[@id="Tr12"]//a[contains(text(),"Enter")]').click()
                #
                # # switching to new window_ family details
                # driver.switch_to.window(driver.window_handles[2])
                #
                # # TODO need to figure out a logic to automate this for multiple family member
                # # name
                # element = WebDriverWait(driver, 30).until(
                #     EC.presence_of_element_located((By.XPATH, '//input[@id="ctl00_HomePageContent_txtName"]')))
                # element.send_keys(dependent_name)
                #
                # # click submit
                # driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_ctrlButtonSave"]').click()
                #
                # time.sleep(5)
                #
                # # click close
                # driver.find_element(By.XPATH, '//input[@name="close_btn"]').click()

                # # switching back to new registration page
                # driver.switch_to.window(driver.window_handles[1])

                # click to enter bank details
                fail_pt = "Bank Details"
                driver.find_element(By.XPATH, '//tr[@id="Tr18"]//a[contains(text(),"Enter Details Here")]').click()

                # switching to new window_ bank details
                driver.switch_to.window(driver.window_handles[2])

                # ifsc code
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@id="ctl00_HomePageContent_txtIFSCcode"]')))
                element.send_keys(ifsc_code)

                # click search
                driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_btnIFSCcode"]').click()

                # Enter account number
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HomePageContent_txtacc_number"]')))
                element.send_keys(account_number)

                # Select account type
                x = driver.find_element(By.XPATH, '//select[@id="ctl00_HomePageContent_ddlAccountType"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(account_type)

                # click submit
                driver.find_element(By.XPATH, '//input[@id="ctl00_HomePageContent_btnsubmit"]').click()

                time.sleep(5)

                # click close
                driver.find_element(By.XPATH, '//input[@id="btnCancel"]').click()

                # switching back to new registration page
                driver.switch_to.window(driver.window_handles[1])

                # Select preferred language
                x = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_ddllanguage"]')
                drop_down = Select(x)
                drop_down.select_by_visible_text(preferred_lan)

                # click Agree
                fail_pt = "Agree checkbox"
                driver.find_element(By.XPATH, '// input[ @ id = "ctl00_HomePageContent_dec_chkbox"]').click()

                # Click Submit #
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_Submit"]').click()

                # # Click Cancel used for debugging only
                # driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_btncancel"]').click()

                # get insurance number
                # time.sleep(20) - commenting this for now
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HomePageContent_ctrlLabelIPNumber"]')))
                insurance_number = element.text

                # add insurance number into the dataframe
                data.loc[row, "New ESIC number"] = insurance_number

                # click close
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_close"]').click()

                print(f"Completed entering {name}- {insurance_number}")

            # TODO if esic number already exists
            # Employee with previous employee number
            else:
                fail_pt = "Previous ESIC number"
                # Enter prev insurance number
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '// *[ @ id = "ctl00_HomePageContent_ctrlTxtIPNumber"]')))
                element.send_keys(prev_esic_num)

                # doa for employee with previous insurance number
                # TODO verify doa
                fail_pt = "DOA- " + doa
                driver.find_element(By.XPATH, '// *[ @ id = "ctl00_HomePageContent_ctrlTxtAppointmentDate"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_cEDOA_title"]').click()

                time.sleep(2)
                driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_cEDOA_title"]').click()

                time.sleep(2)
                for i in range(10):  # 100 years check
                    title_calender = driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_cEDOA_title"]').text

                    # 2020-2029
                    y1 = int(title_calender[0:4])
                    y2 = int(title_calender[5:9])
                    doa_year_yyyy = int(doa[6:10])

                    if y1 <= doa_year_yyyy <= y2:
                        doa_year_last = doa_year_yyyy % y1
                        doa_year_y = str(doa_year_last)
                        # year
                        # xxxx xxx0 xxx1 xxx2
                        # xxx3 xxx4 xxx5 xxx6
                        # xxx7 xxx8 xxx9 xxxx
                        #
                        year_dic = {"0": "0_1", "1": "0_2", "2": "0_3",
                                    "3": "1_0", "4": "1_1", "5": "1_2", "6": "1_3",
                                    "7": "2_0", "8": "2_1", "9": "2_2"}

                        y = year_dic[doa_year_y]
                        year_xpath = f'//*[@id="ctl00_HomePageContent_cEDOA_year_{y}"]'
                        driver.find_element(By.XPATH, year_xpath).click()
                        time.sleep(2)
                        break
                    else:
                        driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_cEDOA_prevArrow"]').click()
                        time.sleep(2)

                # month
                # 3 * 4 array
                # Jan Feb Mar Apr
                # May Jun Jul Aug
                # Sep Oct Nov Dec
                month_dic = {"01": "0_0", "02": "0_1", "03": "0_2", "04": "0_3",
                             "05": "1_0", "06": "1_1", "07": "1_2", "08": "1_3",
                             "09": "2_0", "10": "2_1", "11": "2_2", "12": "2_3"}

                doa_month = doa[3:5]
                m = month_dic[doa_month]
                month_xpath = f'//*[@id="ctl00_HomePageContent_cEDOA_month_{m}"]'
                driver.find_element(By.XPATH, month_xpath).click()
                time.sleep(5)

                # date
                doa_date_dd = doa[0:2]

                # Checking if the date is "0x", then convert it to 'x'
                if doa_date_dd[0] == '0':
                    doa_date = doa_date_dd[1]
                else:
                    doa_date = doa_date_dd

                dates = driver.find_elements(By.XPATH, '//*[@class="ajax__calendar_day"]')

                # finding 1 thereby all the removing previous months dates
                i = 0
                for date in dates:
                    if date.text == '1':
                        break
                    i += 1
                dates = dates[i:-1]

                # find the date matching doa
                for date in dates:
                    if date.text == doa_date:
                        date.click()
                        break

                # Click continue
                driver.find_element(By.XPATH, '// *[ @ id = "ctl00_HomePageContent_btnContinue"]').click()

                # CLick Update
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '// *[ @ id = "ctl00_HomePageContent_ctrlButtonSave"]')))
                element.click()

                # Click Close
                # TODO where is this button, next page or the same page
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '// *[ @ id = "ctl00_HomePageContent_btnGo_Success"]')))
                element.click()

                # Close window and Switch to main window so that new registration link can be clicked
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

                data.loc[row, "New ESIC number"] = "-"

                print(f"Updated {name}")

            success += 1

        except:
            driver.find_element(By.XPATH, '//*[@id="ctl00_HomePageContent_btncancel"]').click()

            print(f"[ERROR] Unable to complete registration of {name} when trying to enter {fail_pt}")
            fail += 1
            data.loc[row, "New ESIC number"] = 'Problem with' + fail_pt

    data.to_excel(xlsx_filename)

    print(f'Total execution time:{(time.time() - start_time) / 60} minutes ')
    print(f"{success}/{no_rows} entries completed successfully")

    driver.quit()


# User inputs
uname = input("Enter the User name:")
passwd = input("Enter password:")
filename = input("Enter the excel filename:")

if filename.find('.xlsx') == -1:  # adding extension if not provided by the user
    filename = filename + '.xlsx'

try:
    esic_registration(username=uname, password=passwd, xlsx_filename=filename)

# if the file not found
except FileNotFoundError:
    print("[Error] File not found")