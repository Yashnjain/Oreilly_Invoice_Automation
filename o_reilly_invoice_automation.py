from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date,datetime
import logging
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import sharepy
import os
from bu_config import get_config
import bu_alerts
import numpy as np


def remove_existing_files(files_location):
    logger.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logger.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logger.info(e)
        raise e


def login_and_download():  
    '''This function downloads log in to the website'''
    try:
        driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)
        logging.info('Accesing website')
        driver.get("https://www.oreilly.com/member/login")
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/section/div[1]/section/section/form/div/label/input"))).send_keys(username)
        time.sleep(1)
        logging.info('Continue Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='orm-Button-root']"))).click()
        time.sleep(1)
        WebDriverWait(driver, 9, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/main/section/div[1]/section/section/section/form/div[1]/div/label/input'))).send_keys(password)        
        time.sleep(1)
        logging.info('click on Sign In Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".orm-Button-root.FullWidth--FD7XS"))).click()
        time.sleep(5)
        logging.info('click on Settings Tab')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/u/preferences/']"))).click()
        time.sleep(5)
        logging.info('Accessing plans and payment')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/main/section/nav/ul/li[5]/a"))).click()
        time.sleep(10)
        logging.info('clicking on Billing History')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/main/section/article/div/div/div/div[3]/div/a[2]"))).click()        
        time.sleep(10)
        logging.info('clicking on Download Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//ul[1]//li[1]//div[4]//button[1]"))).click()
        time.sleep(60)
        logging.info('*****************Download Successfully*************')
        try:
            driver.close()
        except Exception as e: 
            logging.info('driver not closed')
            print("driver not closed") 
            try:
                driver.quit()
            except Exception as e: 
                logging.info('driver quit failed')
                print("driver quit failed") 
    except Exception as e:
        raise e   

def connect_to_sharepoint():
    logging.info('Connecting to sharepoint')
    try:
        site='https://biourja.sharepoint.com'
        username = os.getenv("user") if os.getenv("user") else sp_username
        password = os.getenv("password") if os.getenv("password") else sp_password
        # Connecting to Sharepoint and downloading the file with sync params
        s = sharepy.connect(site, username, password)
        return s
    except Exception as e:
        raise e

def shp_file_upload(s):
    logging.info('Uploading files to sharepoint')
    try:
        global body
        body = ''
        filesToUpload = os.listdir(os.getcwd() + "\\download")
        for fileToUpload in filesToUpload:
            z=path+'\\'+fileToUpload
            locations_list.append(z)     
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "application/pdf"}

            with open(os.path.join(os.getcwd() + "\\download", f'{fileToUpload}'), 'rb') as read_file:
                    content = read_file.read()
            fileToUpload=fileToUpload.replace("'","_")     
            p = s.post(f"{site}{path1}('{path2}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
            nl = '<br>'
            body += (f'{fileToUpload} successfully uploaded, {nl} Attached link for the same:-{nl}{temp_path}{nl}')

        print(f'{fileToUpload} uploaded successfully')
    
        print(f'{job_name} executed succesfully')
        return p   
        
    except Exception as e:
        raise e

def main():
    try:
        no_of_rows=0
        Database=""
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name=processname,database=Database,status='Started',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        remove_existing_files(files_location)
        login_and_download()
        s=connect_to_sharepoint()
        shp_file_upload(s)
        locations_list.append(logfile)
        bu_alerts.bulog(process_name=processname,database=Database,status='Completed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)   
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',multiple_attachment_list=locations_list)
    except Exception as e:
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=Database,status='Failed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
                
if __name__ == "__main__":
    try:
        logging.info("Execution Started")
        time_start=time.time()
        #Global VARIABLES
        site = 'https://biourja.sharepoint.com'
        path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
        path2= "Shared%20Documents/Power%20Reference/Power_Invoices/O_Reilly/"
        # path2= "Shared Documents/Vendor Research/Enverus(PRT)/ERCOT"
        temp_path='https://biourja.sharepoint.com/BiourjaPower/Shared%20Documents/Power%20Reference/Power_Invoices/O_Reilly'
        locations_list=[]
        body = ''

        today_date=date.today()
        # log progress --
        for handler in logging.root.handlers[:]:
           logging.root.removeHandler(handler)
        logfile = os.getcwd() + "\\logs\\" +'O_REILLY_INVOICE_AUTOMATION_Logfile'+str(today_date)+'.txt'

        logging.basicConfig(filename=logfile, filemode='w',
                            format='%(asctime)s %(message)s')
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info('setting paTH TO DOWNLOAD')
        path = os.getcwd() + "\\download"

        logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')
        profile = webdriver.FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.dir', path)
        profile.set_preference('browser.download.useDownloadDir', True)
        profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
        profile.set_preference('pdfjs.disabled', True)
        logging.info('Adding firefox profile')

     
        directories_created=["download","logs"]
        for directory in directories_created:
            path3 = os.path.join(os.getcwd(),directory)  
            try:
                os.makedirs(path3, exist_ok = True)
                print("Directory '%s' created successfully" % directory)
            except OSError as error:
                print("Directory '%s' can not be created" % directory)

        files_location=os.getcwd() + "\\download"
        project_name="OREILLY_INVOICE_AUTOMATION"
        table_name="OREILLY_INVOICE_AUTOMATION"
        credential_dict = get_config('OREILLY_INVOICE_AUTOMATION','OREILLY_INVOICE_AUTOMATION')
        username = credential_dict['USERNAME'].split(';')[0]
        password = credential_dict['PASSWORD'].split(';')[0]
        sp_username = credential_dict['USERNAME'].split(';')[1]
        sp_password =  credential_dict['PASSWORD'].split(';')[1]
        receiver_email = credential_dict['EMAIL_LIST']
        # receiver_email ='yashn.jain@biourja.com'
        job_name=credential_dict['PROJECT_NAME']
        job_id=np.random.randint(1000000,9999999)
        processname = credential_dict['PROJECT_NAME']
        process_owner = credential_dict['IT_OWNER']

    
        main()
        time_end=time.time()
        logging.info(f'It takes {time_start-time_end} seconds to run') 
    except Exception as e:
            logging.exception(str(e))
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)
        

          