from ast import Return
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date
import logging
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import sharepy
import os
from bu_config import get_config
import bu_alerts


today_date=date.today()
# log progress --
logfile = os.getcwd() + "\\Logs\\" +'O_REILLY_INVOICE_AUTOMATION_Logfile'+str(today_date)+'.txt'

logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=logfile)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logging.info('setting paTH TO DOWNLOAD')
path = os.getcwd() + "\\Download"

logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')


profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.dir', path)
profile.set_preference('browser.download.useDownloadDir', True)
profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
profile.set_preference('pdfjs.disabled', True)
logging.info('Adding firefox profile')
driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)

site = 'https://biourja.sharepoint.com'
path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
path2= "Shared%20Documents/Power%20Reference/Power_Invoices/O_Reilly/"
# path2= "Shared Documents/Vendor Research/Enverus(PRT)/ERCOT"

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

def download_wait(path_to_downloads= os.getcwd() + '\\Download'):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = True
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.pdf'):
                dl_wait = False
        seconds += 1
    time.sleep(seconds)
    driver.quit()
    return seconds

def login_and_download():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        driver.get("https://www.oreilly.com/member/login")
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Email Address']"))).send_keys(username)
        time.sleep(1)
        logging.info('Continue Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='orm-Button-root']"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Password']"))).send_keys(password)
        time.sleep(1)
        logging.info('click on Sign In Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".orm-Button-root.FullWidth--FD7XS"))).click()
        time.sleep(5)
        logging.info('click on Settings Tab')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/u/preferences/']"))).click()
        time.sleep(5)
        logging.info('Accessing plans and payment')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT, "Plans & Payment"))).click()
        time.sleep(5)
        logging.info('clicking on Billing History')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT, "See Billing History"))).click()
        time.sleep(5)
        logging.info('clicking on Download Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//ul[1]//li[1]//div[4]//button[1]"))).click()
        time.sleep(10)
        logging.info('*****************Download Successfully*************')
        logging.info(f'File is downloaded in {download_wait()} seconds.')
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
        filesToUpload = os.listdir(os.getcwd() + "\\Download")
        for fileToUpload in filesToUpload:
                
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "application/pdf"}

            with open(os.path.join(os.getcwd() + "\\Download", f'{fileToUpload}'), 'rb') as read_file:
                    content = read_file.read()
            fileToUpload=fileToUpload.replace("'","_")     
            p = s.post(f"{site}{path1}('{path2}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
                # url = f"https://biourja.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('Shared Documents/Vendor Research/Enverus(PRT)/PJMISO')/Files/add(url='dummy.pdf',overwrite=true)"
                # r = s.post(url.format("C:/Users/Yashn.jain/Desktop/First_Project", "Enverus_PJM 90 Price Forecast 02-02-22T.pdf"), data=content, headers=headers)
        print(f'{fileToUpload} uploaded successfully')
    
        print(f'{job_name} executed succesfully')
        return p   
        
    except Exception as e:
        raise e

def main():
    try:
        remove_existing_files(files_location)
        login_and_download()
        s=connect_to_sharepoint()
        shp_file_upload(s)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully, Attached logs',attachment_location = logfile)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
            
    
if __name__ == "__main__":
    logging.info("Execution Started")
    time_start=time.time()
    directories_created=["Download","Logs"]
    for directory in directories_created:
        path = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
        except OSError as error:
            print("Directory '%s' can not be created" % directory)

    files_location=os.getcwd() + "\\Download"
    Project_name="OREILLY_INVOICE_AUTOMATION"
    Table_name="OREILLY_INVOICE_AUTOMATION"
    credential_dict = get_config('OREILLY_INVOICE_AUTOMATION','OREILLY_INVOICE_AUTOMATION')
    username = credential_dict['USERNAME'].split(';')[0]
    password = credential_dict['PASSWORD'].split(';')[0]
    sp_username = credential_dict['USERNAME'].split(';')[1]
    sp_password =  credential_dict['PASSWORD'].split(';')[1]
    receiver_email = credential_dict['EMAIL_LIST']
    job_name='OREILLY_INVOICE_AUTOMATION'
   
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')       