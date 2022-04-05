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
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders

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

def send_mail(receiver_email: str, mail_subject: str, mail_body: str, attachment_locations: list = None, sender_email: str = None, sender_password: str=None) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_locations (list, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    logging.info("INTO THE SEND MAIL FUNCTION")
    done = False
    try:
        logging.info("GIVING CREDENTIALS FOR SENDING MAIL")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            sender_password = r"bY3mLSQ-\Q!9QmXJ"
            # sender_email = r"virtual-out@biourja.com"
            # sender_password = "t?%;`p39&Pv[L<6Y^cz$z2bn"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = "biourjapowerdata@biourja.com"
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        logging.info("Attaching mail body")
        msg.attach(email.mime.text.MIMEText(body, 'html'))
        logging.info("Attching files in the mail")
        for files_locations in attachment_locations:
            with open(files_locations, 'r+b') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % files_locations)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('us-smtp-outbound-1.mimecast.com',
                         587)  # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        logging.info("Email sent successfully")
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done


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
        time.sleep(5)
        logging.info('clicking on Billing History')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/main/section/article/div/div/div/div[3]/div/a[2]"))).click()        
        time.sleep(5)
        logging.info('clicking on Download Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//ul[1]//li[1]//div[4]//button[1]"))).click()
        time.sleep(10)
        logging.info('*****************Download Successfully*************')
        logging.info(f'File is downloaded in {download_wait()} seconds.')
    except Exception as e:
        raise e
    finally:
        driver.close()    

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
        filesToUpload = os.listdir(os.getcwd() + "\\Download")
        for fileToUpload in filesToUpload:
            z=path+'\\'+fileToUpload
            locations_list.append(z)     
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "application/pdf"}

            with open(os.path.join(os.getcwd() + "\\Download", f'{fileToUpload}'), 'rb') as read_file:
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
        remove_existing_files(files_location)
        login_and_download()
        s=connect_to_sharepoint()
        shp_file_upload(s)
        locations_list.append(logfile)
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',attachment_locations = locations_list)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
                
if __name__ == "__main__":
    logging.info("Execution Started")
    time_start=time.time()
    directories_created=["Download","Logs"]
    for directory in directories_created:
        path3 = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path3, exist_ok = True)
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
    # receiver_email ='yashn.jain@biourja.com'
    job_name='OREILLY_INVOICE_AUTOMATION'
   
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')       