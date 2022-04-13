from googleapiclient.errors import HttpError as HTTPError
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload
import io
import requests
import time
import timeit
import os
import glob
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select

from csv import writer
from bs4 import BeautifulSoup
from urllib3.util import Retry
from requests.adapters import HTTPAdapter
from datetime import datetime as dt
# from pprint import pprint

###Utility###
current = os.path.abspath(os.path.dirname(__file__))


def currentDir(fname: str):
    asincurrentDir = os.path.join(current, fname)
    return asincurrentDir


today = dt.today().strftime('%Y%m%d')

###Selenium Work##


def downloadFromAdmin():
    download_dir = current

    '''Login Detail'''
    loginpage = "https://www.rekotrade.com/dms/auth"
    loginID = os.environ["autoxloo-admin-id-relation"]
    loginPW = os.environ["autoxloo-admin-pw-relation"]

    # All the stocks are listed here
    inventoryPage = 'https://www.rekotrade.com/dms/inventory/inventory_reports'
    '''Driver setup'''
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
                                    "download.default_directory": download_dir,
                                    "download.prompt_for_download": False,
                                    "download.directory_upgrade": True,
                                    "plugins.plugins_disabled": ["Chrome PDF Viewer"],
                                    "plugins.always_open_pdf_externally": True,
                                    "profile.default_content_setting_values.automatic_downloads": 1
                                    })

    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument("--disable-extensions")
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    driver.command_executor._commands["send_command"] = (
        "POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {
        'behavior': 'allow', 'downloadPath': download_dir}}
    command_result = driver.execute("send_command", params)

    fById = driver.find_element_by_id
    fByName = driver.find_element_by_name
    fByClass = driver.find_element_by_class_name

    driver.implicitly_wait(10)

    '''Go to Login page '''
    driver.get(loginpage)

    loginElm = fById("login")
    loginElm.clear()
    loginElm.send_keys(loginID)

    pwElm = fById("password")
    pwElm.clear()
    pwElm.send_keys(loginPW)

    signInButton = fById("login2")
    signInButton.click()

    '''Go to Inventory '''
    driver.get(inventoryPage)

    '''Select "All Data" template'''
    tempElm = fByClass("temp-select")
    Select(tempElm).select_by_visible_text("All Data")
    sleep(1)

    '''Click "Download XLS'''
    downloadButton = fByName("print_xls")
    downloadButton.click()
    sleep(3)
    f_list = []
    while len(f_list) == 0:
        sleep(1)
        f_list = glob.glob(currentDir("_home_utc10xloo_www*.xls"))
    driver.close()


### Collaboration with Martin's beforward.jp Scraping ###
s = requests.Session()
headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36"}
all_urls = "https://www.beforward.jp/stocklist/client_wishes_id=/description=/make=/model=/fuel=/fob_price_from=/fob_price_to=/veh_type=/steering=/mission=/mfg_year_from=/mfg_year_to=/mileage_from=/mileage_to=/cc_from=/cc_to=/showmore=/drive_type=/color=/stock_country=35/area=/seats_from=/seats_to=/max_load_min=/max_load_max=/veh_type_sub=/view_cnt=2000/page=1/sortkey=n/sar=/from_stocklist=1/keyword=/kmode=and/"
baseurl = "https://www.beforward.jp"
listofurls = []

'''Retry settings for requests'''
retries = Retry(total=5,  # リトライ回数
                backoff_factor=2,  # sleep時間
                status_forcelist=[500, 502, 503, 504, 429])
s.mount("https://", HTTPAdapter(max_retries=retries))


''' Building Drive API Instance '''
json_key_file = currentDir("relation-gdrive-handler-for-scraping.json")

SCOPES = ['https://www.googleapis.com/auth/drive']
SHARE_FOLDER_ID = 'FOLDERID'

sa_creds = service_account.Credentials.from_service_account_file(
    json_key_file)
scoped_creds = sa_creds.with_scopes(SCOPES)
drive_service = build('drive', 'v3', credentials=scoped_creds)

''' Today's Folder '''
file_metadata = {
    'name': today,
    'mimeType': 'application/vnd.google-apps.folder',
    'parents': [SHARE_FOLDER_ID]
}
response = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
folderID = response.get("id")

'''Upload Function Building'''


def uploadCsvToGdriveFromIO(fname: str, buffer: io.StringIO, excel=False):
    file_metadata = {'name': fname, 'parents': [folderID]}
    media = MediaIoBaseUpload(buffer,
                              mimetype='text/csv') if excel == False else MediaIoBaseUpload(buffer,
                                                                                            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()


def uploadCsvToGdriveFromFile(fname: str, fPath: str, excel=False):
    file_metadata = {'name': fname, 'parents': [folderID]}
    media = MediaFileUpload(fPath,
                            mimetype='text/csv') if excel == False else MediaFileUpload(fPath,
                                                                                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()


''' Start Scraping '''


def getAllCars():
    pagecontent = s.get(all_urls, headers=headers).content
    soup = BeautifulSoup(pagecontent, "html.parser")
    for item in soup.find_all("a", {"class": "vehicle-url-link"}):
        new_url = baseurl+item['href']
        if new_url not in listofurls:
            listofurls.append(new_url)
    with open("listofurls", "w") as f:
        for item in listofurls:
            f.write("%s\n" % item)


def getCarInfo(url):
    carinfo = []
    pagecontent = s.get(url, headers=headers).content
    soup = BeautifulSoup(pagecontent, "html.parser")
    try:
        if soup.find("div", {"class": "list-detail-box-underoffer"}) == None and soup.find("p", {"class": "sold-text"}) == None:
            print("test")
            table = soup.find("table", {"class": "specification"})
            # Re-call BS with new html
            soup1 = BeautifulSoup(str(table), "html.parser")
            for table_item in soup1.find_all("td"):
                temp = table_item.renderContents().decode("utf-8")
                finalstring = temp.strip("\n").strip(
                    "\t").strip("\r").strip("\t").strip("\n")
                carinfo.append(gg)
            price = soup.find(
                "span", {"class": "price ip-usd-price"}).renderContents().decode("utf-8")
            listingname = soup.find(
                "div", {"class": "car-info-flex-box"}).h1.renderContents().decode("utf-8")
            carinfo.append(price)
            carinfo.append(listingname)
            carinfo.append(url)
            with open('beforward.csv', 'a') as fd:
                writer_object = writer(fd)
                writer_object.writerow(carinfo)
                fd.close()
        else:
            pass
    except:
        print(soup)
        print(pagecontent)

### Main Controller ###


def main():
    downloadFromAdmin()
    downloadedFile = glob.glob(currentDir("_home_utc10xloo_www*.xls"))[0]
    uploadCsvToGdriveFromFile("inventory.xls", downloadedFile, excel=True)
    getAllCars()
    with open("listofurls") as file:
        for line in file:
            getCarInfo(line.rstrip())
            print(line.rstrip())
            time.sleep(1)
    uploadCsvToGdriveFromFile("beforward.csv", currentDir("beforward.csv"))


if __name__ == '__main__':
    result = timeit.timeit("main()", globals=globals(), number=1)
    print(result)
