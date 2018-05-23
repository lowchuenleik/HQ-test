from PIL import Image
from PIL import ImageGrab
from selenium import webdriver
import pytesseract
import cv2
from googleapiclient.discovery import build
import pprint
import win32gui
import win32process
import os
import time
from ctypes import windll #DPI SF Correction
# Make program aware of DPI scaling
user32 = windll.user32
user32.SetProcessDPIAware()
###

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from multiprocessing import Process
import psutil
import random
import win32gui
import win32process

v1 = 'v1.png'
ques = 'qs.png'
PHONE_NAME = 'SM G935F'

windows = []

my_api_key = 'AIzaSyDko0nyDdbIkBgnFQChTvBl_97wdRlcdCQ'
my_cse_id = '005645035030761376503:shiwo7rk4sq'
'''
# A list of processes with a name Vysor.exe:
Vysor = [item for item in psutil.process_iter() if item.name() == 'Vysor.exe']
print(Vysor)  # [<psutil.Process(pid=4416, name='Vysor.exe') at 64362512>]

# pid of the first found Vysor.exe process:
pid = next(item for item in psutil.process_iter() if item.name() == 'Vysor.exe').pid
# (raises StopIteration exception if Vysor is not running)

print(pid)



def enum_window_callback(hwnd, pid):
    tid, current_pid = win32process.GetWindowThreadProcessId(hwnd)
    if pid == current_pid and win32gui.IsWindowVisible(hwnd):
        windows.append(hwnd)

windows = [] #Windows now contains all HWND of Vysor

win32gui.EnumWindows(enum_window_callback, pid)
'''

import win32gui

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst
'''
appwindows = get_app_list()
for i in appwindows:
    print(i)
    if PHONE_NAME in i:
    	windows.append(i[0])

driver = webdriver.Chrome("C:/Users/Ley/Desktop/HQ/chromedriver.exe")
driver.get("https://www.google.com")

pytesseract.pytesseract.tesseract_cmd = "C:\\Users\\Ley\\AppData\\Local\\Tesseract-OCR\\tesseract"
'''

def catch():

	start = time.time()

	win32gui.SetForegroundWindow(windows[0])

	box = win32gui.GetWindowRect(windows[0])

	#CORRECTION factor
	a = box[0] + 40 #LEFT X
	b = box[1] + 430 #TOP Y
	c = box[2] - 40 #RIGHT X
	d = box[3] - 1210 #BOTTOM Y

	boxup = (a,b,c,d) #TUPLES DONT SUPPORT ITEM ASSIGNMENT

	img = ImageGrab.grab(bbox=boxup)

	a = box[0] + 40 #LEFT X
	b = box[1] + 930 #TOP Y
	c = box[2] - 40 #RIGHT X
	d = box[3] - 700 #BOTTOM Y

	qbox = (a,b,c,d)

	questions = ImageGrab.grab(bbox=qbox)

	img.save(str(random.randint(0,10000)) + v1)
	questions.save(str(random.randint(0,10000)) + ques)


	'''
	#Following probably unnceessary if speed required.

	#IMPORTANT, HAVE TO SAVE IT FOR OPENCV2 TO READ
	image = cv2.imread(v1)
	qs = cv2.imread(questions)
	#os.remove(v1)
	grayq = cv2.cvtColor(qs, cv2.COLOR_BGR2GRAY)
	gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
	#gray = cv2.medianBlur(gray, 3)

	filename = "{}.png".format(os.getpid())
	cv2.imwrite(filename, gray)

	filenameq = = "{}.png".format(os.getpid())
	cv2.imwrite(filenameq, grayq)

	text = pytesseract.image_to_string(Image.open(filename)) #An array of things
	questions = pytesseract.image_to_string(Image.open(filenameq))

	qna = text.split('\n\n')

	for x in qna:
		x = x.replace('\n',' ')

	question = qna[0].replace('\n',' ')

	print(question)
	print(qna)

	os.remove(filename)
	os.remove(filenameq)
	'''
	question = pytesseract.image_to_string(img)
	question = question.replace('\n',' ')

	answers = pytesseract.image_to_string(questions)
	answers = answers.split('\n\n')

	print(question,answers)

	WebDriverWait(driver, 120).until(
        EC.element_to_be_clickable(
            (By.ID, "lst-ib")
        )
    )

	
	analyse_count(question,answers)#Approach 2
	indiv_search(question,answers) #Approach 3
	#search(question) #Appoach 1

	print('\nTook this many seconds:',time.time()-start)

	#driver.execute_script('''window.open("http://bings.com","_blank");''')
	'''
	driver.switch_to.window(driver.window_handles[-1])
	WebDriverWait(driver, 120).until(
        EC.element_to_be_clickable(
            (By.ID, "sb_form_q")
        )
    )
	elem = driver.find_element_by_id("sb_form_q")
	elem.send_keys(text.replace("\n"," "))
	elem.send_keys(Keys.RETURN)

	time.sleep(0.1)

	driver.switch_to.window(driver.window_handles[0])

	lis = [qna[-1],qna[-2],qna[-3]]
	for x in lis:
		x = question + x
		approach2(x)
	'''
def indiv_search(query,answers):
	for ans in answers:
			results,total = google_search(query + ' ' + ans, my_api_key, my_cse_id, num=10)
			print('This many results:',total,' for ',query+' '+ans,'\n')


def search(query):
	driver.switch_to.window(driver.window_handles[0])
	elem = driver.find_element_by_id("lst-ib")
	elem.clear()
	elem.send_keys(query)
	elem.send_keys(Keys.RETURN)
	driver.switch_to.window(driver.window_handles[0])

def approach2(query):
	driver.execute_script('''window.open("http://google.com","_blank");''')
	driver.switch_to.window(driver.window_handles[-1])
	elem = driver.find_element_by_id("lst-ib")
	elem.send_keys(query)
	elem.send_keys(Keys.RETURN)

	time.sleep(0.1)

	elem = driver.find_element_by_id("resultStats")
	stats = int(elem.text.split()[1].replace(',',''))
	print(stats,'given',query)
	driver.close()

	return stats


def google_search(search_term, api_key, cse_id, **kwargs):
    service = build("customsearch", "v1", developerKey=api_key)
    res = service.cse().list(q=search_term, cx=cse_id, **kwargs).execute()
    tot = res['queries']['request'][0]['totalResults']
    return res['items'], int(tot)

def analyse_count(query,answer):
	results,total = google_search(query, my_api_key, my_cse_id, num=10)
	#results is a list of all entries
	text_alys = ''
	for result in results:
		text_alys += result['title'].replace('\n',' ') + result['snippet'].replace('\n',' ')
	text_alys = text_alys.upper()
	#print(text_alys)
	for ans in answer:
		count = text_alys.count(ans.upper())
		print('\nObtained ',count,' by searching ',ans,' for ',query,'\n')
	return count


