#execute "pip install python-docx PyAutoGUI pynput autopy" if you face any issues in dependencies 
from docx import Document
from docx.shared import Inches
import pyautogui
import os
from pynput.keyboard import Key, Listener
import datetime

currentDir = os.getcwd() 
bad_chars = [';', ':', '!', "*","<",">","\"","/","|","?","*"]
document = Document()

exit_combination_msg='ctrl_l(left ctrl) + Esc'
#exit conditions : Ctrl_l + Esc Key
exit_combination = {Key.ctrl_l, Key.esc}

currently_pressed = set()

# add screen shot to document
def addSSToDoc(imgPath):
  p = document.add_paragraph()
  r = p.add_run()
  r.add_picture(imgPath,width=Inches(5.90551), height=Inches(3.54331))
  p = document.add_paragraph()

#save document
def saveDoc(fileName):
  document.save(fileName+'.docx')
  print()
  print("File Saved as "+fileName+".docx")

#on exit function
def exitFun():
  now = datetime.datetime.now()
  timenow = now.strftime("%H_%M_%S")
  fileName = input("Enter File Name : ")
  for i in bad_chars : 
    fileName = fileName.replace(i, '')

  if(fileName==''):
    fileName = 'testingScreenShots'
  saveDoc(fileName+"_"+timenow)



#on press function
def on_press(key):
    check_key(key)
    if key in exit_combination:
        currently_pressed.add(key)
        if currently_pressed == exit_combination:
        	listener.stop()
        	exitFun()
        	
        

#on release function
def on_release(key):  
    try:
        currently_pressed.remove(key)
    except KeyError:
        pass 
   
#save image as file and add image to doc
def saveImg():
	shot = pyautogui.screenshot()
	now = datetime.datetime.now()
	timenow = now.strftime("%H_%M_%S.%f")[:-3]
	path = currentDir + "\\shots\\"+str(datetime.date.today())
	try:
		shot.save(path+'\\'+timenow+'.png')
		print('File Saved as ' +path+'\\'+timenow+'.png')
		addSSToDoc(path+'//'+timenow+'.png')
	except FileNotFoundError:  
		os.makedirs(path)
		shot.save(path+'//'+timenow+'.png')
		addSSToDoc(path+'//'+timenow+'.png')

# if key is ptrscn
def check_key(key):
    if key == Key.print_screen:
        #print("pressed")
        saveImg()

# exit_combination = {Key.ctrl_l, Key.esc}
currently_pressed = set()

# Collect events until released
with Listener(on_press=on_press,on_release=on_release) as listener:
  print("Created by Chetan Rakheja")
  print("Press "+exit_combination_msg +" to exit and save the document")
  listener.join()

