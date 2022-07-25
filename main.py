from PIL import ImageGrab, Image
import time
import docx
from docx.shared import Cm
import os
import sys
#from pywinauto import Application
#import pywinauto.keyboard as keyboard
import subprocess


def list_pages_win():
    '''
    app = Application().connect(process=2341)
    dlg = app['Untitled - Notepad']
    dlg.set_focus()
    keyboard.send_keys('{LEFT}')
    '''
    p = subprocess.Popen('powershell.exe -ExecutionPolicy RemoteSigned -file "script.ps1"')
    p.communicate()


def list_pages_linux():
    os.system('./list_pages')


def grab_and_save(filename):
    im = ImageGrab.grab(bbox=(320, 60,  1380, 900))
    #im = ImageGrab.grab()
    out = im.transpose(Image.Transpose.ROTATE_270)

    out.save(filename)


def get_file_list():
    file_list = os.listdir(os.getcwd())
    file_list1 = file_list[:]
    for file_name in file_list1:
        if not file_name.endswith('.png'):
            file_list.remove(file_name)
    return sorted(file_list)


if __name__ == '__main__':

    if sys.platform == 'linux':
        list_pages = list_pages_linux
    elif sys.platform == 'win32':
        list_pages = list_pages_win

    time.sleep(3)
    for i in range(5):
        list_pages()
        time.sleep(1)
        grab_and_save('test{}.png'.format(i))

    doc = docx.Document()
    for file_name in get_file_list():
        doc.add_picture(file_name, width=Cm(19))
        #doc.add_page_break()
    doc.save('test.docx')

