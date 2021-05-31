import socket
import win32serviceutil
import servicemanager
import win32event
import win32service
import os
import sys
import wget
import zipfile
import easygui
import csv
import shutil
import pyautogui
import win32gui, win32con
import pandas as pd
from win32com.client import Dispatch
from csv import DictReader
from win32api import GetSystemMetrics
from selenium import webdriver
from getpass import getpass
from time import sleep
from datetime import datetime


class SMWinservice(win32serviceutil.ServiceFramework):
    _svc_name_ = 'ebagiriservice'
    _svc_display_name_ = "EBA GIRIS v0.18"
    _svc_description_ = "EBA'daki derslere otomatik olarak girmesi için hazırlanmıştır"
    @classmethod
    def parse_command_line(cls):
        win32serviceutil.HandleCommandLine(cls)
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
    def SvcStop(self):
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
    def SvcDoRun(self):
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def start(self):   
        def get_version_via_com(filename):
            parser = Dispatch("Scripting.FileSystemObject")
            try:
                version = parser.GetFileVersion(filename)
            except Exception:
                return None
            return version      
        paths = [r"C:/Program Files/Google/Chrome/Application/chrome.exe",
                r"C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"]
        version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]

        if str(version) == "90.0.4430.212" or "90.0.4430.93" or "90.0.4430.85" or "90.0.4430.72":
            googleversion = "90.0.4430.24"
        elif str(version) == "89.0.4389.128" or "89.0.4389.114" or "89.0.4389.90" or "89.0.4389.72":
            googleversion = "89.0.4389.23"

        googleversion = "https://chromedriver.storage.googleapis.com/"+googleversion+"/chromedriver_win32.zip"

        downloadfilepath = "C:/ebagiris"
        if os.path.exists(downloadfilepath):
            shutil.rmtree(downloadfilepath)
        os.mkdir(downloadfilepath)    
        sleep(0.5)
        wget.download(googleversion, downloadfilepath)   
        with zipfile.ZipFile('c:/ebagiris/chromedriver_win32.zip', 'r') as zip_ref:
            zip_ref.extractall('c:/ebagiris') 
        tckimlik = "nothing"
        sifre = "nothing"
        csvdosya = open("C:/ebagiris/config.csv","x")
        with open('C:/ebagiris/config.csv', 'w') as f:
            fieldnames = ['tckimlik', 'sifre']
            thewriter = csv.DictWriter(f, fieldnames=fieldnames)
            thewriter.writeheader()
            thewriter.writerow({'tckimlik' : tckimlik, 'sifre' : sifre})         

    def stop(self):
        shutil.rmtree(downloadfilepath)

    def main(self):
        def girisyap():
            with open('C:/ebagiris/config.csv', 'r') as reader:
                csv_dict_reader = DictReader(reader)
                for row in csv_dict_reader:
                    tckimlik1 = row['tckimlik']
                    sifre1 = row['sifre']
            driver_path = "C:/ebagiris/chromedriver"
            driver      = webdriver.Chrome(driver_path)
            driver.maximize_window()
            driver.get('https://giris.eba.gov.tr/EBA_GIRIS/student.jsp')      
            username    = driver.find_element_by_name('tckn').send_keys(tckimlik1)
            password    = driver.find_element_by_name('password').send_keys(sifre1)
            girisbutonu = driver.find_element_by_xpath('//button[@title="Giriş Butonu"]').click()
            katilbutonu = driver.find_element_by_link_text('Hemen Katıl').click()
            sleep(3)
            dersegiris  = driver.find_element_by_xpath("//a[@id='join']").click()
            yukseklik = GetSystemMetrics(0)
            if str(yukseklik) == "1920":
                sleep(63)
                pyautogui.click(1071,216)
            if str(yukseklik) == "1366":
                sleep(63)
                pyautogui.click(740, 214)
            sleep(10)
            driver.quit()

        def mickapatma():
            pyautogui.FAILSAFE = False
            sleep(0.2)
            startButton = pyautogui.locateOnScreen('C:/muteicon.png', confidence = 0.8)
            pyautogui.moveTo(startButton)
            if startButton:
                pyautogui.click(startButton)

        def ebayisurekliekranalma():
            sleep(30)  
            hwnd = win32gui.FindWindow(None, "Canlı Ders Uygulaması")
            win32gui.SetForegroundWindow(hwnd) 
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 700,200,500,600, win32con.SWP_SHOWWINDOW)         
        while True:
            sleep(5)
            now = datetime.now().strftime("%H:%M")
            mickapatma()
            #burayı if .. or .. ile yazınca çalışmıyor anlamadım ve ebayi sürekli ekrana almayı girisyap fonksiyonundan sonra çalıştırmazsanız ders ekranı beyaz kalıyor
            if "08:41" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "09:21" == now:
                girisyap()
                sleep(62)   
                ebayisurekliekranalma()
            if "10:01" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "10:41" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "11:21" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "12:01" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "12:41" == now:
                girisyap()
                sleep(62)
                ebayisurekliekranalma()
            if "13:21" == now:
                girisyap()
                sleep(62)  
                ebayisurekliekranalma()    
  
if __name__ == '__main__':
    SMWinservice.parse_command_line()