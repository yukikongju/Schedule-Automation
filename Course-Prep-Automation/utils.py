#!/usr/bin/python

import glob
import os
import selenium

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

#  from glob import glob

def convert_pdf_file_to_ppt(file_path, download_dir):
    """ Converting PDF file to powerpoint
        
        Params:
            - file_path: take pdf file path (ex: winter2020/file.pdf)
            - download_dir: the directory in which the ppt should be downloaded

        Return: void

    """
    url = "https://www.adobe.com/ca/acrobat/online/pdf-to-ppt.html"

    # init driver
    driver = webdriver.Chrome()

    button = driver.find_element_by_id("lifecycle-nativebutton")
    button.click()

    pass

def convert_directory_pdf_files_to_ppt(dir_path, download_dir):
    """ Convert all pdf files from directory to powerpoint files in download files 
        Generates powerpoint_dir in file_path
    
        Params:
            - dir_path: take pdf file path (ex: winter2020/file.pdf)
            #  - download_dir: the directory in which the ppt should be downloaded

        Return: void
    
    """
    # get list of pdf in directory
    pdfs = get_pdf_files_from_dir(dir_path)
    #  print(pdfs)

    # init driver
    driver = webdriver.Chrome()
    url = "https://www.adobe.com/ca/acrobat/online/pdf-to-ppt.html"

    # convert all pdf files in powerpoint
    for pdf in pdfs:
        driver.get(url)
        upload_button = driver.find_element_by_id("lifecycle-nativebutton")
        upload_button.click()
        # TODO: import file in upload box
        #  driver.clear()
        upload_button.send_keys(pdf)
        upload_button.send_keys(Keys.RETURN)
    pass
    
def get_pdf_files_from_dir(dir_path): 
    os.chdir(dir_path)
    #  print(dir_path)
    pdfs = []
    for mfile in glob.glob("*.pdf"):
        #  print(mfile)
        file_path = f"{dir_path}/{mfile}"
        pdfs.append(mfile)
    return pdfs

def test():
    #  dir_path = "C:/Users/emuli/OneDrive - Universite de Montreal/Bac-Maths-Info/Computer Science/Bloc A - Data Structure, Computer System, Computation Theory/Computer Systems/CMU 15-213/Slides - CMU 15-213 Computer Systems"
    dir_path = "C:\\Users\\emuli\\OneDrive - Universite de Montreal\\Bac-Maths-Info\\Computer Science\\Bloc A - Data Structure, Computer System, Computation Theory\\Computer Systems\\CMU 15-213\\Slides - CMU 15-213 Computer Systems"
    download_dir = ""
    convert_directory_pdf_files_to_ppt(dir_path, download_dir)
    #  get_pdf_files_from_dir(dir_path)

test()

