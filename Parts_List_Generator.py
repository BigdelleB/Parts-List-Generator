import PySimpleGUI as sg
from pathlib import Path    
import sys
import helper_functions as help
import os


sg.theme('SystemDefault1')

layout = [[sg.Text('Enter CSV Parts List')],
          [sg.Text('File:*', size=(8, 1)), sg.Input(), sg.FileBrowse()],
          [sg.Submit(), sg.Cancel()],
          [sg.Text("\n")],
          [sg.Text("*Parts List columns must match below:")],
          [sg.Text("\"ITEM\", \"QTY\", \"REFERENCE\", \"DESCRIPTION\", \"MANUFACTURER\",\"MANUFACTURER P/N\"")],
          [sg.Text("Make sure that there are no line gaps in the parts list or missing fields")]]


window = sg.Window('Parts List Generator', layout, size = (500,200), keep_on_top=True)

event, values = window.read()

#get the CCA name from the input file
file_path = str(values[0]);
file_name = Path(file_path).name;
CCA_name = file_name[0:6];

base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
cover_page = os.path.join(base_path,"Cover Page.xlsx")


help.main(file_path, CCA_name,cover_page)



window.close()
