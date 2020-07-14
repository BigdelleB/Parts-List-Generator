import PySimpleGUI as sg
from pathlib import Path    
import Parts_list_helper_functions as parts_help
import Was_Is_helper_functions as was_help


sg.theme('SystemDefault1')

layout = [[sg.Text('Choose to create a Parts List or Was-Is List')],
			[sg.Button( 'Parts List Generator')],
		  [sg.Button('Was-Is List Generator')]]
   
  

window = sg.Window('Parts List Generator', layout, size = (500,100), keep_on_top=True)

event, values = window.read()


if event=="Parts List Generator":
	window.close()
	layout = [[sg.Text('Enter CSV Parts List')],
          [sg.Text('File:*', size=(8, 1)), sg.Input(), sg.FileBrowse()],
          [sg.Submit(), sg.Cancel()]];
	window = sg.Window('Parts List Generator', layout, size = (500,500), keep_on_top=True)
	event, values = window.read()
    #get the CCA name from the input file
	file_path = str(values[0]);
	file_name = Path(file_path).name;
	CCA_name = file_name[0:6];
	parts_help.main(file_path, CCA_name)

	#show the parts list prompt

elif event =="Was-Is List Generator":
	window.close()
	layout = [[sg.Text('Enter Original Parts List')],
	[sg.Text('File 1:*', size=(8,1)), sg.Input(), sg.FileBrowse()],
	[sg.Text('Enter New Parts List')],
	[sg.Text('File 2:*', size=(8,1)), sg.Input(), sg.FileBrowse()],
	[sg.Submit(), sg.Cancel()]];

	window = sg.Window('Was-Is Generator', layout, size = (500,200), keep_on_top=True)
	event, values = window.read()

	was_help.main(values[0], values[1])

	

window.close()
