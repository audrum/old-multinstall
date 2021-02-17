import PySimpleGUI as sg

frame_layout_programs = [

            [sg.Text("Please select the programs to install")],
            [sg.Checkbox("Google Chrome"), sg.Checkbox("Firefox")],
            [sg.Checkbox("7-Zip"), sg.Checkbox("WinRAR")],
            [sg.Checkbox("PDF24"), sg.Checkbox("Microsoft Office")],
            [sg.Checkbox("Microsoft Project"), sg.Checkbox("Microsoft Visio")]
]

frame_layout_activations = [

            [sg.Text("Please select the software to activate")],
            [sg.Checkbox("Windows 10"), sg.Checkbox("Microsoft Office")],
            [sg.Checkbox("Microsoft Project"), sg.Checkbox("Microsoft Visio")]
]

layout = [

            [sg.Frame("Programs", frame_layout_programs), sg.Frame("Activations", frame_layout_activations)],
            [sg.Submit(), sg.Cancel()] 
]

window = sg.Window("Multinstaller", layout)

event,  values = window.read()
window.close()  