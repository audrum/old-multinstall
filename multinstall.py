import PySimpleGUI as sg
import subprocess
import sys
import os

column_programs_1 = [

            [sg.Checkbox("Google Chrome", key="-GoogleChrome-")],
            [sg.Checkbox("7-Zip", key="-7zip-")],
            [sg.Checkbox("PDF24", key="-PDF24-")],
            [sg.Checkbox("Microsoft Project", key="-Project-")]
                              
]

column_programs_2 = [

            [sg.Checkbox("Firefox", key="-Firefox-")],
            [sg.Checkbox("WinRAR", key="-WinRAR-")],
            [sg.Checkbox("Microsoft Office", key="-Office-")],
            [sg.Checkbox("Microsoft Visio", key="-Visio-")]
                              
]

column_activations_1 = [

            [sg.Checkbox("Windows 10", key="-ActW10-")],
            [sg.Checkbox("Microsoft Project", key="-ActProject-")]
]

column_activations_2 = [

            [sg.Checkbox("Microsoft Office", key="-ActOffice-")],
            [sg.Checkbox("Microsoft Visio", key="-ActVisio-")]
]

frame_layout_programs = [

            [sg.Text("Please select the programs to install")],
            [sg.Col(column_programs_1), sg.Col(column_programs_2)]
]

frame_layout_activations = [

            [sg.Text("Please select the software to activate")],
            [sg.Col(column_activations_1), sg.Col(column_activations_2)]
]

frame_layout_result = [

            [sg.Output(size=(80,10))]
]

layout = [

            [sg.Text("WORKS ONLY ON WINDOWS 10 x64 bits!", text_color="red")],
            [sg.Frame("Programs", frame_layout_programs), sg.Frame("Activations", frame_layout_activations)],
            [sg.Submit(), sg.Cancel()],
            [sg.Frame("Result", frame_layout_result)]
]

window = sg.Window("Multinstaller", layout) 

agree = sg.popup_yes_no("Checking System", "This program needs Chocolatey in order to work properly, we will check if it is installed, if not it will be installed, do you agree?")

if agree == "Yes":

    if "ChocolateyLastPathUpdate" in os.environ:
        version = subprocess.check_output("choco -v")
        sg.Popup("You already have installed Chocolatey version", version.decode("cp850"))
    
    else:
        choco_install = subprocess.check_output(["Powershell.exe", "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"])
        [sg.popup_scrolled("Chocolatey is not installed, installing...", choco_install.decode("cp850"))]

else:
    exit()

while True:
    event,  values = window.read()
    if event == sg.WIN_CLOSED or event == "Cancel":
        break

    elif event == "Submit":
        if values["-GoogleChrome-"] == True:
            print("Installing Google Chrome...")
            gc_install = subprocess.check_output("cinst GoogleChrome -y")
            print(gc_install.decode("cp850"))
            print("")
        if values["-Firefox-"] == True:
            print("Installing Firefox...")
            fx_install = subprocess.check_output("cinst Firefox -y")
            print(fx_install.decode("cp850"))
            print("")
        if values["-7zip-"] == True:
            print("Installing 7-zip...")
            sevenz_install = subprocess.check_output("cinst 7zip -y")
            print(sevenz_install.decode("cp850"))
            print("")
        if values["-WinRAR-"] == True:
            print("Installing WinRAR...")
            wr_install = subprocess.check_output("cinst winrar -y")
            print(wr_install.decode("cp850"))
            print("")
        if values["-PDF24-"] == True:
            print("Installing PDF24...")
            p24_install = subprocess.check_output("cinst pdf24 -y")
            print(p24_install.decode("cp850"))
            print("")
        if values["-Office-"] == True:
            print("Installing Microsoft Office...")
            mo_install = subprocess.check_output("cinst office365proplus -y")
            print(mo_install.decode("cp850"))
            print("")
        if values["-Project-"] == True:
            print("Installing Microsoft Project...")
            mp_install = subprocess.check_output("cinst microsoft-office-deployment --params="'/64bit /Product:ProjectPro2019Volume'" --force -y")
            print(mp_install.decode("cp850"))
            print("")
        if values["-Visio-"] == True:
            print("Installing Microsoft Visio...")
            mv_install = subprocess.check_output("cinst microsoft-office-deployment --params="'/64bit /Product:VisioPro2019Volume'" --force -y")
            print(mv_install.decode("cp850"))
            print("")
        if values["-ActW10-"] == True:
            print("Activating Windows 10...")
            string = subprocess.check_output(["Powershell.exe", "(Get-WmiObject Win32_OperatingSystem).Caption"])
            if "Microsoft Windows 10 Home Single Language" in string:
                print("Installing license for Microsoft Windows 10 Home Single Language...")
                step1 = subprocess.check_output("slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH")
                print(step1)
                step2 = subprocess.check_output("slmgr /skms kms8.msguides.com")
                print(step2)
                step3 = subprocess.check_output("slmgr /ato")
                print(step3)
            elif "Microsoft Windows 10 Home" in string:
                print("Installing license for Microsoft Windows 10 Home...")
                step1 = subprocess.check_output("slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99")
                print(step1)
                step2 = subprocess.check_output("slmgr /skms kms8.msguides.com")
                print(step2)
                step3 = subprocess.check_output("slmgr /ato")
                print(step3)
            elif "Microsoft Windows 10 Pro" in string:
                print("Installing license for Microsoft Windows 10 Pro...")
                step1 = subprocess.check_output("slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX")
                print(step1)
                step2 = subprocess.check_output("slmgr /skms kms8.msguides.com")
                print(step2)
                step3 = subprocess.check_output("slmgr /ato")
                print(step3)
            elif "Microsoft Windows 10 Pro N" in string:
                print("Installing license for Microsoft Windows 10 Pro...")
                step1 = subprocess.check_output("slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9")
                print(step1)
                step2 = subprocess.check_output("slmgr /skms kms8.msguides.com")
                print(step2)
                step3 = subprocess.check_output("slmgr /ato")
                print(step3)
            else:
                print("Unable to determinate Windows version, Windows not activated")
        if values["-ActOffice-"] == True:
            print("Activating Microsoft Office...")
            pf_dir = os.environ.get("ProgramFiles")
            office_dir = pf_dir+"/Microsoft Office/root/Licenses16/"
            office_script = pf_dir+"/Microsoft Office/Office16/"
            files = os.listdir(office_dir)

            for i in files:
                if "ProPlus2019VL" in i:
                    os.system("cd " + office_script + " && " + "cscript ospp.vbs /inslic:" + '"' + "../root/licenses16/" + i + '"')

            os.system("cd " + office_script + " && " + "cscript ospp.vbs /setprt:1688")
            print("Conecting to port...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /unpkey:6MWKP >nul")
            print("Changing key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP")
            print("Installing new key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /sethst:kms8.msguides.com")
            print("Connecting to server for activation...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /act")
            print("Microsoft Office has been activated!")
        if values["-ActProject-"] == True:
            print("Activating Microsoft Project...")
            pf_dir = os.environ.get("ProgramFiles")
            office_dir = pf_dir+"/Microsoft Office/root/Licenses16/"
            office_script = pf_dir+"/Microsoft Office/Office16/"
            files = os.listdir(office_dir)

            for i in files:
                if "ProjectPro2019VL_KMS" in i:
                    os.system("cd " + office_script + " && " + "cscript ospp.vbs /inslic:" + '"' + "../root/licenses16/" + i + '"')

            os.system("cd " + office_script + " && " + "cscript ospp.vbs /setprt:1688")
            print("Conecting to port...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /unpkey:6MWKP >nul")
            print("Changing key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B")
            print("Installing new key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /sethst:kms8.msguides.com")
            print("Connecting to server for activation...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /act")
            print("Microsoft Project has been activated!")
        if values["-ActVisio-"] == True:
            print("Activating Microsoft Project...")
            pf_dir = os.environ.get("ProgramFiles")
            office_dir = pf_dir+"/Microsoft Office/root/Licenses16/"
            office_script = pf_dir+"/Microsoft Office/Office16/"
            files = os.listdir(office_dir)

            for i in files:
                if "VisioPro2019VL_KMS" in i:
                    os.system("cd " + office_script + " && " + "cscript ospp.vbs /inslic:" + '"' + "../root/licenses16/" + i + '"')

            os.system("cd " + office_script + " && " + "cscript ospp.vbs /setprt:1688")
            print("Conecting to port...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /unpkey:6MWKP >nul")
            print("Changing key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB")
            print("Installing new key...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /sethst:kms8.msguides.com")
            print("Connecting to server for activation...")
            os.system("cd " + office_script + " && " + "cscript ospp.vbs /act")
            print("Microsoft Visio has been activated!")

window.close()