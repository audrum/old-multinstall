import sys
import getopt
import subprocess
import winreg
import os
import colorama
from colorama.ansi import Fore, Style

def usage():
    colorama.init()
    print("%s v0.1" % sys.argv[0])
    print("Useful tool for installing and activating basic software")
    print(Fore.YELLOW + "===============================================================")
    print("FIRST AT ALL PLEASE USE ARGUMENT -I BEFORE USE ANY OTHER OPTION")
    print("===============================================================" + Style.RESET_ALL)
    print("Usage: %s [OPTIONS...]:" % sys.argv[0])
    print("-h   --help                      Get how to use it information")
    print("-I   --install-chocolatey        Installs Chocolatey, mandatory before running any other arguments")
    print("-c   --install-chrome            Installs Google Chrome browser")
    print("-f   --install-firefox           Installs Mozilla Firefox browser")
    print("-z   --install-7zip              Installs 7-zip file archiver")
    print("-p   --install-pdf24             Installs PDF24")
    print("-o   --install-office            Installs Microsoft Office") 
    print("-w   --activate-windows          Activates Microsoft Windows 10")  
    print("-x   --activate-office           Activates Microsoft Office")
    print("")
    print("Example: run install.py -I to install chocolatey")

def install_chocolatey():
    installscript = open("install_choco.ps1", "a")
    installscript.write("Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))")
    installscript.close()

    runscript = subprocess.Popen(["PowerShell.exe", "./install_choco.ps1"], stdout=sys.stdout)
    runscript.communicate()

    runscript.wait()
    os.system("refreshenv")
    os.remove("install_choco.ps1")

def install_chrome():
    runinstall = subprocess.Popen(["cinst", "GoogleChrome", "-y"], stdout=sys.stdout)
    runinstall.wait()

def install_firefox():
    runinstall = subprocess.Popen(["cinst", "Firefox", "-y"], stdout=sys.stdout)
    runinstall.wait()

def install_7zip():
    runinstall = subprocess.Popen(["cinst", "7zip", "-y"], stdout=sys.stdout)
    runinstall.wait()

def install_pdf24():
    runinstall = subprocess.Popen(["cinst", "pdf24", "-y"], stdout=sys.stdout)
    runinstall.wait()

def install_office():
    runinstall = subprocess.Popen(["cinst", "office365proplus", "-y"], stdout=sys.stdout)
    runinstall.wait()

def activate_windows():
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r'SOFTWARE\Microsoft\Windows NT\CurrentVersion')
    value = winreg.QueryValueEx(key, "ProductName")[0]

    print("Detected version %s installed" % value)
    print("Trying to activate %s..." % value)

    if value == "Windows 10 Home Single Language":
        os.system("slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)
    
    elif value == "Windows 10 Home":
        os.system("slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)

    elif value == "Windows 10 Pro":
        os.system("slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)

    elif value == "Windows 11 Home Single Language":
        os.system("slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)
    
    elif value == "Windows 11 Home":
        os.system("slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)

    elif value == "Windows 11 Pro":
        os.system("slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX")
        os.system("slmgr /skms kms8.msguides.com")
        os.system("slmgr /ato")
        print("===%s has been activated..." % value)

    else:
        print("This program can't activate this Windows version, contact the creator for further help")

def activate_office():
    print("Activating..")
    file = open("actoffice.cmd", "a")
    file.write("@echo off\n")
    file.write('''
    @echo off
    title Multinstall&cls&echo =====================================================================================&echo #Project: Activating Microsoft Office&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&set i=1&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
    :skms
    if %i% GTR 10 goto busy
    if %i% EQU 1 set KMS=kms7.MSGuides.com
    if %i% EQU 2 set KMS=s8.now.im
    if %i% EQU 3 set KMS=s9.now.im
    if %i% GTR 3 goto ato
    cscript //nologo ospp.vbs /sethst:%KMS% >nul
    :ato
    echo ============================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #Repository: github.com/audrum&echo.&echo #Please feel free to contact me at t.me/audrum if you have any questions or concerns.&echo.&echo ============================================================================) || (echo The connection to the KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
    goto halt
    :notsupported
    echo ============================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
    :busy
    echo ============================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
    :halt
    pause >nul
    ''')
    file.close()

    activate = subprocess.Popen(["actoffice.cmd"], stdout=sys.stdout)
    activate.wait()

    os.remove("actoffice.cmd")



def main():
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hIcfzpowx", ["help", "install-chocolatey", "install-chrome", "install-firefox", "install-7zip", "install-pdf24", "install-office", "activate-windows", "activate-office"])
    except getopt.GetoptError as err:
        print(err)
        print("Try %s -h for help" % sys.argv[0])
        sys.exit(0)
    for o, a in opts:
        if o in ['-h', '--help']:
            usage()
        elif o in ['-I', '--install_chocolatey']:
            install_chocolatey()
        elif o in ['-c', '--install-chrome']:
            install_chrome()
        elif o in ['-f', '--install-firefox']:
            install_firefox()
        elif o in ['-z', '--install-7zip']:
            install_7zip()
        elif o in ['-p', '--install-pdf24']:
            install_pdf24()
        elif o in ['-o', '--install-office']:
            install_office()
        elif o in ['-w', '--activate-windows']:
            activate_windows()
        elif o in ['-x', '--activate-office']:
            activate_office()
        else:
            usage()

main()