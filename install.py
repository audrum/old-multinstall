import sys
import getopt
import subprocess
import winreg
import os
import platform

def usage():
    print("%s v0.1" % sys.argv[0])
    print("Useful tool for installing and activating basic software")
    print("===============================================================")
    print("FIRST AT ALL PLEASE USE ARGUMENT -c BEFORE USE ANY OTHER OPTION")
    print("===============================================================")
    print("Usage: %s [OPTIONS...]:" % sys.argv[0])
    print("-h   --help                      Get how to use it information")
    print("-I   --install-chocolatey        Installs Chocolatey, mandatory to run other arguments")
    print("-c   --install-chrome            Installs Google Chrome browser")
    print("-f   --install-firefox           Installs Mozilla Firefox browser")
    print("-z   --install-7zip              Installs 7-zip file archiver")
    print("-p   --install-pdf24             Installs PDF24")
    print("-o   --install-office            Installs Microsoft Office") 
    print("-w   --activate_windows          Activates Microsoft Windows 10")   
    print("")
    print("Example: run install.py -c to install chocolatey")

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
    os.system("cinst GoogleChrome -y")

def install_firefox():
    os.system("cinst Firefox -y")

def install_7zip():
    os.system("cinst 7zip -y")

def install_pdf24():
    os.system("cinst pdf24 -y")

def install_office():
    os.system("cinst office365proplus -y")

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

    else:
        print("This program can't activate this Windows version, contact the creator for further help")

def main():
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hIcfzpow", ["help", "install-chocolatey", "install-chrome", "install-firefox", "install-7zip", "install-pdf24", "install-office", "activate-windows"])
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
        else:
            usage()

main()