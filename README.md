# windows_activators

This is a project about open source activators for windows and office written in batch and powershell. **You have to open each one with admin rights**

## Office 2019

The [O2019.cmd](https://github.com/Audrum/windows_activators/blob/master/O2019.cmd) file is a batch script for activate Office 2019, it was written following some instructions from [msguides.com](https://msguides.com/) with a few modifications. Just open it and let it work. You can review the code and feel free to modify it. 

## Office 365

The [O365.cmd](https://github.com/Audrum/windows_activators/blob/master/O365.cmd) file is a batch script for activate Office 365, it was written following some instructions from [msguides.com](https://msguides.com/) with a few modifications. Just open it and let it work. You can review the code and feel free to modify it. 

## Windows 10

The [W10.cmd](https://github.com/Audrum/windows_activators/blob/master/W10.cmd) file is a batch script that can activate the following flavors of Windows 10:

* Windows 10 Home
* Windows 10 Home Single Language
* Windows 10 Pro

Also it can show you a list of serials for each Windows 10 flavor and even use your own serial if you have one. Just open it and follow the on screen instructions. Feel free to review the code and modify it.


## Windows Multinstall

The [multinstall.ps1](https://github.com/Audrum/windows_activators/blob/master/multinstall.ps1) file is a PowerShell script to make a little bit easier and faster the Windows 10 after installation, helping installing automatically some basic tools such as: 

* Google Chrome
* Firefox
* 7-zip
* Optionally Office 365

Also includes the Windows 10 activator ported to PowerShell. Currently working in porting the Office activator. The installing of those programs is made installing and invoking [Chocolatey](https://chocolatey.org/). 

To execute the script, run a PowerShell terminal with admin rights, move to the path where is downloaded the script and execute the script with the command ``.\windows_multinstaler.ps1`` and follow the instructions. Feel free to review the code and modify it.

