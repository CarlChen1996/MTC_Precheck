# -*- coding: utf-8 -*-
# @Time    : 2/26/2019 4:47 PM
# @Author  : Carl
# @Email   : carl.chen@hp.com
# @File    : getVersion.py
# @project : MTC_Precheck
import os,win32con
import subprocess
import  shutil
import time

from win32timezone import *
import uuid
from win32com.client import GetObject
from basic_functions import get_os_type
def get_installed_software_version():
    """
    Most of software can be found by retrieving registry
    they are stored in ("HKEY_LOCAL_MACHINE/SOFTWARE/Microsoft/Windows/CurrentVersion/Uninstall" and
    "HKEY_LOCAL_MACHINE/SOFTWARE/WOW6432Node/Microsoft/Windows/CurrentVersion/Uninstall")
    Each software is made of its "DisplayName", "DisplayVersion" and "Publisher" property
    :return: software list
    """
    current_os = get_os_type()
    all_software = {}
    single_software={}
    branches = []
    if current_os == 'wes7e':
        key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall']
    elif current_os == 'wes7p' or current_os == 'wes8' or current_os == 'win10':
        # ?????????
        key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
               r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall']
        for sub_key in key:
            sub_key_handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
            sub_key_num = winreg.QueryInfoKey(sub_key_handle)[0]
            for i in range(sub_key_num+1):
                try:
                    branch_name = winreg.EnumKey(sub_key_handle, i)
                    branch_path = sub_key + '\\' + branch_name
                    branches.append(branch_path)
                    branch = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, branch_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
                    display_name = winreg.QueryValueEx(branch, 'DisplayName')[0]
                    if 'Adobe Flash Player' in display_name:
                        display_name = 'Adobe Flash Player'
                    display_version = winreg.QueryValueEx(branch, 'DisplayVersion')[0]
                    publisher = winreg.QueryValueEx(branch, 'Publisher')[0]
                    if display_name=='VMware Horizon Client':
                        display_version=display_version[:5]
                    if display_name=='Citrix Receiver Inside':
                        display_version=display_version[:4]
                    if display_name=='Citrix Workspace 1909':
                        display_version=display_version[:6]
                    if display_name=='HP System Default Settings':
                        display_version=display_version[:3]
                    # if display_name =='Intel(R) HD Graphics 610':
                    #     display_version=display_version[6:]
                    if display_name=='HPDM Agent 64-bit':
                        display_version=display_version[:8]
                    if display_name == 'Foxit Reader':
                        display_version=display_version[:5]
                    if display_name == 'Foxit PhantomPDF':
                        display_version=display_version[:3]
                    single_software[display_name] = display_version
                    if ('Microsoft Corporation' not in publisher
                        and 'Hewlett-Packard' not in publisher
                        and 'Citrix Systems, Inc.' not in publisher and
                        'Realtek' not in display_name and 'Intel Corporation' not in publisher
                        and 'Intel(R) Corporation' not in publisher
                        and 'Cisco Systems' not in publisher
                        and 'AMD' not in display_name
                        or display_name == 'Online Plug-in'
                        or display_name == 'Configuration Manager Client'
                                ):
                        all_software.update(single_software)
                except WindowsError:
                    pass
    else:
        print('Cannot find efficient registry')
    # filter_rep_list = []
    # for sw in all_software:
    #     if sw not in filter_rep_list:
    #         filter_rep_list.append(sw)
    return all_software
def get_all_device_drivers():
    try:
        wmi = GetObject(r'winmgmts:\\.\root\cimv2')
        # wmi = GetObject('winmgmts:') #??????
        # ??????????,?????????,?????????,????????????????
        # ??Realtek audio, ramdriver, USB, wlan, graphic, TPM, Intel TXE, chipset, AMD SMBus
        common_wql = ("Select * from Win32_PnPSignedDriver Where "
                      "(DeviceClass = 'MEDIA' and (DeviceName Like 'Realtek%')) or "
                      "DeviceClass = 'Bluetooth' and Not DeviceName Like 'Microsoft%' or "
                      "DeviceClass = 'Ramdrive' or "
                      "DeviceClass = 'Biometric' or "
                      "DeviceClass = 'Proximity' or "
                      "(DeviceClass = 'USB' and Not DeviceName Like '%USB Composite Device%')or "
                      "(DeviceClass = 'Net' and Not  DeviceName Like 'WAN Miniport%') or "
                      "DeviceClass = 'Display' or "
                      "DeviceClass = 'SecurityDevices' or "
                      "(DeviceClass = 'System' and (DeviceName Like 'Intel(R) Trusted Execution%' or DeviceName Like 'Intel(R)%Celeron%' or DeviceName Like '%AMD SMBus%'))"
                      )
        # common_wql=("Select * from Win32_PnPSignedDriver where "
        #             "not DeviceName Like '%USB Composite Device%'"
        #             "and not DeviceName like '%Microsoft%'"
        #             )
        drivers = wmi.ExecQuery(common_wql)
        driver_dict = {}
        for driver in drivers:
            if 'Bluetooth(R)' in str(driver.DeviceName):
                driver.DeviceName = 'Intel(R) Wireless Bluetooth(R)'
            if driver.DeviceName == 'Intel(R) HD Graphics 610':
                driver.DriverVersion = driver.DriverVersion[6:]
            driver_dict[str(driver.DeviceName)] = str(driver.DriverVersion)
        return driver_dict
    except WindowsError as e:
        print('sql syntax error\n', e)
# software_version.append(device_drivers)

if __name__ == "__main__":
    print(get_all_device_drivers())



