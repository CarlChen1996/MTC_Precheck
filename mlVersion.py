# -*- coding: utf-8 -*-
# Created by Carl on 2/22/2019
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
from datetime import datetime
import pandas as pd
import os


def get_MlFile(file_dir):
    for root, dirs, files in os.walk(file_dir):
        # print('root_dir:', root)  # 当前目录路径
        # print('sub_dirs:', dirs)  # 当前路径下所有子目录
        # print('files:', files)  # 当前路径下所有非目录子文件
        index = [x for x in range(0,len(files))]
        d = {}
        for i in range(len(files)):
            d[index[i]] = files[i]
        return d

def get_Version(MlName,platform):
    # df = pd.read_excel('ml_list/mt45.xls',usecols= ['Name', 'Version'])
    ml_path='ml_list/'+platform+'/'+MlName
    df = pd.read_excel(ml_path,usecols= ['Name','Version'])
    ml_data = df.loc[:].values  # 读取指定多行的话，就要在loc[]里面嵌套列表指定行数
    final_dict={}
    final_list=[]

    # 7 items
    def mt45(ml_data):
        for j in ml_data:
            if 'Realtek Ethernet Controller' in j[0]:
                name = 'Realtek PCIe GbE Family Controller'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'AMD' and 'Video Driver' in j[0]:
                name = 'AMD Radeon(TM) Vega 6 Graphics'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'AMD' and 'Chipset' in j[0]:
                name = 'AMD SMBus'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Intel WLAN Driver' in j[0]:
                name = 'Intel(R) Wireless-AC 9260 160MHz'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Collaboration Keyboard' in j[0]:
                name = 'HP Collaboration Keyboard for Skype for Business'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP lt4210 LTE/HSPA+' in j[0]:
                name = 'Intel(R) XMM(TM) 7360 LTE-A Driver Package'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Touch Fingerprint' in j[0]:
                name = 'Synaptics VFS7552 Touch Fingerprint Sensor with PurePrint(TM)'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'NFC' in j[0]:
                name = 'NxpNfcClientDriver'
                final_list.append(name)
                final_dict[name] = j[1]

    # 5 items
    def mt31(ml_data):
        for j in ml_data:
            if 'Realtek Ethernet Controller' in j[0]:
                name = 'Realtek PCIe GbE Family Controller'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Video Driver' in j[0]:
                name = 'Intel(R) HD Graphics 610'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Chipset' in j[0]:
                name = 'Intel(R) Chipset Device Software'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Intel WLAN Driver' in j[0]:
                name = 'Intel(R) Dual Band Wireless-AC 8265'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Universal Camera Driver' in j[0]:
                name = 'HP Universal Camera Driver'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Dynamic Platform and Thermal Framework' in j[0]:
                name = 'Intel(R) Dynamic Platform and Thermal Framework'
                final_list.append(name)
                final_dict[name] = j[1]

    # 9 items
    def mt21(ml_data):
        for j in ml_data:
            if '2017 Synaptics Mouse Driver' == j[0]:
                name = 'Synaptics Pointing Device Driver'
                final_dict[name] = j[1]
            if 'HP User State' in j[0]:
                name = 'HP User State Tool'
                final_dict[name] = j[1]
            if 'Foxit PhantomPDF' in j[0]:
                name = 'Foxit PhantomPDF'
                final_dict[name] = j[1]
            if 'HP Hotkey Filter' in j[0]:
                name = 'HP Hotkey Filter'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Realtek USB and PCIe Media Card Reader' in j[0]:
                name = 'Realtek Card Reader'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Intel WLAN Driver' in j[0]:
                name = 'Intel(R) Dual Band Wireless-AC 8265'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Chipset' in j[0]:
                name = 'Intel(R) Chipset Device Software'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Realtek Ethernet Controller' in j[0]:
                name = 'Realtek Ethernet Controller Driver'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Dynamic Platform and Thermal Framework' in j[0]:
                name = 'Intel(R) Dynamic Platform and Thermal Framework'
                final_list.append(name)
                final_dict[name] = j[1]

    # 6 items
    def mt22(ml_data):
        for j in ml_data:
            if 'Realtek Ethernet Controller' in j[0]:
                name = 'Realtek PCIe GbE Family Controller'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Intel WLAN Driver' in j[0]:
                name = 'Intel(R) Wireless-AC 9560 160MHz'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Touch Fingerprint' in j[0]:
                name = 'Synaptics FS7604 Touch Fingerprint Sensor with PurePrint(TM)'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Chipset' in j[0]:
                name = 'Intel(R) Chipset Device Software'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Citrix Workspace' in j[0]:
                name = 'Citrix Workspace 1909'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Intel Dynamic Tuning' in j[0]:
                name = 'Intel(R) Dynamic Tuning'
                final_list.append(name)
                final_dict[name] = j[1]
    def mt44(ml_data):
        pass

    # 18 items
    def common(ml_data):
        for j in ml_data:
            if 'CyberLink' in j[0]:
                name = 'CyberLink Power Media Player 14'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Connection Optimizer' in j[0]:
                name = 'HP Connection Optimizer'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Documentation' in j[0]:
                name = 'HP Documentation'
                final_list.append(name)
                j[1] = j[1][:-2] + '0.' + j[1][-2] + '.' + j[1][-1]
                # if j[1]=='1.01':
                #     j[1]='1.0.0.1'
                # if j[1]=='1.00':
                #     j[1]='1.0.0.0'
                final_dict[name] = j[1]
            if 'HP Easy Shell' in j[0]:
                name = 'HP Easy Shell'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Function Key Filter' in j[0]:
                name = 'HP Function Key Filter'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Hotkey Support' in j[0]:
                name = 'HP Hotkey Support'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Ramdrive Driver' in j[0]:
                name = 'HP RAM Disk Device (disk)'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP System Information' in j[0]:
                name = 'HP System Information'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP ThinUpdate' in j[0]:
                name = 'HP ThinUpdate'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP USB Port Manager' in j[0]:
                name = 'HP USB Port Manager'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Foxit Reader' in j[0]:
                name = 'Foxit Reader'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'VMware Horizon Client' in j[0]:
                name = 'VMware Horizon Client'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HP Write Filter Manager' in j[0]:
                name = 'HP Write Manager'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'HPDM Agent' and '64-bit' in j[0]:
                name = 'HPDM Agent 64-bit'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'System Default Settings' in j[0]:
                name = 'HP System Default Settings'
                final_list.append(name)
                final_dict[name] = j[1]
            # if 'HP Wireless Button Driver' in j[0]:
            #     name = 'HP Wireless Button Driver'
            #     final_dict[name] = j[1]
            if 'Intel Bluetooth Driver' in j[0]:
                name = 'Intel(R) Wireless Bluetooth(R)'
                final_list.append(name)
                final_dict[name] = j[1]
            if 'Citrix Receiver' in j[0]:
                name = 'Citrix Receiver Inside'
                final_list.append(name)
                final_dict[name] = j[1]

    common(ml_data)
    if platform == 'mt21':
        mt21(ml_data)
    if platform == 'mt22':
        mt22(ml_data)
    if platform == 'mt31':
        mt31(ml_data)
    if platform == 'mt44':
        mt44(ml_data)
    elif platform == 'mt45':
        mt45(ml_data)
    a = pd.DataFrame(final_dict.items(),columns=['Name','Version'])
    a.to_excel(os.getcwd()+'/report/ml_version.xls')
    return final_dict


if __name__ == '__main__':
    print(get_Version('18WWLTAJ6ao.xls'))
