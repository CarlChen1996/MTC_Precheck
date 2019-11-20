# -*- coding: utf-8 -*-
# Created by Carl on 2/22/2019
import os
import time

from jinja2 import Environment, FileSystemLoader
from twisted.python.compat import raw_input
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import pandas as pd
import basic_functions as bas,complex_functions as com
from getVersion import get_all_device_drivers,get_installed_software_version
from mlVersion import get_MlFile,get_Version


def check():
    while True:
        platforms=['mt21','mt22','mt31','mt44','mt45']
        for platform in platforms:
            print(str(platforms.index(platform))+':'+platform)
        platform_index = int(raw_input('Choose your platform.'))
        if platform_index in range(len(platforms)):
            ml = get_MlFile('ml_list/'+platforms[platform_index])
            if ml:
                for k, v in ml.items():
                    print(k, v)
                ml_index = raw_input('Choose your ml please.')
                if ml_index.isdigit():
                    if int(ml_index) in ml.keys():
                        ml_name = ml[int(ml_index)]
                        print(ml_name)
                        break
            else:
                print('-There is no ml file of your platform.')
    # 本地时间
    local_time = bas.get_local_time()
    # MAC地址
    mac = bas.get_mac()
    # kb
    kb = bas.get_kb()
    # 时区
    timezone = bas.get_timezone()
    # 系统类型
    os_type = bas.get_os_type()
    # Ml
    ml = com.get_ml()
    # BIOS
    bios = com.get_bios()
    # ml
    d1 = get_Version(ml_name,platforms[platform_index])  # ml
    # local
    d = get_installed_software_version().copy()
    d.update(get_all_device_drivers())
    d2 = d.copy()
    a = pd.DataFrame(d2.items(), columns=['Name', 'Version'])
    a.to_excel(os.getcwd() + '/report/local_version.xls')
    software_right = []
    version_right = []
    software_wrong = []
    version_wrong = []
    for k1, v1 in d1.items():
        if k1 in d2.keys():
            if v1 == d2[k1]:
                version_right.append([k1,v1])
            else:
                version_wrong.append([k1,v1+'|'+d2[k1]])
            software_right.append([k1,v1])
        else:
            software_wrong.append([k1,v1])
    # print(software_right,len(software_right))
    # print(version_right,len(version_right))
    # print(version_wrong,len(version_wrong))
    # print(software_wrong,len(software_wrong))
    result = {}
    result['Time'] = local_time
    result['OStype'] = os_type
    result['BIOS'] = bios
    result['TimeZone'] = timezone
    result['mac'] = mac
    result['kb'] = kb
    result['uninstalled'] = software_wrong
    result['wrong-version'] = version_wrong
    result['version_right'] = version_right
    result['ml_name'] = ml_name[:-4]
    return result


def result(result):
    env = Environment(loader=FileSystemLoader('./templates', encoding='utf-8'))
    uninstalled = len(result['uninstalled'])
    wrong_version = len(result['wrong-version'])
    version_right = len(result['version_right'])
    count = uninstalled + wrong_version + version_right
    total = {'uninstalled':uninstalled, 'wrong-version': wrong_version,'version_right':version_right,
             'Count':count }
    data = [{'value': total['wrong-version'], 'name': 'Wrong-version', 'itemStyle': {'color': '#FFEC8B'}},
            {'value': total['uninstalled'], 'name': 'Uninstalled', 'itemStyle': {'color': '#d9534f'}},
            {'value': total['version_right'], 'name': 'Version_right', 'itemStyle': {'color': '#5cb85c'}},
            ]
    template = env.get_template('base.html')
    html = template.render(result=result, data=data, total=total,
                           encoding='utf-8')  # unicode string
    with open('report/{}.html'.format(result['ml_name']), 'w', encoding='utf-8') as f:
        f.write(html)
    return result['ml_name']+'.html'


if __name__ == '__main__':
    data = check()
    html = result(data)
    os.popen(os.getcwd()+'/report/'+html)
