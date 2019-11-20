from basic_functions import  *
passlist=[]
faillist=[]
def get_ml():
    try:
        mlpath="SYSTEM\CurrentControlSet\Control\WindowsEmbedded\RunTimeID"
        if check_registry_exist(winreg.HKEY_LOCAL_MACHINE,mlpath,"RunTimeOEMRev"):
            return read_reg(winreg.HKEY_LOCAL_MACHINE,mlpath,"RunTimeOEMRev")
    except:
        return "ML string not exist"
def get_bios():
    try:
        mlpath="HARDWARE\DESCRIPTION\System\BIOS"
        if check_registry_exist(winreg.HKEY_LOCAL_MACHINE,mlpath,"BIOSVersion"):
            return read_reg(winreg.HKEY_LOCAL_MACHINE,mlpath,"BIOSVersion")
    except:
        return "BIOS string not exist"
def check_mlgm():
    if (get_os_type().lower()=="wes7e") | (get_os_type().lower()=="wes7p"):
        if check_file_exist("C:\SYSTEM.SAV\FLAGS\MLGM.flg"):
            return "mlgm"
        else:
            return "normal"
    else:
        return "normal"
def read_txt_16(path):
    try:
        with open(path, 'r', encoding='utf-16') as file_object:
            contents = file_object.read()
            # print(contents)
            return contents
    except:
        print('Read txt file fail')
def remove_empty(strs:str):
    try:
        return strs.replace(" ", "")
    except:
        print("error remove empty string")
def record(path:str,lists:list):
    try:

        for i in lists:
            write_txt(path,i+"\n")
    except:
        return  False
def logs(strings):

    debug_log_folder=get_current_dir()+"\\log\\debug"
    debug_log="log.txt"
    if check_file_exist(debug_log_folder):
        debug_log=debug_log_folder+"\\"+debug_log
    else:
        create_folder(debug_log_folder)
        debug_log = debug_log_folder + "\\" + debug_log

    if check_file_exist(debug_log):
        write_txt(debug_log,get_local_time()+"_"+strings+"\n")
    else:
        create_txt(debug_log)
        write_txt(debug_log, get_local_time() + "_" + strings + "\n")
def run_ps1(command):
    subprocess.run("powershell -ExecutionPolicy Unrestricted -File " + command)
def check_image_info(language):
    mls=get_ml()
    temp=mls.split("#")
    lans=language
    #temp=get_ml().split("#")
    if temp:
        #print(temp[0][-2:])
        if temp[0][-2:].isdigit():
            data_list=[]
            for i in read_txt(get_current_dir() + r"\data\ML_ibr.txt").split("\n"):
                if(i):
                    data_list.append(list(i.split(":")))
            #print(data_list)
            index=-1
            for i in data_list:
                if i[0].lower()==lans.lower():
                    index=data_list.index(i)
                    break
            #print(index)
            if index>=0:
                if temp[2].lower()==data_list[index][1].lower():
                    print("Check HP system information:pass:" + mls)
                    passlist.append("Check HP system information:pass:" + mls)
                else:
                    print("Check HP system information:fail:" + mls + ":data incorrect")
                    faillist.append("Check HP system information:fail:" + mls + ":data incorrect")
            else:
                print("Check HP system information:fail:" + mls + ":not found language in ML_ibr.txt")
                faillist.append("Check HP system information:fail:" + mls + ":not found language in ML_ibr.txt")
        else:
            data_list = []
            for i in read_txt(get_current_dir()+r"\data\ML_dash.txt").split("\n"):
                #print(i)
                if(i):
                    data_list.append(list(i.split(":")))
           # print(data_list)
            index=-1
            for i in data_list:
                if i[0].lower()==lans.lower():
                    index=data_list.index(i)
                    break
            #print(index)
            if index>=0:
                if (temp[1].lower()==data_list[index][1].lower()) &(temp[2].lower()==data_list[index][2].lower()):
                    print("Check HP system information:pass:" + mls)
                    passlist.append("Check HP system information:pass:" + mls)
                else:
                    print("Check HP system information:fail:" + mls + ":data incorrect")
                    faillist.append("Check HP system information:fail:" + mls + ":data incorrect")
            else:
                print("Check HP system information:fail:"+mls+":not found language in ML_dash.txt")
                faillist.append("Check HP system information:fail:"+mls+":not found language in ML_dash.txt")
def check_qfe():
    checkpath=r"c:\system.sav\Logs\OSIT\QFESuppPack.LOG"
    if check_file_exist(checkpath):
        data_list = []
        for i in read_txt(checkpath).split("\n"):
            #print(i)
            if(i):
                data_list.append(i)
       # print(data_list)
        check_qfe_count=0
        for i in data_list:
            if i:
                if (len(i.split(":"))==2)&("Error" in i):
                    if (i.split(":")[0].strip()=="Error")&(i.split(":")[1].strip().isdigit()):
                     #print(i)
                     faillist.append("Check QFE install log no errors:fail:" + checkpath + ":"+i)
                     print("Check QFE install log no errors:fail:" + checkpath + ":"+i)
                     check_qfe_count=check_qfe_count+1
                #else:
                    #passlist.append("check_qfe:pass:" + checkpath)
                    #print("check_qfe:pass:" + checkpath)
        if check_qfe_count==0:
            passlist.append("Check QFE install log no errors:pass:" + checkpath )
            print("Check QFE install log no errors:pass:" + checkpath )
    else:
        faillist.append("Check QFE install log no errors:fail:" + checkpath + ":not found log file")
        print("Check QFE install log no errors:fail:" + checkpath + ":not found log file")
def check_defragmentation():
    rootpath=winreg.HKEY_LOCAL_MACHINE
    regPath="SYSTEM\CurrentControlSet\Services\defragsvc"
    if check_registry_exist(rootpath,regPath,"start"):
        regValue=read_reg(rootpath,regPath,"start")
        if regValue==4:
            passlist.append("check defragmentation Disable:pass:HKEY_LOCAL_MACHINE\%s\start=%d"%(regPath,regValue))
            logs("Defragmentation Check Result:PASS")
        else:
            faillist.append("check defragmentation Disable:fail:HKEY_LOCAL_MACHINE\%s\start=%d"%(regPath,regValue))
            logs("Defragmentation Check Result:(Value Error)FAIL")
    else:
        faillist.append("check Defragmentation Disable:fail:Registry value not exist!!")
        logs("Defragmentation Check Result:(Reg not exist)FAIL")
def check_bcd():
    '''
     After eventview.ps1, check Z:\\bcd.txt exist
     Z:\\bcd.txt contains "IgnoreAllFailures", check if exist key word
    '''
    filePath="z:\\bcd.txt"
    if check_file_exist(filePath):
        content=read_txt_16(filePath)
        if "IgnoreAllFailures" in content:
            passlist.append("check BCD:pass:%s"%content.replace("\n",""))
            logs("BCD Check Result:PASS")
        else:
            faillist.append("check BCD:fail:%s"%content.replace("\n",""))
            logs("%s:BCD Check Result:(Content Error)FAIL")
    else:
        faillist.append("check BCD:fail:Z:\\bcd.txt not exist!!")
        logs("BCD Check Result:(File not exist)FAIL")
def check_hiberboot():
    '''
    Check hiberboot option in registry:
    HiberbootEnabled=0
    '''
    rootpath = winreg.HKEY_LOCAL_MACHINE
    regPath="SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    if check_registry_exist(rootpath,regPath,"HiberbootEnabled"):
        regValue=read_reg(rootpath,regPath,"HiberbootEnabled")
        if regValue==0:
            passlist.append("check:Hiberboot:pass:HKEY_LOCAL_MACHINE\%s\RPSessionInterval=%d"%(regPath,regValue))
            logs("Hiberboot Check Result:PASS")
        else:
            faillist.append("check Hiberboot:fail:HKEY_LOCAL_MACHINE\%s\RPSessionInterval=%d" % ( regPath, regValue))
            logs("Hiberboot Check Result:(Value Error)FAIL")
    else:
        faillist.append("check Hiberboot:fail:Registry value not exist!!")
        logs("Hiberboot Check Result:(Reg not exist)FAIL")
def check_system_protection():
    rootpath = winreg.HKEY_LOCAL_MACHINE
    regPath="SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
    if check_registry_exist(rootpath,regPath,"RPSessionInterval"):
        regValue=read_reg(rootpath,regPath,"RPSessionInterval")
        if regValue==0:
            passlist.append("check system protection:pass:HKEY_LOCAL_MACHINE\%s\RPSessionInterval=%d"%(regPath,regValue))
            logs("System_Protection Check Result:PASS")
        else:
            faillist.append("check system protection:fail:HKEY_LOCAL_MACHINE\%s\RPSessionInterval=%d" % ( regPath, regValue))
            logs("System_Protection Check Result:(Value Error)FAIL")
    else:
        faillist.append("check System protection:fail:Registry value not exist!")
        logs("System_Protection Check Result:(Reg not exist)FAIL")
def check_timezone(language):
    try:
        f = open(get_current_dir()+"\\data\\timezone.yml", 'rb')
        y = yaml.load(f)
        actual_timezone = get_timezone()[0]
        os = get_os_type()
        if language in y["normal language"]:
            expected_timezone = y["normal language"][language]
            if actual_timezone == expected_timezone:
                logs("check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone)
                result = "check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone
                passlist.append(result)
            else:
                logs("check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone))
                result = "check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone)
                faillist.append(result)
        elif language in y["special language"]["care about OS"]:
            if os in y["special language"]["care about OS"][language]:
                expected_timezone = y["special language"]["care about OS"][language][os]
                if actual_timezone == expected_timezone:
                    logs("check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone)
                    result = "check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone
                    passlist.append(result)
                else:
                    logs("check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone))
                    result = "check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone)
                    faillist.append(result)
            else:
                expected_timezone = y["special language"]["care about OS"][language]["else"]
                if actual_timezone == expected_timezone:
                    logs("check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone)
                    result = "check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone
                    passlist.append(result)
                else:
                    logs("check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone))
                    result = "check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone)
                    faillist.append(result)
        elif language in y["special language"]["not care about OS"]:
            expected_timezone = y["special language"]["not care about OS"][language]
            if actual_timezone in expected_timezone:
                logs("check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone)
                result = "check timezone:pass:actual_timezone match expected_timezone:%s" % expected_timezone
                passlist.append(result)
            else:
                logs("check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone))
                result = "check timezone:fail:actual_timezone:%s doesn't match expected_timezone:%s" % (actual_timezone, expected_timezone)
                faillist.append(result)
        else:
            logs("check timezone:fail:The specified language can't be found.")
            result = "check timezone:fail:The specified language can't be found."
            faillist.append(result)
    except Exception as e:
        print(e)
def check_ime(language):
    try:
        f = open(get_current_dir() + "\\data\\IME.yml", 'rb')
        y = yaml.load(f)
        if language in y:
            flag = 0
            for key in y[language]:
                if check_file_exist(y[language][key]):
                    flag = 1
                else:
                    flag = 0
            if flag == 1:
                logs("Check IME:pass:The current language is %s" % language)
                result = "Check IME:pass:The current language is %s" % language
                passlist.append(result)
            elif flag == 0:
                logs("Check IME:fail:The current language is %s: not found"% language)
                result = "Check IME:fail:The current language is %s: not found" % language
                faillist.append(result)
        else:
            logs("Check IME:N/A:The current language is %s, don't need to check IME" % language)
    except Exception as e:
        print(e)
def check_event_viewer():
    filepath1 = "z:\\1.txt"
    filepath2 = "z:\\2.txt"
    filepathbcd = "z:\\bcd.txt"
    ######## 为了能正常执行bcdedit.exe,通过下面命令将windows\system32下面的文件copy到windows\\syswow64目录下
    ######## 备注：win32 python 在64bit os上windows\system32会默认指向到windows\syswow64目录，通过sysnative别名可以访问到system32
    ######## 代码中体现：windows\system32代表实际的windows\syswow64,代码中windows\sysnative实际代表windows\system32）
    copy_file("c:\\windows\\sysnative\\bcdedit.exe", "c:\\windows\\system32")
    if check_file_exist(filepath1):
        delete_file(filepath1)
    if check_file_exist(filepath2):
        delete_file(filepath2)
    if check_file_exist(filepathbcd):
        delete_file(filepathbcd)
    copy_file(get_current_dir()+r"\data\eventviewer.ps1", "z:\\")
    run_powershell("z:\\eventviewer.ps1")
    # while not(check_file_exist(filepath1) and check_file_exist(filepath2) and check_file_exist(filepathbcd)):
    #     time.sleep(1000)
    #     break
    # msgbox("Start to check event view.", "Information", "ok")
    EventsCount = 0
    if check_file_exist(filepath1):
        with open(filepath1, "r", encoding="utf-16") as f1:
            data1 = f1.read()
            failcount=0
            if data1:
                errors=data1.split("\n")
                for i in errors:
                    if i!="\n":
                        if i!="":
                            print("i:"+i)
                            if i.split("_")[0].split(":")[1].isnumeric():
                                print("check event viewer:fail:"+i+":critical")
                                faillist.append("check event viewer:fail:"+i+":critical")
                                failcount=failcount+1
            else:
                print("check event viewer:fail:no data")
                faillist.append("check event viewer:fail:no data")
                failcount = failcount + 1
            if failcount==0:
                print("check event viewer critical:pass")
                passlist.append("check event viewer critical:pass")
    if check_file_exist(filepath2):
        with open(filepath2, "r", encoding="utf-16") as f2:
            data2 = f2.read()
            failcount = 0
            if data2:
                errors=data2.split("\n")
                for i in errors:
                    if i != "\n":
                        if i != "":
                            if i.split("_")[0].split(":")[1].isnumeric():
                                print("check event viewer:fail:"+i+":error")
                                faillist.append("check event viewer:fail:"+i+":error")
                                failcount = failcount + 1
            else:
                print("check event viewer:fail:no data")
                faillist.append("check event viewer:fail:no data")
                failcount = failcount + 1
            if failcount==0:
                print("check event viewer error:pass")
                passlist.append("check event viewer error:pass")
def loadDataSet():
    filename = get_current_dir()+r'\data\display_language&location&format.txt'
    with open(filename,encoding="utf-8") as file_project:
        List_row = file_project.readlines()
        List_sourse = []
        for list_line in List_row:
            list_line = list(list_line.strip().split(','))
            s = []
            for i in list_line:
                s.append(i)
            List_sourse.append(s)
        return List_sourse
#print(loadDataSet())
def check_display_language(ostype,language):
    try:
        if check_registry_exist(winreg.HKEY_CURRENT_USER,'Control Panel\Desktop','PreferredUILanguages') and \
                check_registry_exist(winreg.HKEY_CURRENT_USER, 'Control Panel\Desktop\MuiCached',
                                 'MachinePreferredUILanguages'):

            value1 = read_reg(winreg.HKEY_CURRENT_USER,'Control Panel\Desktop','PreferredUILanguages')
            value2 = read_reg(winreg.HKEY_CURRENT_USER,'Control Panel\Desktop\MuiCached','MachinePreferredUILanguages')
            if value1 == value2:
                value = value1
        else:
            check_registry_exist(winreg.HKEY_CURRENT_USER,'Control Panel\Desktop\MuiCached','MachinePreferredUILanguages')
            value = read_reg(winreg.HKEY_CURRENT_USER,'Control Panel\Desktop\MuiCached','MachinePreferredUILanguages')

 #       print(value)
        for i in range(61):
            if loadDataSet()[i][0] == ostype and loadDataSet()[i][1] == language:
#               print(i)
                if value[0] == loadDataSet()[i][2]:
                    result = 'Check default display language is ' + value[0] + ':pass'
                    print(result)
                    passlist.append(result)
                else:
                    result = 'Check default display language is ' + value[0] + \
                             ':fail, default display language should be '+ loadDataSet()[i][2]
                    print(result)
                    faillist.append(result)
    except:
         print("can't find default display language!!!" )
 # check location
def check_location(ostype,language):
    try:
        if check_registry_exist(winreg.HKEY_CURRENT_USER,'Control Panel\International','sCountry'):
            value = read_reg(winreg.HKEY_CURRENT_USER,'Control Panel\International','sCountry')
 #           print(value)
            for i in range(61):
                if loadDataSet()[i][0] == ostype and loadDataSet()[i][1] == language:
 #                   print(i)
                    if value == loadDataSet()[i][4]:
                        result = 'Check region is ' + value + ':pass'
                        print(result)
                        passlist.append(result)
                    else:
                        result = 'Check region is ' + value + ': fail, region should be ' + loadDataSet()[i][4]
                        print(result)
                        faillist.append(result)

    except:
        print("can't find location!!!")
# check format
def check_mui_format(ostype,language):
    try:
        if check_registry_exist(winreg.HKEY_CURRENT_USER,'Control Panel\International','Locale'):
            value = read_reg(winreg.HKEY_CURRENT_USER,'Control Panel\International','Locale')
            for i in range(61):
                if loadDataSet()[i][0] == ostype and loadDataSet()[i][1] == language:
#                    print(i)
                    if value == loadDataSet()[i][3]:
                        result = 'check format is ' + value + ':pass'
                        print(result)
                        passlist.append(result)
                    else:
                        result = 'check MUI format is ' + value + ':fail, MUI format should be ' + loadDataSet()[i][3]
                        print(result)
                        faillist.append(result)

    except:
        print("can't find MUI format")
def check_hardware(platform, config):
    try:

        os_type = get_os_type()
        # file_path = r'\\15.83.240.98/wes/precheck_py/rookie/Precheck-2018 Fall Refresh/ReleasedNote/'
        # file_path = r'D:/Mine/BaiduNetDisk/'
        file_path = get_current_dir() + '\\data\\ReleasedNote\\'
        common_file = (
                    file_path + platform + '\\Hardware\\HW-ReleasedNote-' + platform + '-' + os_type + ' - common.txt')
        with open(common_file)as f1:
            note_release_org = f1.readlines()
            del note_release_org[0]  # ??'Common list'
        with open(
                file_path + platform + '\\Hardware\\HW-ReleasedNote-' + platform + '-' + os_type + ' - ' + config + '.txt')as f2:
            note_special = f2.readlines()
            del note_special[0]  # ??'Config list'
        note_release_org.extend(note_special)
        local_info = get_all_device_drivers()
        note_release = []
        for ele in note_release_org:
            new_ele = ele.strip('\n')  # ??\n
            note_release.append(new_ele)
        sorted_local_info = sorted(local_info)  # ??????
        sorted_note_release = sorted(note_release)
        passlist_unfilter = []
        faillist_unfilter = []
        local_info_new = []
        note_info_new = []
        for driver in sorted_local_info:
            new_driver = driver.split(',')
            local_info_new.append(new_driver)  # ????
        for note in sorted_note_release:
            new_note = note.split(',')
            note_info_new.append(new_note)  # ????
        for i in range(len(local_info_new)):
            for j in range(len(note_info_new)):
                local_version = local_info_new[i][1]
                note_version = note_info_new[j][1]
                if (note_info_new[j][0] == local_info_new[i][0]) and (local_version == note_version):
                    pass_result = 'Check ' + local_info_new[i][
                        0] + ': pass: Local version,' + local_version + ' ReleasedNote version,' + note_version
                    passlist_unfilter.append(pass_result)
                elif (note_info_new[j][0] == local_info_new[i][0]) and (local_version != note_version):
                    fail_result = 'Check ' + local_info_new[i][
                        0] + ': fail: Local version,' + local_version + ' ReleasedNote version,' + note_version
                    faillist_unfilter.append(fail_result)
        local_devices_names = []
        release_notes_names = []
        for i in range(len(local_info_new)):
            local_devices_names.append(local_info_new[i][0])
        for j in range(len(note_info_new)):
            release_notes_names.append(note_info_new[j][0])
        for one in local_devices_names:
            if one not in release_notes_names:
                unexpected_result = 'Check ' + one + ': fail: Unexpected'
                faillist_unfilter.append(unexpected_result)
        for one in release_notes_names:
            if one not in local_devices_names:
                missing_result = 'Check ' + one + ':fail:  Not Found'
                faillist_unfilter.append(missing_result)

        for one_pass in passlist_unfilter:
            if one_pass not in passlist:
                passlist.append(one_pass)
        for one_fail in faillist_unfilter:
            if one_fail not in faillist:
                faillist.append(one_fail)
    except Exception:
        raise
    return None
def check_protected_file_exist(path):
    filepath, filename = os.path.split(path)
    if os.path.exists(filepath):
        file_list = os.listdir(filepath)
        if filename in file_list:
            return True
    else:
        print("Can't find the file: " + path)
        return False
def check_software(platform):
    passlist_unfilter = []
    faillist_unfilter = []

    try:
        os_type = get_os_type()
        # energy star ??:'C:/hp/DATA/EStar/energystar.ico', 'C:/hp/DATA/EStar/Estar.dll'
        if os_type == 'win10' and (check_protected_file_exist('C:/hp/DATA/EStar/Estar.dll') and
                                   check_protected_file_exist('C:/hp/DATA/EStar/energystar.ico')):
            found_estar = 'FOUND Energy Star'
            passlist_unfilter.append(found_estar)
        elif os_type == 'win10' and ((not check_protected_file_exist('C:/hp/DATA/EStar/energystar.ico')) or
                                     (not check_protected_file_exist('C:/hp/DATA/EStar/Estar.dll'))):
            missing_estar = 'MISSING Energy Star'
            faillist_unfilter.append(missing_estar)
        local_software = get_installed_software_version()

        local_software_win10 = ['C:/Program Files/internet explorer/iexplore.exe',
                                'C:/Windows/SysNative/Macromed/Flash/FlashUtil_ActiveX.exe',
                                'C:/Windows/SysNative/AutoLogCfg.exe',
                                'C:/Program Files/Windows Media Player/wmplayer.exe',
                                'C:/Program Files (x86)/Citrix/ICA Client/Receiver/Receiver.exe',
                                ]
        local_software_wes7p = ['C:/Program Files/internet explorer/iexplore.exe',
                                'C:/Windows/SysNative/AutoLogCfg.exe',
                                'C:/Program Files/Windows Media Player/wmplayer.exe',
                                'C:/Program Files (x86)/Citrix/ICA Client/Receiver/Receiver.exe'
                                ]
        local_software_wes7e = ['C:/Program Files/internet explorer/iexplore.exe',
                                'C:/Windows/System32/AutoLogCfg.exe',
                                'C:/Program Files/Windows Media Player/wmplayer.exe',
                                'C:/Program Files/Citrix/ICA Client/Receiver/Receiver.exe'
                                ]
        if os_type == 'win10':
            local_software_single = local_software_win10
        elif os_type == 'wes7p':
            local_software_single = local_software_wes7p
        elif os_type == 'wes7e':
            local_software_single = local_software_wes7e
        for one in local_software_single:
            if check_protected_file_exist(one):
                single_software = get_single_software_version(one)
                local_software.append(single_software)
        # ?? win10?RDP??(get_single_software_version()?????????RDP??)
        rdp = os.popen('wmic datafile where name="C:\\\\Windows\\\\system32\\\\mstsc.exe" get version').read()
        rdp_info = 'Remote Desktop Connection,' + rdp.strip().split('\n')[-1].strip()
        local_software.append(rdp_info)  # ??RDP????
        sorted_local_software = sorted(local_software)  # ???????
        # file_path = r'\\15.83.240.98/wes/precheck_py/rookie/Precheck-2018 Fall Refresh/ReleasedNote/'
        file_path = get_current_dir() + '\\data\\ReleasedNote\\'
        sw_file = file_path + platform + '\\Software\\SW-ReleasedNote-' + platform + '-' + os_type + '.txt'
        with open(sw_file, 'r')as f1:
            sw = f1.readlines()
        note_software = []
        for one in sw:
            new_one = one.strip('\n')  # ??\n
            note_software.append(new_one)
        sorted_note_software = sorted(note_software)  # ???????
        local_software_new = []
        note_software_new = []
        for one in sorted_local_software:
            new_local = one.split(',')
            local_software_new.append(new_local)  # ????
        for one in sorted_note_software:
            new_note = one.split(',')
            note_software_new.append(new_note)  # ????
        for i in range(len(local_software_new)):
            for j in range(len(note_software_new)):
                local_version = local_software_new[i][1]
                local_name = local_software_new[i][0]
                note_version = note_software_new[j][1]
                note_name = note_software_new[j][0]
                if local_name == note_name and local_version == note_version:
                    pass_result = 'Check ' + local_name + ': pass: Local version,' + local_version + ' Releasednote version,' + note_version
                    passlist_unfilter.append(pass_result)
                elif local_name == note_name and local_version != note_version:
                    fail_result = 'Check ' + local_name + ': fail: Local version,' + local_version + ' Releasednote version,' + note_version
                    faillist_unfilter.append(fail_result)
        local_software_names = []
        notes_software_names = []
        for i in range(len(local_software_new)):
            local_software_names.append(local_software_new[i][0])
        for j in range(len(note_software_new)):
            notes_software_names.append(note_software_new[j][0])
        for one_local in local_software_names:
            if one_local not in notes_software_names:
                unexpected_result = 'Check ' + one_local + ':fail: Unexpected'
                faillist_unfilter.append(unexpected_result)
        for one_note in notes_software_names:
            if one_note not in local_software_names:
                missing_result = 'Check ' + one_note + ':fail:  Not Found'
                faillist_unfilter.append(missing_result)

        # ???
        for one_pass in passlist_unfilter:
            if one_pass not in passlist:
                passlist.append(one_pass)
        for one_fail in faillist_unfilter:
            if one_fail not in faillist:
                faillist.append(one_fail)
    except Exception:
        raise
    return None
def generate_project_kb(platform):
    """
    capture the new project installed KBs and write to kb ReleasedNote txt
    :return:
    """
    first_ml_kb = get_kb()
    os_type = get_os_type()
    file_path = get_current_dir() + '\\data\\ReleasedNote\\'
    kb_file = file_path + platform + '\\KB\\KB-ReleasedNote-' + platform + '-' + os_type + '.txt'
    for one in first_ml_kb:
        write_txt(kb_file, one+'\n')
    return None
def check_kb(platform):
    try:
        os_type = get_os_type()
        file_path = get_current_dir() + '\\data\\ReleasedNote\\'
        kb_file = file_path + platform + '\\KB\\KB-ReleasedNote-' + platform + '-' + os_type + '.txt'
        if not check_protected_file_exist(kb_file):
            generate_project_kb(platform)
            fail_log = "new KB_releasenote created under " + kb_file + ", please check all it's correct, then run precheck again!"
            faillist.append(fail_log)
        else:
            local_kb_org = get_kb()
            local_kb = []
            for one in local_kb_org:
                if one in local_kb:
                    new_one = one + 'p'
                    local_kb.append(new_one)
                elif one not in local_kb:
                    local_kb.append(one)
            with open(kb_file, 'r')as f1:
                kb = f1.readlines()
            note_kb = []
            for one in kb:
                new_one = one.strip('\n')  # ??\n
                note_kb.append(new_one)
            for one in local_kb:
                if one not in note_kb:
                    unexpected_result = 'Check ' + one + ': fail: Unexpected'
                    faillist.append(unexpected_result)
                else:
                    pass_result = 'Check ' + one + ': pass'
                    passlist.append(pass_result)
            for one in note_kb:
                if one not in local_kb:
                    missing_result = 'Check ' + one + ': fail: Not Found'
                    faillist.append(missing_result)
    except Exception:
        raise
    return None
def check_BSOD():
    """
    还可以通过读取注册表里RunTimeID的值判断系统类型
    :return:
    """

    try:
        bsod=read_reg(winreg.HKEY_LOCAL_MACHINE, 'SYSTEM\CurrentControlSet\Control\CrashControl','AutoReboot')
       # key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,SYSTEM\CurrentControlSet\Control\CrashControl\AutoReboot,0, winreg.KEY_ALL_ACCESS)
        # 当查询注册表最后一级目录时，QueryValueEx函数返回的是所有键所对应的值，类型是列表
        #reg_ml = winreg.QueryValueEx(key, AutoReboot)
    except Exception:
        raise
    if int(bsod) == 0:
        passlist.append("Check BSOD  not auto reboot:Pass,RegValue:{}".format(bsod))
        logs("Check BSOD  not auto reboot:Pass")
    else:
        faillist.append("Check BSOD  not auto reboot:Fail, RegValue:{}".format(bsod))
        logs("Check BSOD  not auto reboot:(value Error)Fail")
def check_pagefile():
    Owmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    colPageFiles = Owmi.ExecQuery("Select * from Win32_PageFileUsage")
    if colPageFiles.count != 0:
        logs("check page file:fail:colPageFiles:%s"%colPageFiles[0].Name)
        result = "check page file:fail:colPageFiles:%s"%colPageFiles[0].Name
        print(result)
        faillist.append(result)
    else:
        logs("check page file:pass, Pagefile count:{}".format(colPageFiles.count))
        result = "check page file:pass, Pagefile count:{}".format(colPageFiles.count)
        print(result)
        passlist.append(result)
def check_windows_update():
    c = wmi.WMI()
    for service in c.Win32_Service():
        if service.Caption == "Windows Update":
            print(service.Caption, service.StartMode, service.State)
            if service.State == "Stopped" and service.StartMode == "Disabled":
                logs("check windows update:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check windows update:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                passlist.append(result)
            else:
                logs("check windows update:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check windows update:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                faillist.append(result)
def check_sms():
    c = wmi.WMI()
    for service in c.Win32_Service():
        if service.Caption == "SMS Agent Host":
            print(service.Caption, service.StartMode, service.State)
            if service.State == "Stopped" and service.StartMode == "Manual":
                logs("check sms:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check sms:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                passlist.append(result)
            else:
                logs("check sms:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check sms:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                faillist.append(result)
def check_windows_search():
    c = wmi.WMI()
    for service in c.Win32_Service():
        if service.Caption == "Windows Search":
            print(service.Caption, service.StartMode, service.State)
            if service.State == "Stopped":
                logs("check windows search:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check windows search:pass:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                passlist.append(result)
            else:
                logs("check windows search:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode))
                result = "check windows search:fail:service.State:%s:service.StartMode:%s"%(service.State,service.StartMode)
                faillist.append(result)
def check_service():
    try:
        check_windows_update()
        check_sms()
        check_windows_search()
    except:
        print("fail")
def load_data_set_languagepackages():
    filename = get_current_dir()+r'\data\check_language_packages.txt'
    with open(filename,encoding="utf-8") as file_project:
        List_row = file_project.readlines()
        List_source = []
        for line in List_row:
            list_line = list(line.strip().split(','))
            s = []
            for i in list_line:
                s.append(i)
            List_source.append(s)
        return List_source
# print(load_data_set())

def check_language_packages(ostype, language):

    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\MUI\UILanguages")
    try:
        x = 0
        lists = []
        while True:
            subkey = winreg.EnumKey(key, x)
            # print(subkey)
            lists.append(subkey)   # language packages subkeys list
            x = x + 1
    except WindowsError:
        pass
    # print("subkeylist: ", lists)
    character = ''
    for l in lists:
        character = character + l + '|'   # combine items in lists to a long string, separated with '|'
    # print("character", character)

    count = x
    pass_count = 0

    with open(get_current_dir()+r'\data\check_language_packages.txt', encoding="utf-8") as f:
        rows = len(f.readlines())
        # print("rows: ", rows)
    for n in range(rows):
        type = load_data_set_languagepackages()[n][0]
        lan = load_data_set_languagepackages()[n][1]
        coun = load_data_set_languagepackages()[n][5]
        package = load_data_set_languagepackages()[n][6]
        if ostype in type and language in lan:
            if int(coun) == count:
                for item in lists:
                    if item in package:
                        # print('{},{},{},{}'.format(type, lan, coun, package))
                        pass_count = pass_count + 1
                if pass_count == count:
                    result = "check language packages passed. The package is: " + character
                    print(result)
                    passlist.append(result)
                else:
                    result = "check language packages failed. The actual package is: " + character +\
                             " the expected value should be " + load_data_set_languagepackages()[n][6]
                    print(result)
                    faillist.append(result)
            else:
                result = "check language packages count failed, the count is: " + str(count) + \
                         " the expected count should be: "+ str(coun)
                print(result)
                faillist.append(result)

def load_data_set_keyboardlayouts():
    filename = get_current_dir()+r'\data\check_keyboard_layouts.txt'
    with open(filename,encoding="utf-8") as file_project:
        List_row = file_project.readlines()
        List_source = []
        for line in List_row:
            list_line = list(line.strip().split(','))
            s = []
            for i in list_line:
                s.append(i)
            List_source.append(s)
        return List_source
# print(load_data_set_keyboardlayouts())

def check_keyboard_layouts(ostype, language):
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Keyboard Layout\Preload")
    try:
        lists = []
        y = 0
        while True:
            # name, data, type = winreg.EnumValue(key, y)
            t = winreg.EnumValue(key, y)
            # print(t)
            lists.append(t)
            y = y + 1
    except WindowsError:
        pass
    # print(lists)
    character = ''
    for l in lists:
        valuestring = l[0]+','+l[1]
        character = character + valuestring + '|'   # combine items in lists to a long string, separated with '|'
    # print("character: ", character)

    count = len(lists)
    pass_count = 0

    with open(get_current_dir()+r'\data\check_keyboard_layouts.txt', encoding="utf-8") as f:
        rows = len(f.readlines())
        # print("rows: ", rows)
    for n in range(rows):
        if ostype in load_data_set_keyboardlayouts()[n][0] and language in load_data_set_keyboardlayouts()[n][1]:
            if int(load_data_set_keyboardlayouts()[n][2]) == count:
                for item in lists:
                    actualvalue = item[0]+item[1]
                    # print(actualvalue)
                    if actualvalue in load_data_set_keyboardlayouts()[n][3]:
                        pass_count = pass_count + 1
                if pass_count == count:
                    result = "check keyboard layouts passed. The layout is: " + character
                    print(result)
                    passlist.append(result)
                else:
                    result = "check keyboard layouts failed. The actual layouts is: " + character + \
                             " the expected value should be one of " + load_data_set_keyboardlayouts()[n][3]
                    print(result)
                    faillist.append(result)
            else:
                result = "check keyboard layouts count failed, the count is: " + str(count) + \
                         " the expected count should be: " + str(load_data_set_keyboardlayouts()[n][2])
                print(result)
                faillist.append(result)
if __name__ == '__main__':
    check_qfe()
    #check_image_info()
    print(passlist)
    print(faillist)

