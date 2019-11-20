import os,win32con
import subprocess
import  shutil
from win32timezone import *
import uuid
from win32com.client import GetObject

def read_txt(path):
    try:
        with open(path, 'r', encoding='utf-8') as file_object:
            contents = file_object.read()
            #print(contents)
            return contents
    except:
        print('Read txt file fail')
def write_txt(filename, str):
    try:
        with open(filename, 'a', encoding="utf-8") as file_object:
            file_object.write(str)
    except:
        print('write file fail')
def create_txt(filename):
    try:
        open(filename, 'w').close()
    except:
        print('Create file fail')
def get_local_time():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
def check_file_exist(file):
    return os.path.exists(file)
def rename_folder(src, dst):
    try:
        os.rename(src, dst)
        return True
    except:
        return False
def rename_file(src, dst):
        try:
            os.rename(src, dst)
            return True
        except:
            return False
def copy_file(src, dst):
    try:
        shutil.copy(src, dst)
        return True
    except:
        return False
def delete_file(file):
    try:
        os.remove(file)
        return True
    except:
        return False
def create_folder(path):
    try:
        os.makedirs(path)
        return True
    except:
        return False
def copy_folder(src, dst):
    try:
        shutil.copytree(src, dst)
        return True
    except:
        return False
def delete_folder(path):
    try:
        shutil.rmtree(path)
        return True
    except:
        return False
def run_cmd(command):
    return os.popen(command).read()
def run_exe(command):
    return os.popen(command).read()
def run_powershell(command):
    args = [r"powershell", "-ExecutionPolicy", "Unrestricted"]
    if type(command)==list:
        command=args+command
    else:
        command=args+[command]
    return subprocess.Popen(command, stdout=subprocess.PIPE).stdout.read()
def msgbox(text, title, t="ok"):
    if t.lower() == "error":
        type = win32con.MB_ICONERROR
    if t.lower() == "warning":
        type = win32con.MB_ICONWARNING
    if t.lower() == "ok":
        type = win32con.MB_OK
    win32api.MessageBox(0, text, title, type)
def get_current_dir():
    return os.getcwd()
def get_local_time():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def get_mac():
    try:
      node = uuid.getnode()
      mac = uuid.UUID(int = node).hex[-12:]
      return mac
    except:
        print("get mac error")
        return None
def check_single_service(serviceName):
    wmi = GetObject('winmgmts:/root/cimv2')
    runningServices = wmi.ExecQuery("Select * from Win32_Service")
    for service in runningServices:
        sName = service.Caption
        sStartMode = service.StartMode
        sState = service.State
        if sName == serviceName:
            return sName, sStartMode, sState
            break
def get_kb():
    wmi = GetObject('winmgmts:/root/cimv2')
    kbs = wmi.ExecQuery("Select * from Win32_QuickFixEngineering")
    lists = []
    for kb in kbs:
	    lists.append(kb.HotfixID)
    return lists
def get_timezone():
    wmi = GetObject('winmgmts:/root/cimv2')
    timezones = wmi.ExecQuery('Select * from Win32_TimeZone')
    lists = []
    for timezone in timezones:
        lists.append(timezone.Caption)
    return lists
def create_registry(rootpath,path, valuename):
    reg_root =rootpath
    reg_path = path
    reg_flags = win32con.WRITE_OWNER | win32con.KEY_WOW64_64KEY | win32con.KEY_ALL_ACCESS
    try:
        key = win32api.RegOpenKeyEx(reg_root, reg_path, 0, reg_flags)
        win32api.RegSetValueEx(key, valuename, 0, win32con.REG_SZ, '')
        win32api.RegCloseKey(key)
        return True
    except:
        return False
def delete_registry(rootpath,path, valuename):
    reg_root = rootpath
    reg_path = path
    reg_flags = win32con.WRITE_OWNER | win32con.KEY_WOW64_64KEY | win32con.KEY_ALL_ACCESS
    try:
        # 删除value
        key = win32api.RegOpenKeyEx(reg_root, reg_path, 0, reg_flags)
        win32api.RegDeleteValue(key, valuename)
        win32api.RegCloseKey(key)
        return True
    except:
        return False
def read_reg(rootpath,path,name):
    try:
        registry_key = winreg.OpenKey(rootpath, path, 0,
                                       winreg.KEY_READ)
        value, regtype = winreg.QueryValueEx(registry_key, name)
        winreg.CloseKey(registry_key)
        return value
    except WindowsError:
        return None
def check_registry_exist(rootpath,path,name):
    try:
        registry_key = winreg.OpenKey(rootpath, path, 0,
                                       winreg.KEY_READ)
        value, regtype = winreg.QueryValueEx(registry_key, name)
        winreg.CloseKey(registry_key)
        return True
    except WindowsError:
        return False
def edit_registry(rootpath,path, valuename, value):
    reg_root = rootpath
    reg_path = path
    reg_flags = win32con.WRITE_OWNER | win32con.KEY_WOW64_64KEY | win32con.KEY_ALL_ACCESS

    if check_registry_exist(path, valuename):
        key = win32api.RegOpenKeyEx(reg_root, reg_path, 0, reg_flags)
        win32api.RegSetValueEx(key, valuename, 0, win32con.REG_SZ, value)
        win32api.RegCloseKey(key)
        return True
    else:
        return False
def rename_reg(rootpath,path,inname,outname):
    if(check_registry_exist(rootpath,path,inname)):
        temp=read_reg(rootpath,path,inname)
        delete_registry(rootpath,path,inname)
        create_registry(rootpath,path,outname,temp)
    else:
        print("not found value name")
def get_os_type():

    """
    Win32_OperatingSystem???????,??:
    Caption,OSArchitecture,Version,CodeSet,OSLanguage,OSType,SerialNumber,MUILanguages
    Status,CountryCode,CurrentTimeZone,ServicePackMajorVersion,ServicePackMinorVersion,Locale
    :return:
    """
    try:
        wmi = GetObject('winmgmts:/root/cimv2')
        operating_systems = wmi.ExecQuery("Select * from Win32_OperatingSystem ")
    except Exception as e:
        print("Can't query os from Win32_OperatingSystem\n", e)
    for o_s in operating_systems:
        os_name = o_s.Caption
        os_bit = o_s.OSArchitecture
    os_type = 'incorrect os type'
    if 'Microsoft Windows 10' in os_name:
        os_type = 'win10'
    elif 'Microsoft Windows Embedded Standard' in os_name and '64' in os_bit:
        os_type = 'wes7p'
    elif 'Microsoft Windows Embedded Standard' in os_name and '32' in os_bit:
        os_type = 'wes7e'
    return os_type
def get_single_software_version(filepath):
    """
    Read all properties of the given file and return them as a dictionary.
    ??????????????????(???props)
    :param filepath:
    :return file_info:
    """

    StringFileInfo_structure = ('Comments', 'InternalName', 'ProductName',
                                'CompanyName', 'LegalCopyright', 'ProductVersion',
                                'FileDescription', 'LegalTrademarks', 'PrivateBuild',
                                'FileVersion', 'OriginalFilename', 'SpecialBuild')

    props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None}
    try:
        # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
        # ???????
        VS_FIXEDFILEINFO_structure = win32api.GetFileVersionInfo(filepath, '\\')
        props['FixedFileInfo'] = VS_FIXEDFILEINFO_structure
        props['FileVersion'] = "%d.%d.%d.%d" % (VS_FIXEDFILEINFO_structure['FileVersionMS'] / 65536,
                                                VS_FIXEDFILEINFO_structure['FileVersionMS'] % 65536,
                                                VS_FIXEDFILEINFO_structure['FileVersionLS'] / 65536,
                                                VS_FIXEDFILEINFO_structure['FileVersionLS'] % 65536)
        # ??????????
        Major_Version_Number = win32api.HIWORD(VS_FIXEDFILEINFO_structure['FileVersionMS'])
        Minor_Version_Number = win32api.LOWORD(VS_FIXEDFILEINFO_structure['FileVersionMS'])
        Revision_Number = win32api.HIWORD(VS_FIXEDFILEINFO_structure['FileVersionLS'])
        Build_Number = win32api.LOWORD(VS_FIXEDFILEINFO_structure['FileVersionLS'])
        version = '{}.{}.{}.{}'.format(Major_Version_Number, Minor_Version_Number, Revision_Number, Build_Number)

        # \VarFileInfo\Translation returns list of available (language, codepage)
        # pairs that can be used to retrieve string info. We are using only the first pair.
        # ??????????????,?(2052, 1200), 2052????,1200??Unicode
        lang, codepage = win32api.GetFileVersionInfo(filepath, '\\VarFileInfo\\Translation')[0]

        # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
        # two are language/codepage pair returned from above
        StringFileInfo_dict = {}
        for item in StringFileInfo_structure:
            strInfopath = '\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, item)
            StringFileInfo_dict[item] = win32api.GetFileVersionInfo(filepath, strInfopath)
        props['StringFileInfo'] = StringFileInfo_dict
    except Exception as e:
        print(e, '\n', filepath)
    if not props["StringFileInfo"]:
        return None
    else:
        name = props["StringFileInfo"]["FileDescription"]
        if 'AutoLogon' in name:
            name = 'HP Logon Manager'
        if 'Adobe' in name:
            name = 'Adobe Flash Player'
        if filepath == 'C:/Windows/SysNative/mstsc.exe' or filepath == 'C:/Windows/System32/mstsc.exe':
            name = 'Remote Desktop Connection'
        if filepath == 'C:/Program Files/Windows Media Player/wmplayer.exe':
            name = 'Windows Media Player'
        version = props['FileVersion']
        file_info = name+','+version
        return file_info

def get_all_abnormal_devices():
    """
    ??ConfigManagerErrorCode????0,?device??
    :return: ?????????????
    """
    try:
        wmi = GetObject('winmgmts:/root/cimv2')
        abnormal_devices = wmi.ExecQuery("Select * from Win32_PnpEntity Where "
                                         "ConfigManagerErrorCode != 0")
    except Exception as e:
        print("Can't query abnormal devices from Win32_PnpEntity\n", e)
    device_list = []
    for device in abnormal_devices:
        info = str(device.Name)
        device_list.append(info)
    return device_list


if __name__ == '__main__':
    msgbox("aa".lower(),"aa","error")
