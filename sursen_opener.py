import os
import sys
import shutil
import ctypes
import winreg
"""
Sursen Reader打开文件助手 CLI版本
作者：Zhai @ 20250519
版本：1.0
# .\.venv\Scripts\activate.ps1
# pyinstaller --uac-admin -F -w sursen_opener.py --icon=sursen.ico
"""
TEMP_DIR = r"C:\cratmp"
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
def run_as_admin():
    if is_admin():
        return True
    else:
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
            return False
        except:
            return False
def get_registry_value(key_path, value_name):
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
        value, _ = winreg.QueryValueEx(key, value_name)
        winreg.CloseKey(key)
        return value
    except Exception as e:
        print(f"Error getting registry value: {e}")
        return None
    
def get_short_filename(filepath):
    base = os.path.basename(filepath)
    name, ext = os.path.splitext(base)
    short_name = name[:32]
    candidate = f"{short_name}{ext}"
    i = 1
    while os.path.exists(os.path.join(TEMP_DIR, candidate)):
        candidate = f"{short_name}_{i}{ext}"
        i += 1
    return candidate

def open_with_sursen(filepath):
    # print("Opening with Sursen:", cmd_path,filepath)
    # 要执行的.exe文件路径
    exe_path = "start /B \""+cmd_path+"\" \""+ filepath+"\" >nul 2>&1"
    exe_path = cmd_path + " \"" + filepath + "\" "
    # os.system(exe_path)
    import subprocess

    # 使用subprocess.Popen执行.exe文件，并传入STARTUPINFO对象
    subprocess.Popen(exe_path)
def main():
    if len(sys.argv) < 2:
        print("请通过双击.gw文件或拖拽文件到本程序运行。")
        sys.exit(1)
    filepath = sys.argv[1]
    if not os.path.isfile(filepath):
        print("文件不存在:", filepath)
        sys.exit(1)
    print("文件路径:", filepath, "长度:", len(filepath))
    if len(filepath) <= 150:
        open_with_sursen(filepath)
    else:
        if not os.path.exists(TEMP_DIR):
            os.makedirs(TEMP_DIR)
        short_filename = get_short_filename(filepath)
        temp_path = os.path.join(TEMP_DIR, short_filename)
        shutil.copy2(filepath, temp_path)
        open_with_sursen(temp_path)

if __name__ == "__main__":
    print("Sursen Reader GW opener by Zhai @ 20250519")
    isadmin=run_as_admin()
    if os.path.exists(TEMP_DIR):
        pass
    else:
        os.makedirs(TEMP_DIR)
    for i in os.listdir(TEMP_DIR):
        if i.endswith(".gw"):
            try:
                os.remove(os.path.join(TEMP_DIR, i))
            except:
                pass
    key_path = r"SursenReader.gw\\shell\\open\\command"
    value_name = ""
    registry_value = get_registry_value(key_path, value_name)
    if registry_value:
        print(f"Found Sursen Installed : {registry_value}")
        
        i=0
        lc=0
        for c in registry_value:
            i+=1
            if c=="\\":
                lc=i
        folder=registry_value[:lc]
        cmd_path=folder+r"SursenReader.exe"
    main()