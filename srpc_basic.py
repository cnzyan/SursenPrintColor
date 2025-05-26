# 读取注册表
# .\.venv\Scripts\activate.ps1
# pyinstaller -F srpc_basic.py
import winreg
import os
import ctypes
import sys
def wait_for_any_key():
    """等待用户按下任意键继续（跨平台实现）"""
    if os.name == 'nt':
        # Windows系统使用msvcrt
        import msvcrt
        print("\n按任意键继续...", end='', flush=True)
        msvcrt.getch()
    else:
        # 类Unix系统（Linux/MacOS）使用termios
        import tty
        import termios
        print("\n按任意键继续...", end='', flush=True)
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
    print()  # 换行保持格式整洁
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

def add_runas(cmd_path):
    '''添加管理员权限到注册表'''
    '''
    exe_path = sys.executable
    # 判断当前运行的Python是否具有管理员权限，没有则申请
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, __file__, None, 1)
    '''
    reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
    reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, access=winreg.KEY_SET_VALUE | winreg.KEY_READ)
    runas_value = "~ RUNASADMIN"
    try:
        value = winreg.QueryValueEx(reg_key, cmd_path)
    except FileNotFoundError:
        winreg.SetValueEx(reg_key, cmd_path, 0, winreg.REG_SZ, runas_value)
    else:
        if runas_value[2:] not in value[0]:
            winreg.SetValueEx(reg_key, cmd_path, 0, winreg.REG_SZ, value[0] + ' ' + runas_value[2:])
    print(f"Added runas administrator done to {cmd_path}")
    winreg.CloseKey(reg_key)
def get_registry_value(key_path, value_name):
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
        value, _ = winreg.QueryValueEx(key, value_name)
        winreg.CloseKey(key)
        return value
    except Exception as e:
        print(f"Error getting registry value: {e}")
        return None

if __name__ == "__main__":
    isadmin=run_as_admin()
    if not isadmin:
        print("Trying to run this program as administrator.")
        os._exit(0)
    print("Sursen Reader Print Color Fixer by Zhai @ 20250124")
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
        set_runas_admin = False
        if set_runas_admin:
            add_runas(cmd_path)
            print(f"Added runas administrator done to {cmd_path}")
        else:
            print(f"Not added runas administrator done to {cmd_path}")

        # print(folder)
        try:
            content=''
            with open(folder+'\\SursenPrint.ini','r') as f:
                for line in f:
                    if line.startswith('PrintBlackAndWhite='):
                        line='PrintBlackAndWhite=0\n'
                    content+=line
            with open(folder+'\\SursenPrint.ini','w') as f:
                f.write(content)
            print(f"Fixed SursenPrint.ini for {folder}") 
        except:
            print("FAIL to open file in system folder. what a fail. You need run with administrator privileges.")
    for dir in os.listdir("C:\\Users\\"):
        # print(dir)
        
        try:
            content=''
            with open("C:\\Users\\"+dir+'\\AppData\\Local\\VirtualStore\\Program Files (x86)\\Sursen\\Reader\\SursenPrint.ini','r') as f:
                for line in f:
                    if line.startswith('PrintBlackAndWhite='):
                        line='PrintBlackAndWhite=0\n'
                    content+=line
            with open("C:\\Users\\"+dir+'\\AppData\\Local\\VirtualStore\\Program Files (x86)\\Sursen\\Reader\\SursenPrint.ini','w') as f:
                f.write(content)
            print(f"Fixed SursenPrint.ini for {dir}")
        except:
            pass
    wait_for_any_key()