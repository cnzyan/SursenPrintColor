import os
import shutil
import subprocess
import sys
import glob
import winreg
import ctypes
'''
使用python，编写一个安装程序，将本目录所有exe文件复制到%programfiles%下的sursen_enhancer_by_crazyan文件夹，然后运行修复程序srpc_cli.exe，将sursen_opener.exe设置为.gw格式文件的默认打开方式，支持windows7到win11的所有windows系统版本
# .\.venv\Scripts\activate.ps1
# pyinstaller --uac-admin -F sursen_enhancer_setup.py --icon=sursen.ico --name "Sursen Reader修复工具安装程序"

'''
def wait_for_any_key():
    """等待用户按下任意键继续（跨平台实现）"""
    if os.name == 'nt':
        import msvcrt
        print("\n按任意键继续...", end='', flush=True)
        msvcrt.getch()
    else:
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
    print()

def get_program_files_dir():
    if 'PROGRAMFILES(X86)' in os.environ:
        return os.environ['PROGRAMFILES']
    return os.environ.get('PROGRAMFILES', r'C:\Program Files')

def copy_exe_files(src_dir, dest_dir):
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    for exe_file in glob.glob(os.path.join(src_dir, '*.exe')):
        try:
            # 复制文件到目标目录
            shutil.copy2(exe_file, dest_dir)
        except Exception as e:
            pass  # 处理文件复制异常

def run_srpc_cli(dest_dir):
    srpc_path = os.path.join(dest_dir, 'srpc_cli.exe')
    if os.path.exists(srpc_path):
        subprocess.run([srpc_path], shell=True)

def set_gw_file_association(dest_dir):
    opener_path = os.path.join(dest_dir, 'sursen_opener.exe')
    if not os.path.exists(opener_path):
        print("sursen_opener.exe 未找到，跳过文件关联设置。")
        return
    main_key="sursen_gwfile"
    # main_key="SursenReader.gw"
    try:
        # 设置完整的文件类型信息
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, '.gw') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, main_key)
            winreg.SetValueEx(key, 'Content Type', 0, winreg.REG_SZ, 'application/sursen-gw')

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, main_key) as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, 'Sursen GW Document')
            winreg.SetValueEx(key, 'FriendlyTypeName', 0, winreg.REG_SZ, '书生 GW 文档')
            winreg.SetValueEx(key, 'EditFlags', 0, winreg.REG_DWORD, 0x00010000)

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, main_key+r'\DefaultIcon') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, f'"{opener_path}",0')

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, main_key+r'\shell\open\command') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, f'"{opener_path}" "%1"')
        print("成功设置 .gw 文件关联！")
    except Exception as e:
        print(f"设置文件关联失败: {str(e)}")
    try:
        # 设置用户可见的默认关联
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.gw\UserChoice') as key:
            winreg.SetValueEx(key, 'Progid', 0, winreg.REG_SZ, 'SursenReader.gw')

        # 刷新系统设置
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        print("成功设置 .gw 默认关联！")
    #except PermissionError:
    #    print("权限不足，请以管理员身份运行安装程序。")
    except Exception as e:
        print(f"设置默认关联失败: {str(e)}")

def main():
    src_dir = os.path.dirname('./')
    dest_dir = os.path.join(get_program_files_dir(), 'sursen_enhancer_by_crazyan')

    print(f"正在复制文件从 {src_dir} 到 {dest_dir}...")
    copy_exe_files(src_dir, dest_dir)

    print("正在运行 srpc_cli.exe...")
    run_srpc_cli(dest_dir)

    print("正在设置 .gw 文件关联...")
    set_gw_file_association(dest_dir)

    print("安装完成！")
    wait_for_any_key()

if __name__ == '__main__':
    if not sys.platform.startswith('win'):
        print("本安装程序仅支持 Windows 操作系统。")
        sys.exit(1)
    main()