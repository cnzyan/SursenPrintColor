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
# pyinstaller --uac-admin -F sursen_enhancer_setup.py --icon=sursen.ico

'''
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
def get_program_files_dir():
    # Prefer 64-bit Program Files on 64-bit Windows
    if 'PROGRAMFILES(X86)' in os.environ:
        return os.environ['PROGRAMFILES']
    return os.environ.get('PROGRAMFILES', r'C:\Program Files')

def copy_exe_files(src_dir, dest_dir):
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    for exe_file in glob.glob(os.path.join(src_dir, '*.exe')):
        shutil.copy2(exe_file, dest_dir)

def run_srpc_cli(dest_dir):
    srpc_path = os.path.join(dest_dir, 'srpc_cli.exe')
    if os.path.exists(srpc_path):
        subprocess.run([srpc_path], shell=True)

def set_gw_file_association(dest_dir):
    opener_path = os.path.join(dest_dir, 'sursen_opener.exe')
    if not os.path.exists(opener_path):
        print("sursen_opener.exe not found, skipping file association.")
        return

    # Set .gw file association in registry
    try:
        # 1. Set .gw to sursen_gwfile
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, '.gw') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, 'sursen_gwfile')

        # 2. Set sursen_gwfile shell open command
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r'sursen_gwfile\shell\open\command') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, f'"{opener_path}" "%1"')

        # 3. Set icon (optional)
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r'sursen_gwfile\DefaultIcon') as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, f'"{opener_path}",0')

        # 4. Notify Windows of the change
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        print(".gw file association set successfully.")
    except Exception as e:
        print(f"Failed to set file association: {e}")

def main():
    src_dir = os.path.dirname('./')
    dest_dir = os.path.join(get_program_files_dir(), 'sursen_enhancer_by_crazyan')

    print(f"Copying exe files from {src_dir} to {dest_dir} ...")
    copy_exe_files(src_dir, dest_dir)

    print("Running srpc_cli.exe ...")
    run_srpc_cli(dest_dir)

    print("Setting .gw file association ...")
    set_gw_file_association(dest_dir)

    print("Installation complete.")
    wait_for_any_key()

if __name__ == '__main__':
    if not sys.platform.startswith('win'):
        print("This installer only supports Windows.")
        sys.exit(1)
    main()