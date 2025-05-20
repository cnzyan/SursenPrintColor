import sys
import os
import threading
import shutil
import ctypes
import winreg
import subprocess
from ctypes import wintypes
import keyboard
"""
Sursen Reader打开文件助手 CLI版本
作者：Zhai @ 20250519
版本：2.0
更新内容：
1. 完全解决管理员会话路径问题
2. 跨会话UNC路径处理
3. 智能提权机制
# pyinstaller -F -w sursen_opener.py --icon=sursen.ico 
# pyinstaller -F sursen_opener.py --icon=sursen.ico --name "sursen_opener_dbg"
"""

TEMP_DIR = os.path.join(os.environ['APPDATA'], 'SursenTemp')
MAX_PATH = 150


def exit_program():
    print("\nTimeout reached. Exiting.")
    os._exit(0)  # 立即退出所有线程


def wait_for_any_key():
    # 创建并启动超时定时器
    print("\n按任意键继续...", end='', flush=True)
    timer = threading.Timer(10.0, exit_program)
    timer.daemon = True  # 设置为守护线程
    timer.start()
    """根据操作系统检测任意键输入"""
    if os.name == 'nt':
        # Windows 系统使用msvcrt
        import msvcrt
        while not msvcrt.kbhit():
            threading.Event().wait(0.1)
        msvcrt.getch()
    else:
        # Unix/Linux/Mac 使用termios
        import tty
        import termios
        # 保存原始终端设置
        old_settings = termios.tcgetattr(sys.stdin)
        try:
            tty.setcbreak(sys.stdin.fileno())
            sys.stdin.read(1)
        finally:
            # 恢复终端设置
            termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)


# Windows API结构体定义


class NETRESOURCE(ctypes.Structure):
    _fields_ = [
        ('dwScope', wintypes.DWORD),
        ('dwType', wintypes.DWORD),
        ('dwDisplayType', wintypes.DWORD),
        ('dwUsage', wintypes.DWORD),
        ('lpLocalName', wintypes.LPWSTR),
        ('lpRemoteName', wintypes.LPWSTR),
        ('lpComment', wintypes.LPWSTR),
        ('lpProvider', wintypes.LPWSTR),
    ]


def get_real_unc_path(path):
    """智能路径转换函数（跨会话版）"""
    try:
        # 如果已经是UNC路径直接返回
        if path.startswith('\\\\'):
            return os.path.normpath(path)

        # 获取所有活跃的网络连接
        net_resources = []
        dwResult = 0
        hEnum = ctypes.c_void_p()
        dwScope = 2  # RESOURCE_CONNECTED
        dwType = 0x3  # DISK
        dwUsage = 0
        dwFlags = 0

        # 枚举所有网络连接
        dwResult = ctypes.windll.mpr.WNetOpenEnumW(
            dwScope, dwType, dwUsage, None, ctypes.byref(hEnum))

        if dwResult != 0:
            return path

        buffer = (NETRESOURCE * 64)()
        entries = wintypes.DWORD(0xFFFFFFFF)
        bufferSize = wintypes.DWORD(ctypes.sizeof(buffer))

        dwResult = ctypes.windll.mpr.WNetEnumResourceW(
            hEnum, ctypes.byref(entries), buffer, ctypes.byref(bufferSize))

        if dwResult == 0:
            for i in range(entries.value):
                nr = buffer[i]
                if nr.lpLocalName:
                    net_resources.append((
                        nr.lpLocalName.upper(),
                        nr.lpRemoteName.upper()
                    ))

        ctypes.windll.mpr.WNetCloseEnum(hEnum)

        # 匹配最佳路径
        original_drive = os.path.splitdrive(path)[0].upper()
        for local, remote in net_resources:
            if local == original_drive:
                unc_base = remote
                relative_path = os.path.normpath(path[len(local):])
                return os.path.join(unc_base, relative_path).replace('\\', '/')

        # 尝试通过本地路径转换
        if not path.startswith('\\\\'):
            return '\\\\?\\' + os.path.abspath(path)

        return path
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 路径转换失败: {path}:{str(e)}")
        return path


def is_admin():
    """检查管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def elevate_with_params():
    """带参数提权（保持路径完整性）"""
    if is_admin():
        return True

    # 构造带参数的提权请求
    script = sys.argv[0]
    params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])

    try:
        ctypes.windll.shell32.ShellExecuteW(
            None,
            'runas',
            sys.executable,
            f'"{script}" {params}',
            None,
            1
        )
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 提权失败: {str(e)}")
    return False


def get_registry_value(key_path, value_name):
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
        value, _ = winreg.QueryValueEx(key, value_name)
        winreg.CloseKey(key)
        return value
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 注册表读取失败: {str(e)}")
        return None


def get_sursen_path():
    """获取Sursen安装路径（强化版）"""
    try:
        key_path = r"SursenReader.gw\\shell\\open\\command"
        value_name = ""
        registry_value = get_registry_value(key_path, value_name)
        if registry_value:
            if flag_dgb:
                print(f"[DEBUG] Found Sursen Installed : {registry_value}")

            i = 0
            lc = 0
            for c in registry_value:
                i += 1
                if c == "\\":
                    lc = i
            folder = registry_value[:lc]
            return folder
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 注册表读取失败: {str(e)}")
        else:
            sys.exit(1)


def handle_long_path(filepath):
    """处理长路径问题"""
    try:
        # 创建临时目录
        os.makedirs(TEMP_DIR, exist_ok=True)

        # 生成短文件名
        base_name = os.path.basename(filepath)
        name, ext = os.path.splitext(base_name)
        short_name = name[:32]
        candidate = f"{short_name}{ext}"

        i = 1
        while os.path.exists(os.path.join(TEMP_DIR, candidate)):
            candidate = f"{short_name}_{i}{ext}"
            i += 1
        temp_path = os.path.join(TEMP_DIR, candidate)

        # 复制文件
        shutil.copy2(filepath, temp_path)
        return temp_path
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 处理失败的路径: {filepath}:{str(e)}")
        else:
            sys.exit(1)


def main():
    global final_path
    """主逻辑函数"""
    if len(sys.argv) < 2:
        print("请通过文件关联或拖拽方式使用")
        sys.exit(1)

    # 原始路径处理
    raw_path = sys.argv[1]
    if flag_dgb:
        print(f"[DEBUG] 原始输入路径: {raw_path}")

    # 路径标准化
    normalized_path = os.path.abspath(raw_path)
    if flag_dgb:
        print(f"[DEBUG] 标准化路径: {normalized_path}")

    # 处理长路径/全部复制到临时文件夹

    final_path = handle_long_path(normalized_path)
    if flag_dgb:
        print(f"[DEBUG] 最终路径: {final_path}")


if __name__ == "__main__":
    flag_dgb = False
    # 文件名是否包含dbg
    if "dbg" in os.path.basename(sys.argv[0]):
        flag_dgb = True
        print("[DEBUG] 调试模式已启用")

    # 初始化环境
    if flag_dgb:
        print("Sursen Reader启动器 v2.0 (调试模式)")

    final_path = ""

    # 获取Sursen路径
    sursen_dir = get_sursen_path()
    cmd_path = os.path.join(sursen_dir, "SursenReader.exe")
    if flag_dgb:
        print(f"[DEBUG] 主程序路径: {cmd_path}")

    # 清理临时目录
    if os.path.exists(TEMP_DIR):
        for f in os.listdir(TEMP_DIR):
            try:
                os.remove(os.path.join(TEMP_DIR, f))
            except:
                pass
    else:
        os.makedirs(TEMP_DIR, exist_ok=True)

    # 执行主逻辑
    main()

    # 启动Sursen Reader
    try:
        subprocess.Popen([cmd_path, final_path], shell=True)
        if flag_dgb:
            print(f"[DEBUG] 启动命令: {cmd_path} {final_path}")
    except FileNotFoundError:
        if flag_dgb:
            print(f"[DEBUG] 找不到文件: {cmd_path}")
    except Exception as e:
        if flag_dgb:
            print(f"[DEBUG] 启动失败: {str(e)}")
    if flag_dgb:
        wait_for_any_key()
    sys.exit(0)
