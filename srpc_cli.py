#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sursen Reader打印色彩修复工具 CLI版本
作者：Zhai @ 20250124
版本：3.1
# pyinstaller --uac-admin -F srpc_cli.py --icon=sursen.ico
"""

import winreg
import os
import ctypes
import sys
import argparse
import logging
from pathlib import Path
from typing import Optional, Tuple
import colorama

# 初始化colorama
colorama.init()

# ======================
# 颜色配置（改用colorama）
# ======================
COLOR = {
    "RED": colorama.Fore.RED,
    "GREEN": colorama.Fore.GREEN,
    "YELLOW": colorama.Fore.YELLOW,
    "BLUE": colorama.Fore.BLUE,
    "RESET": colorama.Style.RESET_ALL
}

# ======================
# 命令行参数解析
# ======================
def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="Sursen Reader打印色彩修复工具",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('-s', '--silent', 
                        action='store_true',
                        help='静默模式运行（不显示控制台输出）')
    parser.add_argument('-l', '--log', 
                        type=Path,
                        default='sursen_fix.log',
                        help='指定日志文件路径（默认：sursen_fix.log）')
    return parser.parse_args()

# ======================
# 权限管理模块
# ======================
class PrivilegeManager:
    """管理员权限管理"""
    
    @staticmethod
    def is_admin() -> bool:
        """检查管理员权限"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    @staticmethod
    def elevate() -> bool:
        """提升权限"""
        try:
            params = ' '.join([f'"{x}"' if ' ' in x else x for x in sys.argv])
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, params, None, 1
            )
            return True
        except Exception as e:
            logging.error(f"权限提升失败: {e}")
            return False

# ======================
# 注册表操作模块
# ======================
class RegistryOperator:
    """注册表操作类"""
    
    @staticmethod
    def get_install_path() -> Optional[str]:
        """获取Sursen安装路径"""
        key_path = r"SursenReader.gw\shell\open\command"
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path) as key:
                value, _ = winreg.QueryValueEx(key, "")
                return value.strip('"').replace('\\"', '').replace('\\\\', '\\')
        except FileNotFoundError:
            logging.error("未找到Sursen注册表项")
            return None
        except Exception as e:
            logging.error(f"注册表访问错误: {e}")
            return None

    @staticmethod
    def set_runas(exe_path: str) -> Tuple[bool, str]:
        """设置管理员运行权限"""
        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path,
                              0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                try:
                    current_value = winreg.QueryValueEx(key, exe_path)[0]
                except FileNotFoundError:
                    new_value = "~ RUNASADMIN"
                else:
                    new_value = f"{current_value} RUNASADMIN" if "RUNASADMIN" not in current_value else current_value

                winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, new_value)
                return True, "管理员权限设置成功"
        except Exception as e:
            return False, f"注册表写入失败: {e}"

# ======================
# 配置文件修复模块
# ======================
class ConfigFixer:
    """配置文件修复器"""
    
    @staticmethod
    def modify_config(config_path: Path) -> Tuple[bool, str]:
        """修改单个配置文件"""
        if not config_path.exists():
            return False, "配置文件不存在"
        
        try:
            modified = False
            with open(config_path, 'r+', encoding='utf-8') as f:
                lines = f.readlines()
                for i, line in enumerate(lines):
                    if line.startswith('PrintBlackAndWhite='):
                        if line.strip() != 'PrintBlackAndWhite=0':
                            lines[i] = 'PrintBlackAndWhite=0\n'
                            modified = True
                            break
                
                if modified:
                    f.seek(0)
                    f.writelines(lines)
                    f.truncate()
                    return True, "配置文件修改成功"
            return False, "配置文件无需修改"
        except PermissionError:
            return False, "权限不足"
        except Exception as e:
            return False, f"文件操作失败: {e}"

    @classmethod
    def fix_all_configs(cls, install_path: Optional[str]) -> None:
        """修复所有配置文件"""
        # 系统配置文件
        if install_path:
            sys_config = Path(install_path).parent / "SursenPrint.ini"
            status, msg = cls.modify_config(sys_config)
            logging.info(f"系统配置: {msg}")

        # 用户配置文件
        users_dir = Path("C:/Users")
        for user_dir in users_dir.iterdir():
            if user_dir.is_dir() and user_dir.name not in ['Public', 'Default', 'All Users']:
                user_config = user_dir / 'AppData'/'Local'/'VirtualStore'/'Program Files (x86)'/'Sursen'/'Reader'/'SursenPrint.ini'
                if user_config.exists():
                    status, msg = cls.modify_config(user_config)
                    logging.info(f"用户 {user_dir.name}: {msg}")

# ======================
# CLI界面模块
# ======================
class CLIInterface:
    """命令行界面控制器"""
    
    def __init__(self, silent: bool = False):
        self.silent = silent

    def print_colored(self, message: str, color: str) -> None:
        """跨平台彩色输出"""
        if not self.silent:
            color_code = COLOR.get(color.upper(), COLOR["RESET"])
            print(f"{color_code}{message}{COLOR['RESET']}")

    def show_header(self) -> None:
        """显示标题"""
        if self.silent:
            return
        self.print_colored("=== Sursen Reader打印色彩修复工具 ===", "BLUE")
        self.print_colored("作者：Zhai @ 20250124 | 版本：3.2\n", "YELLOW")
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
# ======================
# 主程序逻辑
# ======================
def main():
    # 初始化
    args = parse_args()
    cli = CLIInterface(args.silent)
    
    # 配置日志（兼容Python 3.8）
    log_format = '[%(asctime)s] %(levelname)s: %(message)s'
    if sys.version_info >= (3, 9):
        logging.basicConfig(
            filename=args.log,
            level=logging.INFO,
            format=log_format,
            encoding='utf-8'
        )
    else:
        logging.basicConfig(
            filename=args.log,
            level=logging.INFO,
            format=log_format
        )
        # 手动处理文件编码
        if Path(args.log).exists():
            for handler in logging.root.handlers[:]:
                if isinstance(handler, logging.FileHandler):
                    handler.close()
                    logging.root.removeHandler(handler)
            file_handler = logging.FileHandler(args.log, mode='a', encoding='utf-8')
            file_handler.setFormatter(logging.Formatter(log_format))
            logging.root.addHandler(file_handler)

    # 显示标题
    cli.show_header()

    # 权限检查
    if not PrivilegeManager.is_admin():
        cli.print_colored("错误: 需要管理员权限运行！", "RED")
        if not args.silent and input("尝试获取权限？(y/n): ").lower() == 'y':
            if PrivilegeManager.elevate():
                sys.exit(0)
        sys.exit(1)

    try:
        # 获取安装路径
        install_path = RegistryOperator.get_install_path()
        if install_path:
            cli.print_colored(f"检测到安装路径: {install_path}", "GREEN")
            set_runas_admin = False
            # 设置管理员权限
            if set_runas_admin:
                # 设置管理员权限
                success, msg = RegistryOperator.set_runas(install_path)
                cli.print_colored(msg, "GREEN" if success else "RED")
            else:
                cli.print_colored("不设置管理员权限", "YELLOW")
                
            # 修复配置文件
            ConfigFixer.fix_all_configs(install_path)
            cli.print_colored("\n设置彩色打印操作完成！请重启Sursen Reader生效", "GREEN")
        else:
            cli.print_colored("未检测到Sursen安装", "YELLOW")


        wait_for_any_key()
        # 按任意键继续

    except KeyboardInterrupt:
        cli.print_colored("\n操作已取消", "RED")
        sys.exit(130)
    except Exception as e:
        cli.print_colored(f"\n发生致命错误: {str(e)}", "RED")
        logging.exception("程序异常终止")
        sys.exit(2)

if __name__ == "__main__":
    main()