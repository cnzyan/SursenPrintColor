import winreg
import os
import ctypes
import sys
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import threading

# .\.venv\Scripts\activate.ps1
# pyinstaller -F -w --icon=your_icon.ico srpc_gui.py
# pyinstaller -F --uac-admin -w srpc_gui.py
class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        
    def write(self, message):
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        
    def flush(self):
        pass

class SursenFixerApp:
    def __init__(self, master):
        self.master = master
        master.title("Sursen Reader修复工具 v1.0")
        master.geometry("800x600")

        # 权限检查
        if not self.is_admin():
            self.show_admin_warning()
            return

        # 界面组件
        self.create_widgets()
        sys.stdout = StdoutRedirector(self.output_area)

    def create_widgets(self):
        # 顶部说明
        self.header = ttk.Label(self.master, 
                               text="Sursen Reader打印色彩修复工具\n作者：Zhai @20250124",
                               font=('Arial', 12),
                               justify='center')
        self.header.pack(pady=10)

        # 输出区域
        self.output_area = scrolledtext.ScrolledText(self.master, wrap=tk.WORD)
        self.output_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 进度条
        self.progress = ttk.Progressbar(self.master, orient=tk.HORIZONTAL, mode='indeterminate')
        
        # 按钮框架
        self.btn_frame = ttk.Frame(self.master)
        self.btn_frame.pack(pady=10)

        self.start_btn = ttk.Button(self.btn_frame, 
                                   text="开始修复", 
                                   command=self.start_repair)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.exit_btn = ttk.Button(self.btn_frame,
                                  text="退出",
                                  command=self.master.destroy)
        self.exit_btn.pack(side=tk.LEFT, padx=5)

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def show_admin_warning(self):
        messagebox.showerror("权限错误", "请以管理员身份运行本程序！")
        self.master.after(1000, self.master.destroy)

    def start_repair(self):
        self.start_btn.config(state=tk.DISABLED)
        self.progress.pack(pady=5)
        self.progress.start()
        
        repair_thread = threading.Thread(target=self.run_repair, daemon=True)
        repair_thread.start()

    def run_repair(self):
        try:
            print("="*50)
            print("开始修复流程...")
            
            # 原main逻辑
            key_path = r"SursenReader.gw\\shell\\open\\command"
            value_name = ""
            registry_value = self.get_registry_value(key_path, value_name)
            
            if registry_value:
                print(f"发现Sursen阅读器安装路径: {registry_value}")
                exe_path = self.process_registry_value(registry_value)
                self.add_runas(exe_path)
                self.fix_config_files(exe_path)
            
            self.fix_user_configs()
            
            print("\n彩色打印修复完成！")
            messagebox.showinfo("完成", "所有修复操作已完成！")
            
        except Exception as e:
            print(f"\n发生错误: {str(e)}")
            messagebox.showerror("错误", f"修复过程中发生错误：{str(e)}")
        finally:
            self.progress.stop()
            self.progress.pack_forget()
            self.start_btn.config(state=tk.NORMAL)

    # 以下是原功能函数（稍作修改）
    def get_registry_value(self, key_path, value_name):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
            value, _ = winreg.QueryValueEx(key, value_name)
            winreg.CloseKey(key)
            return value
        except Exception as e:
            print(f"注册表访问错误: {e}")
            return None

    def process_registry_value(self, value):
        last_backslash = value.rfind('\\') + 1
        folder = value[:last_backslash]
        return os.path.join(folder, "SursenReader.exe")

    def add_runas(self, exe_path):
        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 
                              access=winreg.KEY_SET_VALUE | winreg.KEY_READ) as key:
                try:
                    current_value = winreg.QueryValueEx(key, exe_path)[0]
                except FileNotFoundError:
                    new_value = "~ RUNASADMIN"
                else:
                    new_value = f"{current_value} RUNASADMIN" if "RUNASADMIN" not in current_value else current_value
                
                winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, new_value)
                print(f"成功添加管理员权限到：{exe_path}")
        except Exception as e:
            print(f"注册表修改失败: {e}")
            raise

    def fix_config_files(self, folder_path):
        config_path = os.path.join(os.path.dirname(folder_path), "SursenPrint.ini")
        try:
            with open(config_path, 'r+') as f:
                content = []
                modified = False
                for line in f:
                    if line.startswith('PrintBlackAndWhite='):
                        if '0' not in line:
                            line = 'PrintBlackAndWhite=0\n'
                            modified = True
                    content.append(line)
                
                if modified:
                    f.seek(0)
                    f.writelines(content)
                    f.truncate()
                    print(f"成功修复配置文件：{config_path}")
                else:
                    print(f"配置文件无需修改：{config_path}")
        except Exception as e:
            print(f"配置文件修改失败: {e}")
            raise

    def fix_user_configs(self):
        users_dir = "C:\\Users\\"
        print("\n正在扫描用户配置文件...")
        for username in os.listdir(users_dir):
            config_path = os.path.join(
                users_dir, username,
                'AppData\\Local\\VirtualStore\\Program Files (x86)\\Sursen\\Reader\\SursenPrint.ini'
            )
            if os.path.exists(config_path):
                try:
                    with open(config_path, 'r+') as f:
                        content = []
                        modified = False
                        for line in f:
                            if line.startswith('PrintBlackAndWhite='):
                                if '0' not in line:
                                    line = 'PrintBlackAndWhite=0\n'
                                    modified = True
                            content.append(line)
                        
                        if modified:
                            f.seek(0)
                            f.writelines(content)
                            f.truncate()
                            print(f"成功修复用户 {username} 的配置文件")
                except Exception as e:
                    print(f"用户 {username} 配置文件修改失败: {e}")

def main():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        return

    root = tk.Tk()
    app = SursenFixerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()