import psutil
import subprocess
import time
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext
from queue import Queue, Empty
from threading import Thread
from qemutool_pe import QemuTool

class DiskImageWriter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.columns = [
            "index", "device", "model", "size", "type", "status"
        ]
        self.fields = [
            'disk_id', 'type', 'status', 'path', 'target', 'lun_id',
            'location_path', 'current_readonly_state', 'readonly', 'boot_disk',
            'pagefile_disk', 'hibernation_file_disk', 'crashdump_disk', 'clustered_disk'
        ]
        self.qemu_thread = None
        self.qemu_tool = None
        self.queue = Queue()
        self.init_ui()

    def run_diskpart_command(self, commands):
        process = subprocess.Popen(["diskpart"], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate(input=commands.encode())
        return stdout.decode('gbk', errors='ignore')

    def get_physical_disks(self):
        physical_disks = []
        try:
            output = self.run_diskpart_command("list disk\n")
            
            disk_info_pattern = re.compile(r"^(\*?)\s+(\w+)\s+(\d+)\s+(\w+)\s+(\d+\s+\w+)\s+(\d+\s+\w+)(?:\s+(\w*))?(?:\s+(\*?))?\r?$", re.MULTILINE)
            disks = disk_info_pattern.findall(output)
            
            for disk in disks:
                current, drive, index, status, size, free, dyn, gpt = disk
                disk_info = {
                    'has_partitions': True,
                    'current': True if current else False,
                    'device': f'\\\\.\\PHYSICALDRIVE{index}',
                    'index': index,
                    'status': status,
                    'size': size,
                    'free': free,
                    'dyn': dyn,
                    'gpt': True if gpt else False
                }

                detail_output = self.run_diskpart_command(f"select disk {index}\ndetail disk\n")
                if '###' not in detail_output:
                    disk_info['has_partitions'] = False
                detail_lines = detail_output.splitlines()
                start_index = 0
                count = 0
                for i, line in enumerate(detail_lines):
                    if line == 'DISKPART> ':
                        count += 1
                    if count == 2:
                        disk_info['model'] = detail_lines[i + 1].strip()
                        start_index = i + 2
                        break

                detail_lines = detail_lines[start_index:]
                for i, line in enumerate(detail_lines[:len(self.fields)]):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        disk_info[self.fields[i]] = parts[1].strip()
                output_log = f'扫描到{drive}{index}: '
                for key, value in disk_info.items():
                    output_log += f'{key}: {value}, '
                self.log(output_log[:-2])

                physical_disks.append(disk_info)
        except Exception as e:
            self.log(f"获取磁盘信息失败: {e}")

        if not physical_disks:
            self.log("请删除硬盘分区。")
        return physical_disks

    def init_ui(self):
        self.title("DiskImageWriter")
        self.geometry("840x400")

        self.disk_table = ttk.Treeview(self, columns=self.columns, show="headings")
        self.disk_table.heading("index", text="索引")
        self.disk_table.heading("device", text="设备")
        self.disk_table.heading("model", text="型号")
        self.disk_table.heading("size", text="大小")
        self.disk_table.heading("type", text="类型")
        self.disk_table.heading("status", text="状态")
        self.disk_table.pack(fill=tk.BOTH, expand=True)

        for col in self.columns:
            self.disk_table.column(col, minwidth=0, width=100, stretch=tk.YES)

        button_frame = tk.Frame(self)
        button_frame.pack(fill=tk.X)

        self.refresh_button = tk.Button(button_frame, text="重新扫描", command=self.refresh_disk_list)
        self.refresh_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.start_button = tk.Button(button_frame, text="刷入固件", command=self.start_write_and_extend)
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.refresh_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)

        self.log_output = scrolledtext.ScrolledText(self, state='disabled')
        self.log_output.pack(fill=tk.BOTH, expand=True)

        command_frame = tk.Frame(self)
        command_frame.pack(fill=tk.X)

        self.command_line = tk.Entry(command_frame)
        self.command_line.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        self.command_line.bind("<Return>", lambda event: self.send_command())

        self.send_button = tk.Button(command_frame, text="发送", command=self.send_command)
        self.send_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.refresh_disk_list()
        self.process_queue()

    def process_queue(self):
        try:
            while True:
                message = self.queue.get_nowait()
                self.log(message)
                if message == 'FINISHED':
                    self.on_write_finished()
        except Empty:
            time.sleep(0.1)
            pass
        self.after(100, self.process_queue)

    def log(self, message):
        self.log_output.config(state='normal')
        self.log_output.insert(tk.END, message + "\n")
        self.log_output.config(state='disabled')
        self.log_output.yview(tk.END)

    def on_write_finished(self):
        self.start_button.config(state=tk.NORMAL)

    def refresh_disk_list(self):
        disks = self.get_physical_disks()
        for item in self.disk_table.get_children():
            self.disk_table.delete(item)
        for disk in disks:
            if disk['current'] or disk['has_partitions']:
                continue
            self.disk_table.insert("", tk.END, values=(
                disk['index'],
                disk['device'],
                disk['model'],
                disk['size'],
                disk['type'],
                disk['status']
            ))

    def send_command(self):
        if not self.qemu_tool:
            return

        command = self.command_line.get()
        self.qemu_tool.command_queue.put(command)
        self.command_line.delete(0, tk.END)

    def start_write_and_extend(self):
        selected_item = self.disk_table.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请选择一个磁盘。")
            return

        device = self.disk_table.item(selected_item[0], 'values')[1]

        management_id = simpledialog.askstring("请输入管理ID", f"您选择了设备{device}\n请输入管理ID:")
        if not management_id:
            messagebox.showwarning("警告", "管理ID不能为空。")
            return

        device_id = simpledialog.askstring("请输入设备标识", f"您选择了设备: {device}\n您的管理ID是: {management_id}\n请输入设备标识:")
        if not device_id:
            messagebox.showwarning("警告", "设备标识不能为空。")
            return

        confirm = messagebox.askquestion("确认", f"您选择了设备: {device}\n您的管理ID是: {management_id}\n您的设备标识是: {device_id}\n您确定要刷入固件吗？这将擦除磁盘上的所有数据。")
        if confirm != 'yes':
            return
        
        self.start_button.config(state=tk.DISABLED)

        if not self.terminate_qemu_process():
            self.start_button.config(state=tk.NORMAL)
            return
        
        self.qemu_tool = QemuTool(device, self.queue, management_id, device_id)
        self.qemu_thread = Thread(target=self.qemu_tool.run)
        self.qemu_thread.start()

    def terminate_qemu_process(self):
        for proc in psutil.process_iter(['pid', 'name']):
            if 'qemu-system-x86_64.exe' != proc.name():
                continue
            try:
                p = psutil.Process(proc.info['pid'])
                p.terminate()
                p.wait(timeout=3)
                return True
            except psutil.AccessDenied:
                self.log('无法终止管道，权限不足。')
                return False
            except psutil.NoSuchProcess:
                return True
            except psutil.TimeoutExpired:
                self.log('终止管道超时。')
                return False
        return True

if __name__ == "__main__":
    app = DiskImageWriter()
    app.mainloop()
