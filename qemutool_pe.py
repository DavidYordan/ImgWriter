import os
import socket
import subprocess
import sys
import threading
import time
import yaml
from queue import Queue
from uuid import uuid4

class QemuTool:
    def __init__(self, device, queue, management_id, device_id):
        self.setup_paths(device, management_id, device_id)
        self.command_queue = Queue()
        self.core_port = None
        self.core_socket = None
        self.legacy_boot = False
        self.monitor_port = None
        self.monitor_socket = None
        self.queue = queue
        self.running = True
        self.tasks_queue = Queue()
        self.setup_tasks()
        self.current_state = self.tasks_queue.get()

    def connect_core(self):
        self.queue.put('尝试连接内核...')
        time.sleep(1)
        while True:
            try:
                if not self.core_socket:
                    self.core_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    self.core_socket.connect(('127.0.0.1', self.core_port))
                self.queue.put('内核连接成功。')
                return
            except:
                self.queue.put('连接内核失败, 正在重试...')
                time.sleep(1)

    def connect_monitor(self):
        self.queue.put('尝试连接监视器...')
        time.sleep(1)
        while True:
            try:
                if not self.monitor_socket:
                    self.monitor_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    self.monitor_socket.connect(('127.0.0.1', self.monitor_port))
                self.queue.put('监视器连接成功。')
                return
            except:
                self.queue.put('连接监视器失败, 正在重试...')
                time.sleep(1)

    def find_available_port(self, start_port=50000, end_port=60000):
        for port in range(start_port, end_port):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                try:
                    sock.bind(('127.0.0.1', port))
                    if not self.core_port:
                        self.core_port = port
                        continue
                    if not self.monitor_port:
                        self.monitor_port = port
                        return
                except socket.error:
                    continue
        raise RuntimeError("No available ports found in the specified range.")

    def prepare_optool_command(self):
        self.find_available_port()
        return [
            self.qemu,
            '-m', '512M',
            '-drive', f'file={self.optoolImg},format=raw,if=none,id=disk0',
            '-device', 'virtio-scsi-pci,id=scsi0',
            '-device', 'scsi-hd,drive=disk0,bus=scsi0.0',
            '-chardev', f'socket,id=char0,host=127.0.0.1,port={self.core_port},server,nowait',
            '-serial', 'chardev:char0',
            '-monitor', f'tcp:127.0.0.1:{self.monitor_port},server,nowait',
            '-nographic'
        ]

    def setup_tasks(self):
        self.tasks_queue.put(self.initial_state)
        self.tasks_queue.put(self.ready_state)
        self.tasks_queue.put(self.physicaldrive_check_state)
        self.tasks_queue.put(self.netflex_check_state)
        self.tasks_queue.put(self.format_disk_state)
        self.tasks_queue.put(self.write_img_state)
        self.tasks_queue.put(self.extend_disk_state)
        self.tasks_queue.put(self.mount_disk_state)
        self.tasks_queue.put(self.umount_disk_state)
        self.tasks_queue.put(self.end_state)
        self.tasks_queue.put(self.pass_state)

    def initial_state(self, line):
        if 'Please' in line:
            time.sleep(1)
            self.current_state = self.tasks_queue.get()
            self.queue.put('平台已就绪。')
            self.command_queue.put('')
            self.command_queue.put('')

    def ready_state(self, line):
        if '#' in line:
            time.sleep(0.5)
            self.current_state = self.tasks_queue.get()
            self.add_drives('physicaldrive')

    def physicaldrive_check_state(self, line):
        if 'Attached' not in line:
            return
        time.sleep(0.5)
        self.current_state = self.tasks_queue.get()
        self.command_queue.put('')
        time.sleep(0.5)
        self.add_drives('netflex')

    def netflex_check_state(self, line):
        if 'Attached' not in line:
            return
        time.sleep(0.5)
        self.current_state = self.tasks_queue.get()
        self.command_queue.put('')
        time.sleep(0.5)
        self.command_queue.put(f'parted /dev/sdb --script mklabel msdos')

    def format_disk_state(self, line):
        if 'msdos' not in line:
            return
        time.sleep(2)
        self.current_state = self.tasks_queue.get()
        self.queue.put(f'{self.device}刷入固件...')
        self.command_queue.put(f'dd if=/dev/sdc of=/dev/sdb bs=4M')

    def write_img_state(self, line):
        if 'out' in line:
            time.sleep(0.5)
            self.current_state = self.tasks_queue.get()
            self.queue.put(f'修复{self.device}...')
            self.command_queue.put(f'parted /dev/sdb')

    def extend_disk_state(self, line):
        if self.legacy_boot:
            if 'ext2' in line:
                time.sleep(0.5)
                self.command_queue.put('resizepart 2 100%')
                self.legacy_boot = False
            else:
                self.queue.put(f'硬盘格式异常，请寻求远程支持。')
                self.current_state = self.pass_state
                self.command_queue.put('quit')
        elif 'I/O' in line:
            time.sleep(0.5)
            self.command_queue.put('Retry')
        elif 'Welcome' in line:
            time.sleep(0.5)
            self.command_queue.put('print')
        elif 'corrupt' in line:
            time.sleep(0.5)
            self.command_queue.put('OK')
        elif 'current' in line:
            time.sleep(0.5)
            self.queue.put(f'修复{self.device}分区表...')
            self.command_queue.put('Fix')
        elif 'legacy_boot' in line:
            self.legacy_boot = True
        elif 'resizepart' in line:
            time.sleep(0.5)
            self.command_queue.put('quit')
        elif 'quit' in line:
            time.sleep(0.5)
            self.queue.put(f'检修{self.device}分区...')
            self.command_queue.put(f'e2fsck -f -p /dev/sdb2')
        elif 'inconsistency' in line:
            time.sleep(0.5)
            self.queue.put(f'硬盘格式异常，请尝试删除分区。')
            self.current_state = self.pass_state
        elif 'contiguous' in line:
            time.sleep(0.5)
            self.queue.put(f'扩容{self.device}空间...')
            self.command_queue.put(f'resize2fs /dev/sdb2')
        elif 'long' in line:
            time.sleep(1)
            self.current_state = self.tasks_queue.get()
            self.queue.put(f'挂载{self.device}...')
            self.command_queue.put(f'mkdir -p /mnt/disk && mount /dev/sdb2 /mnt/disk')

    def mount_disk_state(self, line):
        if 'argument' in line:
            self.current_state = self.pass_state
            self.queue.put(f'挂载失败，请重启软件重试...')
            return
        if 'mkdir' not in line:
            return
        time.sleep(1)
        self.current_state = self.tasks_queue.get()
        self.command_queue.put(f'echo -e "{self.yaml}" > /mnt/disk/etc/system.yaml')

    def umount_disk_state(self, line):
        if 'argument' in line:
            self.current_state = self.pass_state
            self.queue.put(f'挂载失败，请重启软件重试...')
            return
        if 'UUID' not in line:
            return
        time.sleep(0.5)
        self.current_state = self.tasks_queue.get()
        self.queue.put(f'卸载{self.device}...')
        self.command_queue.put('umount /mnt/disk')

    def end_state(self, line):
        if 'umount' not in line:
            return
        self.queue.put('固件刷入成功。')
        time.sleep(0.5)
        self.current_state = self.tasks_queue.get()
        self.queue.put('关闭固件平台...')
        self.command_queue.put('poweroff')
        self.running = False
        self.queue.put('FINISHED')

    def pass_state(self, line):
        pass

    def process_line(self, line):
        try:
            self.current_state(line)
        except Exception as e:
            self.queue.put(f'Processing Error: {e}')

    def read_core(self):
        buffer = b''
        while self.running:
            try:
                data = self.core_socket.recv(4096)
                if not data:
                    time.sleep(0.2)
                    continue
                buffer += data
                if not b'\n' in buffer:
                    time.sleep(0.2)
                    continue
                lines = buffer.split(b'\n')
                for line in lines[:-1]:
                    line = line.strip()
                    if line:
                        line_str = line.decode('utf-8')
                        print(line_str)
                        self.process_line(line_str)
                    buffer = lines[-1]
            except Exception as e:
                self.queue.put(f'读取内核出错: {e}')
                time.sleep(0.2)

    def run(self):
        self.write_img_to_disk()

    def run_qemu(self, command):
        return subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    def send_command(self):
        while self.running:
            command = self.command_queue.get()
            print(f'发送命令: {command}')
            try:
                self.core_socket.sendall(f'{command}\n'.encode())
            except Exception as e:
                self.queue.put(f'Sending Error: {e}')
            finally:
                if command == 'poweroff':
                    return

    def setup_paths(self, device, management_id, device_id):
        if hasattr(sys, '_MEIPASS'):
            sysPath = sys._MEIPASS
        else:
            sysPath = os.path.abspath('.')
        self.device = device
        self.qemu = os.path.join(sysPath, 'qemutools', 'qemu-system-x86_64.exe')
        self.netflexImg = os.path.join(sysPath, 'img', 'netflex.img')
        self.optoolImg = os.path.join(sysPath, 'img', 'optool.img')
        self.yaml = yaml.dump(
            {
                'UUID': str(uuid4()),
                'ManagementID': management_id,
                'DeviceID': device_id,
                'LocalPort': 56765,
                'Servers': [f'clent{i}.duoruduochu.com' for i in range(1, 7)],
                'HeartbeatInterval': 10,
                'HeartbeatRetries': 3
            },
            default_flow_style=False
        ).replace('\n', '\\n').replace('"', '\\"')

    def write_img_to_disk(self):
        process = self.run_qemu(self.prepare_optool_command())
        self.connect_core()
        self.connect_monitor()
        self.queue.put('加载固件平台...')

        read_thread = threading.Thread(target=self.read_core)
        read_thread.start()

        write_thread = threading.Thread(target=self.send_command)
        write_thread.start()

        try:
            read_thread.join()
            write_thread.join()
        except Exception as e:
            self.queue.put(f'Writing Error: {e}')
        finally:
            if self.core_socket:
                self.core_socket.close()
            if self.monitor_socket:
                self.monitor_socket.close()
            process.terminate()
            process.wait()

    def add_drives(self, drive_type):
        if drive_type == 'physicaldrive':
            self.queue.put('装载硬盘...')
            self.send_monitor_command(f'drive_add 0 file={self.device},if=none,id=disk1,format=raw')
            self.send_monitor_command('device_add scsi-hd,drive=disk1,bus=scsi0.0')
        elif drive_type == 'netflex':
            self.queue.put('装载固件...')
            self.send_monitor_command(f'drive_add 0 file={self.netflexImg},if=none,id=disk2,format=raw')
            self.send_monitor_command('device_add scsi-hd,drive=disk2,bus=scsi0.0')

    def send_monitor_command(self, command):
        print(f'发送监视器命令: {command}')
        try:
            self.monitor_socket.sendall(f'{command}\n'.encode())
        except Exception as e:
            self.queue.put(f'Monitor Error: {e}')
