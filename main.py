import psutil
import subprocess
import sys
import win32com.client
from PyQt6 import QtWidgets
from PyQt6.QtGui import QTextCursor
from PyQt6.QtCore import QThread

from qemutool import QemuTool

class DiskImageWriter(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.qemu_thread = None
        self.qemu_tool = None
        self.init_ui()

    def get_physical_disks(self):
        physical_disks = []
        c = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        connection = c.ConnectServer(".", r"root\cimv2")
        for disk in connection.ExecQuery("Select * from Win32_DiskDrive"):
            if self.has_no_partitions(disk.Index):
                physical_disks.append({
                    'device': disk.DeviceID,
                    'index': disk.Index,
                    'manufacturer': disk.Manufacturer,
                    'model': disk.Model,
                    'size': f'{int(int(disk.Size if disk.Size else 0) // (1024 ** 3))}GB',
                    'serial_number': disk.SerialNumber.strip()
                })
        if physical_disks:
            self.log(f"扫描到硬盘: {physical_disks}")
        else:
            self.log("请删除硬盘分区。")
        return physical_disks
    
    def init_ui(self):
        self.setWindowTitle("DiskImageWriter")
        self.setGeometry(100, 100, 840, 400)

        self.layout = QtWidgets.QVBoxLayout(self)

        self.disk_table = QtWidgets.QTableWidget(self)
        self.disk_table.setColumnCount(6)
        self.disk_table.setHorizontalHeaderLabels(["设备", "制造商", "型号", "大小", "序列号", "索引"])
        self.disk_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.disk_table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.disk_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.layout.addWidget(self.disk_table)

        hlayout = QtWidgets.QHBoxLayout()
        self.refresh_button = QtWidgets.QPushButton("重新扫描", self)
        self.refresh_button.clicked.connect(self.refresh_disk_list)
        hlayout.addWidget(self.refresh_button)

        self.start_button = QtWidgets.QPushButton("刷入固件", self)
        self.start_button.clicked.connect(self.start_write_and_extend)
        hlayout.addWidget(self.start_button)

        self.layout.addLayout(hlayout)

        self.log_output = QtWidgets.QTextEdit(self)
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output)

        self.command_layout = QtWidgets.QHBoxLayout()

        self.command_line = QtWidgets.QLineEdit(self)
        self.command_line.returnPressed.connect(self.send_command)
        self.command_layout.addWidget(self.command_line)

        self.send_button = QtWidgets.QPushButton("发送", self)
        self.send_button.clicked.connect(self.send_command)
        self.command_layout.addWidget(self.send_button)

        self.layout.addLayout(self.command_layout)

        self.refresh_disk_list()

    def log(self, message):
        self.log_output.append(message)
        self.log_output.textChanged.connect(lambda: self.log_output.moveCursor(QTextCursor.MoveOperation.End))

    def has_no_partitions(self, disk_index):
        script = f"select disk {disk_index} \n detail disk"
        process = subprocess.Popen("diskpart", stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate(input=script.encode())

        output = stdout.decode('gbk', errors='ignore')
        return '###' not in output

    def on_write_finished(self):
        self.start_button.setEnabled(True)

    def refresh_disk_list(self):
        disks = self.get_physical_disks()
        self.disk_table.setRowCount(0)
        for disk in disks:
            row_position = self.disk_table.rowCount()
            self.disk_table.insertRow(row_position)
            self.disk_table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(disk['device']))
            self.disk_table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(disk['manufacturer']))
            self.disk_table.setItem(row_position, 2, QtWidgets.QTableWidgetItem(disk['model']))
            self.disk_table.setItem(row_position, 3, QtWidgets.QTableWidgetItem(disk['size']))
            self.disk_table.setItem(row_position, 4, QtWidgets.QTableWidgetItem(disk['serial_number']))
            self.disk_table.setItem(row_position, 5, QtWidgets.QTableWidgetItem(str(disk['index'])))

    def send_command(self):
        if not self.qemu_tool:
            return

        command = self.command_line.text()
        self.qemu_tool.command_queue.put(command)
        self.command_line.clear()

    def start_write_and_extend(self):
        selected_row = self.disk_table.currentRow()
        if selected_row == -1:
            QtWidgets.QMessageBox.warning(self, "警告", "请选择一个磁盘。")
            return

        device = self.disk_table.item(selected_row, 0).text()

        confirm = QtWidgets.QMessageBox.question(self, "确认", f"你确定要刷入固件至 {device} 吗？这将擦除磁盘上的所有数据。")
        if confirm != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        
        self.start_button.setEnabled(False)

        if not self.terminate_qemu_process():
            self.start_button.setEnabled(True)
            return
        
        
        self.qemu_tool = QemuTool(device)
        self.qemu_thread = QThread()
        self.qemu_tool.moveToThread(self.qemu_thread)
        self.qemu_tool.output_signal.connect(self.log)
        self.qemu_tool.finished_signal.connect(self.on_write_finished)
        self.qemu_thread.started.connect(self.qemu_tool.run)
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
    app = QtWidgets.QApplication(sys.argv)
    ex = DiskImageWriter()
    ex.show()
    sys.exit(app.exec())
