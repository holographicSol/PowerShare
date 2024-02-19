""" Written by Benjamin Jack Cullen """

import os.path
import sys
import subprocess
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5.QtCore import QTimer
import win32api

info = subprocess.STARTUPINFO()
info.dwFlags = 1
info.wShowWindow = 0
tray_icon = []

_program_name = 'PowerShare'

def setExecutionPolicyUnrestricted():
    cmd = 'powershell Set-ExecutionPolicy Unrestricted'
    xcmd = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    while True:
        output = xcmd.stdout.readline()
        if output == '' and xcmd.poll() is not None:
            break
        if output:
            output = output.decode("utf-8").strip()
            print(output)
        else:
            break


class SystemTrayIcon(QtWidgets.QSystemTrayIcon):

    def __init__(self, icon, parent=None):

        QtWidgets.QSystemTrayIcon.__init__(self, icon, parent)

        self.context_menu_style = """QMenu 
                                   {background-color: rgb(33, 33, 33);
                                   color: rgb(255, 255, 255);
                                   border-top:2px solid rgb(0, 0, 0);
                                   border-bottom:2px solid rgb(0, 0, 0);
                                   border-right:2px solid rgb(0, 0, 0);
                                   border-left:2px solid rgb(0, 0, 0);
                                   }
                                   QMenu::item::selected
                                   {background-color : rgb(33, 33, 33);
                                   color: rgb(33, 33, 255);
                                   }
                                   """

        # initiate QMenu and set context menu style
        menu = QtWidgets.QMenu(parent)
        menu.setStyleSheet(self.context_menu_style)

        # initiate context menu items
        menu.addSeparator()
        menu.addAction(QtGui.QIcon("./icon.ico"), _program_name)
        menu.addSeparator()
        on_action = menu.addAction(QtGui.QIcon("./img_user.ico"), "Enable Share All (Network share every disk)")
        menu.addSeparator()
        don_action = menu.addAction(QtGui.QIcon("./img_admin.ico"), "Enable Default Shares")
        menu.addSeparator()
        off_action = menu.addAction(QtGui.QIcon("./img_admin.ico"), "Disable Share All (Disable network share every disk)")
        menu.addSeparator()
        isolate_action = menu.addAction(QtGui.QIcon("./img_isolate.ico"),
                                        "Isolate (Disable Share All + Disable Admin Shares)")
        menu.addSeparator()
        exit_action = menu.addAction(QtGui.QIcon("./img_exit.ico"), "Exit")

        # set context menuQtGui.QIcon
        self.setContextMenu(menu)

        # plug context menu items into functions
        exit_action.triggered.connect(self.exit)
        on_action.triggered.connect(self.net_share_express_0)
        off_action.triggered.connect(self.net_share_express_1)
        isolate_action.triggered.connect(self.net_share_express_2)
        don_action.triggered.connect(self.net_share_express_3)

        # monitor interval
        self.timer_0 = QTimer(self)
        self.timer_0.setInterval(500)
        self.timer_0.timeout.connect(self.timer_0_function)
        self.timer_0_start_function()

    def exit(self):
        QtCore.QCoreApplication.exit()

    def net_share_express_0(self):
        self.net_share_express(_command=int(0))

    def net_share_express_1(self):
        self.net_share_express(_command=int(1))

    def net_share_express_2(self):
        self.net_share_express(_command=int(2))

    def net_share_express_3(self):
        self.net_share_express(_command=int(3))

    @QtCore.pyqtSlot()
    def net_share_express(self, _command: int):

        # Enable user shares for every disk
        drives = win32api.GetLogicalDriveStrings()
        drives = drives.split('\000')[:-1]
        for drive in drives:
            if _command == int(0):
                cmd = 'net share ' + str(drive.strip(':\\').lower() + '=' + str(drive) + ' /unlimited /grant:everyone,FULL')
                subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

        # Disable shares
        if _command == int(1) or _command == int(2):
            output_idx = 0
            xcmd = subprocess.Popen('net share', shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            while True:
                output = xcmd.stdout.readline()
                if output == '' and xcmd.poll() is not None:
                    break
                if output:
                    output = output.decode("utf-8").strip()
                    output_split = output.split()
                    if output == 'There are no entries in the list.':
                        break

                    # Apply output filter
                    if output_idx >= 4 and len(output_split) >= 1 and output_split[0] != 'The':
                        cmd = f'net share {output_split[0]} /delete'

                        # Disable user shares only
                        if _command == int(1):
                            if not output_split[0].endswith('$') and _command == 1:
                                subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

                        # Disable admin shares and user shares
                        elif _command == int(2):
                            subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                else:
                    break

                output_idx += 1

        elif _command == int(3):
            xcmd = subprocess.Popen('powershell Set-ItemProperty -Name AutoShareWks -Path HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters -Value 1', shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            while True:
                output = xcmd.stdout.readline()
                if output == '' and xcmd.poll() is not None:
                    break
                if output:
                    output = output.decode("utf-8").strip()
                    print(output)
                else:
                    break
            xcmd = subprocess.Popen('powershell ./default_admin_shares.ps1', shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            while True:
                output = xcmd.stdout.readline()
                if output == '' and xcmd.poll() is not None:
                    break
                if output:
                    output = output.decode("utf-8").strip()
                    print(output)
                else:
                    break

    @QtCore.pyqtSlot()
    def timer_0_start_function(self):
        self.timer_0.start()

    @QtCore.pyqtSlot()
    def timer_0_stop_function(self):
        self.timer_0.stop()

    @QtCore.pyqtSlot()
    def timer_0_function(self):
        """ This function monitors for active network shares """

        admin_shares = False
        user_shares = False
        output_idx = 0
        xcmd = subprocess.Popen('net share', shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

        while True:
            output = xcmd.stdout.readline()
            if output == '' and xcmd.poll() is not None:
                break
            if output:
                output = output.decode("utf-8").strip()
                output_split = output.split()

                # Apply output filter
                if output_idx >= 4 and len(output_split) >= 1 and output_split[0] != 'The':

                    # No shares found
                    if 'There are no entries in the list.' in output:
                        admin_shares = False
                        user_shares = False
                        break

                    # Admin shares found
                    elif output_split[0].endswith('$'):
                        admin_shares = True

                    # User shares found
                    elif os.path.exists(output_split[1]):
                        user_shares = True
            else:
                break
            output_idx += 1
        xcmd.poll()

        # Disk Letter Shares Enabled
        if user_shares is True and admin_shares is True:
            tray_icon.setIcon(QtGui.QIcon('./img_adminuser.ico'))

        # Admin Shares Enabled
        elif admin_shares is True and user_shares is False:
            tray_icon.setIcon(QtGui.QIcon('./img_admin.ico'))

        # User Shares Enabled
        elif admin_shares is False and user_shares is True:
            tray_icon.setIcon(QtGui.QIcon('./img_user.ico'))

        # Shares Disabled
        elif admin_shares is False and user_shares is False:
            tray_icon.setIcon(QtGui.QIcon('./img_isolate.ico'))


def main(image):
    global tray_icon

    setExecutionPolicyUnrestricted()

    app = QtWidgets.QApplication(sys.argv)
    w = QtWidgets.QWidget()
    tray_icon = SystemTrayIcon(QtGui.QIcon(image), w)
    tray_icon.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    on = 'icon.ico'
    main(on)
