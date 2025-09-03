# canDrive @ 2020
# 生成单文件可执行程序：pyinstaller -F main.spec
#----------------------------------------------------------------
import serial
import canSniffer_ui
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem, QHeaderView, QFileDialog, QRadioButton
from PyQt5.QtWidgets import QVBoxLayout, QSizeGrip
from PyQt5.QtCore import Qt,QEvent 
from PyQt5.QtGui import QColor
import serial.tools.list_ports

import sys
import os
import time
import qtmodern
from qtmodern import styles
from qtmodern import windows
import csv

import HideOldPackets
import SerialReader
import SerialWriter
import FileLoader

# 启用高DPI缩放和高DPI图标
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

class canSnifferGUI(QMainWindow, canSniffer_ui.Ui_MainWindow):
    def __init__(self):
        super(canSnifferGUI, self).__init__()
        self.setupUi(self)
        
        self.setWindowFlags(Qt.Window | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("CAN Sniffer")
        self.setGeometry(200, 200, 600, 400)

        # 连接各个按钮和控件的信号与槽
        self.portScanButton.clicked.connect(self.scanPorts)
        self.portConnectButton.clicked.connect(self.serialPortConnect)
        self.portDisconnectButton.clicked.connect(self.serialPortDisconnect)
        self.startSniffingButton.clicked.connect(self.startSniffing)
        self.stopSniffingButton.clicked.connect(self.stopSniffing)
        self.saveSelectedIdInDictButton_can2.clicked.connect(self.saveIdLabelToDictCallback)
        self.saveSessionToFileButton_can2.clicked.connect(self.saveSessionToFile)
        self.loadSessionFromFileButton.clicked.connect(self.loadSessionFromFile)
        self.showOnlyIdsLineEdit_can2.textChanged.connect(self.can2showOnlyIdsTextChanged)
        self.hideIdsLineEdit_can2.textChanged.connect(self.can2hideIdsTextChanged)
        self.clearLabelDictButton.clicked.connect(self.clearLabelDict)
        self.serialController = serial.Serial()
        self.can2MessageTableWidget.cellClicked.connect(self.cellWasClicked)
        self.newTxTableRow.clicked.connect(self.newTxTableRowCallback)
        self.removeTxTableRow.clicked.connect(self.removeTxTableRowCallback)
        self.sendTxTableButton.clicked.connect(self.sendTxTableCallback)
        self.abortSessionLoadingButton.clicked.connect(self.abortSessionLoadingCallback)
        self.showSendingTableCheckBox.clicked.connect(self.showSendingTableButtonCallback)
        self.addToDecodedPushButton_can2.clicked.connect(self.addToDecodedCallback)
        self.deleteDecodedPacketLinePushButton.clicked.connect(self.deleteDecodedLineCallback)
        self.decodedMessagesTableWidget.itemChanged.connect(self.decodedTableItemChangedCallback)
        self.clearTableButton_can2.clicked.connect(self.clearTableCallback)
        self.sendSelectedDecodedPacketButton.clicked.connect(self.sendDecodedPacketCallback)
        self.playbackMainTableButton_can2.clicked.connect(self.playbackMainTableCallback)
        self.stopPlayBackButton_can2.clicked.connect(self.stopPlayBackCallback)
        self.hideAllPacketsButton_can2.clicked.connect(self.hideAllPackets)
        self.showControlsButton.hide()
        
        # 初始化串口读写和文件加载线程
        self.serialWriterThread = SerialWriter.SerialWriterThread(self.serialController)
        self.serialReaderThread = SerialReader.SerialReaderThread(self.serialController)
        self.serialReaderThread.receivedPacketSignal.connect(self.serialPacketReceiverCallback)
        self.fileLoaderThread = FileLoader.FileLoaderThread()
        self.fileLoaderThread.newRowSignal.connect(self.fileLoaderCallback)
        self.fileLoaderThread.loadingFinishedSignal.connect(self.fileLoadingFinishedCallback)
        self.hideOldPacketsThread = HideOldPackets.HideOldPacketsThread()
        self.hideOldPacketsThread.hideOldPacketsSignal.connect(self.hideOldPacketsCallback)

        # 控件初始状态设置
        self.stopPlayBackButton_can2.setVisible(False)
        self.playBackProgressBar_can2.setVisible(False)
        self.sendingGroupBox.hide()
        self.hideOldPacketsThread.enable(5)
        self.hideOldPacketsThread.start()
        
        # 导出解码列表时间戳格式设置
        self.exportDecodedListInMillisecTimestamp = False

        self.scanPorts()
        self.startTime = 0
        self.receivedPackets = 0
        self.playbackMainTableIndex = 0
        self.labelDictFile = None
        self.idDict = dict([])
        self.can1showOnlyIdsSet = set([])
        self.can2showOnlyIdsSet = set([])
        self.linshowOnlyIdsSet = set([])
        self.can1hideIdsSet = set([])
        self.can2hideIdsSet = set([])
        self.linhideIdsSet = set([])

        self.idLabelDict = dict()
        self.isInited = False
        self.init()
        
        # 创建保存文件夹
        if not os.path.exists("save"):
            os.makedirs("save")

        # 设置表格列宽
        #can2
        for i in range(5, self.can2MessageTableWidget.columnCount()):
            self.can2MessageTableWidget.setColumnWidth(i, 32)
        for i in range(5, self.can2MessageTableWidget.columnCount()):
            self.decodedMessagesTableWidget.setColumnWidth(i, 32)
        #can1
        for i in range(5, self.can1MessageTableWidget.columnCount()):
            self.can1MessageTableWidget.setColumnWidth(i, 32)
        for i in range(5, self.can1MessageTableWidget.columnCount()):
            self.decodedMessagesTableWidget.setColumnWidth(i, 32)

        #lin
        for i in range(5, self.linMessageTableWidget.columnCount()):
            self.linMessageTableWidget.setColumnWidth(i, 32)
        for i in range(5, self.linMessageTableWidget.columnCount()):
            self.decodedMessagesTableWidget.setColumnWidth(i, 32)

        self.decodedMessagesTableWidget.setColumnWidth(1, 150)
        self.decodedMessagesTableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.txTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.showNormal()
        # 然后再进入全屏
        #self.showFullScreen()
        self.showMaximized()
    '''   
    def event(self, event):
        print('-----event = ',event.type())
        # 捕获窗口被重新激活
        if event.type() == 24:
            print('-----ActivationChange = ',event.type())
            #if self.isMinimized():
            print("任务栏点击激活，恢复窗口")
            self.showNormal()
            self.raise_()
            self.activateWindow()
        return super().event(event,)
    def changeEvent(self, event):
        print('-----changeEvent = ',event.type())
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized():
                # 最小化时可以做一些处理
                self.update()  # 强制刷新缩略图
            elif self.isActiveWindow():
                # 恢复时强制显示
                self.showNormal()
        super().changeEvent(event)
    '''
    # 停止回放
    def stopPlayBackCallback(self):
        try:
            self.serialWriterThread.packetSentSignal.disconnect()
        except:
            pass
        self.serialWriterThread.clearQueues()
        self.playbackMainTableButton_can2.setVisible(True)
        self.stopPlayBackButton_can2.setVisible(False)
        self.playBackProgressBar_can2.setVisible(False)

    # 设置单选按钮状态
    def setRadioButton(self, radioButton:QRadioButton, mode):
        radioButton.setAutoExclusive(False)
        if mode == 0:
            radioButton.setChecked(False)
        if mode == 1:
            radioButton.setChecked(True)
        if mode == 2:
            radioButton.setChecked(not radioButton.isChecked())
        radioButton.setAutoExclusive(True)
        QApplication.processEvents()

    # 回放一条主表数据
    def playbackMainTable1Packet(self):
        row = self.playbackMainTableIndex

        if row < 0:
            self.stopPlayBackCallback()
            return
        maxRows = self.can2MessageTableWidget.rowCount()
        txBuf = ""
        id = ((self.can2MessageTableWidget.item(row, 1).text()).split(" "))[0]
        if len(id) % 2:
            txBuf += '0'
        txBuf += id + ',' + self.can2MessageTableWidget.item(row, 2).text() + ',' + \
                 self.can2MessageTableWidget.item(row, 3).text() + ','
        for i in range(5, self.can2MessageTableWidget.columnCount()):
            txBuf += self.can2MessageTableWidget.item(row, i).text()
        txBuf += '\n'
        if row < maxRows - 1:
            dt = float(self.can2MessageTableWidget.item(row, 0).text()) - float(
                self.can2MessageTableWidget.item(row + 1, 0).text())
            sec_to_ms = 1000
            if '.' not in self.can2MessageTableWidget.item(row, 0).text():
                sec_to_ms = 1       # 时间戳已为毫秒
            dt = abs(int(dt * sec_to_ms))
            self.serialWriterThread.setNormalWriteDelay(dt)
        self.playBackProgressBar_can2.setValue(int((maxRows - row) / maxRows * 100))
        self.playbackMainTableIndex -= 1

        self.serialWriterThread.write(txBuf)

    # 开始回放主表
    def playbackMainTableCallback(self):
        self.playbackMainTableButton_can2.setVisible(False)
        self.stopPlayBackButton_can2.setVisible(True)
        self.playBackProgressBar_can2.setVisible(True)
        self.playbackMainTableIndex = self.can2MessageTableWidget.rowCount() - 1
        self.serialWriterThread.setRepeatedWriteDelay(0)
        print('playing back...')
        self.serialWriterThread.packetSentSignal.connect(self.playbackMainTable1Packet)
        self.playbackMainTable1Packet()

    # 清空主表
    def clearTableCallback(self):
        self.idDict.clear()
        self.can2MessageTableWidget.setRowCount(0)

    # 发送解码包
    def sendDecodedPacketCallback(self):
        self.newTxTableRowCallback()
        newRow = 0
        decodedCurrentRow = self.decodedMessagesTableWidget.currentRow()
        newId = str(self.decodedMessagesTableWidget.item(decodedCurrentRow, 1).text()).split(" ")
        newItem = QTableWidgetItem(newId[0])
        self.txTable.setItem(newRow, 0, QTableWidgetItem(newItem))
        for i in range(1, 3):
            self.txTable.setItem(newRow, i, QTableWidgetItem(self.decodedMessagesTableWidget.item(decodedCurrentRow, i+1)))
        newData = ""
        for i in range(int(self.decodedMessagesTableWidget.item(decodedCurrentRow, 4).text())):
            newData += str(self.decodedMessagesTableWidget.item(decodedCurrentRow, 5 + i).text())
        self.txTable.setItem(newRow, 3, QTableWidgetItem(newData))
        self.txTable.selectRow(newRow)
        if self.sendTxTableButton.isEnabled():
            self.sendTxTableCallback()

    # 解码表内容变更回调
    def decodedTableItemChangedCallback(self):
        if self.isInited:
            self.saveTableToFile(self.decodedMessagesTableWidget, "save/decodedPackets.csv")

    # 删除解码行
    def deleteDecodedLineCallback(self):
        self.decodedMessagesTableWidget.removeRow(self.decodedMessagesTableWidget.currentRow())

    # 添加到解码表
    def addToDecodedCallback(self):
        newRow = self.decodedMessagesTableWidget.rowCount()
        self.decodedMessagesTableWidget.insertRow(newRow)
        for i in range(1, self.decodedMessagesTableWidget.columnCount()):
            new_item = QTableWidgetItem(self.can2MessageTableWidget.item(self.can2MessageTableWidget.currentRow(), i))
            self.decodedMessagesTableWidget.setItem(newRow, i, new_item)

    # 显示/隐藏发送表
    def showSendingTableButtonCallback(self):
        if self.showSendingTableCheckBox.isChecked():
            self.sendingGroupBox.show()
        else:
            self.sendingGroupBox.hide()

    # 隐藏所有包
    def hideAllPackets(self):
        text = ""
        for id in self.idDict:
            text += id + " "
        self.hideIdsLineEdit_can2.setText(text)
        self.clearTableCallback()

    # 隐藏旧包回调
    def hideOldPacketsCallback(self):
        if not self.hideOldPacketsCheckBox_can2.isChecked():
            return
        if not self.groupModeCheckBox.isChecked():
            return
        for i in range(self.can2MessageTableWidget.rowCount()):
            if self.can2MessageTableWidget.isRowHidden(i):
                continue
            packetTime = float(self.can2MessageTableWidget.item(i, 0).text())
            if (time.time() - self.startTime) - packetTime > self.hideOldPeriod.value():
                self.can2MessageTableWidget.setRowHidden(i, True)

    # 发送发送表数据
    def sendTxTableCallback(self):
        self.setRadioButton(self.txDataRadioButton, 2)
        for row in range(self.txTable.rowCount()):
            if self.txTable.item(row, 0).isSelected():
                txBuf = ""
                for i in range(self.txTable.columnCount()):
                    subStr = self.txTable.item(row, i).text() + ","
                    if not len(subStr) % 2:
                        subStr = '0' + subStr
                    txBuf += subStr
                txBuf = txBuf[:-1] + '\n'
                if self.repeatedDelayCheckBox.isChecked():
                    self.serialWriterThread.setRepeatedWriteDelay(self.repeatTxDelayValue.value())
                else:
                    self.serialWriterThread.setRepeatedWriteDelay(0)
                self.serialWriterThread.write(txBuf)

    # 文件加载完成回调
    def fileLoadingFinishedCallback(self):
        self.abortSessionLoadingButton.setEnabled(False)

    # 中止文件加载回调
    def abortSessionLoadingCallback(self):
        self.fileLoaderThread.stop()
        self.abortSessionLoadingButton.setEnabled(False)

    # 移除发送表行
    def removeTxTableRowCallback(self):
        try:
            self.txTable.removeRow(self.txTable.currentRow())
        except:
            print('cannot remove')

    # 新增发送表行
    def newTxTableRowCallback(self):
        newRow = 0
        self.txTable.insertRow(newRow)

    # 仅显示指定ID回调
    def can2showOnlyIdsTextChanged(self):
        self.can2showOnlyIdsSet.clear()
        self.can2showOnlyIdsSet = set(self.showOnlyIdsLineEdit_can2.text().split(" "))

    # 仅显示指定ID回调
    def can1showOnlyIdsTextChanged(self):
        self.can1showOnlyIdsSet.clear()
        self.can1showOnlyIdsSet = set(self.showOnlyIdsLineEdit_can1.text().split(" "))

    # 仅显示指定ID回调
    def linshowOnlyIdsTextChanged(self):
        self.linshowOnlyIdsSet.clear()
        self.linshowOnlyIdsSet = set(self.showOnlyIdsLineEdit_lin.text().split(" "))

    # 隐藏指定ID回调
    def can1hideIdsTextChanged(self):
        self.can1hideIdsSet.clear()
        self.can1hideIdsSet = set(self.hideIdsLineEdit_can1.text().split(" "))

    def can2hideIdsTextChanged(self):
        self.can2hideIdsSet.clear()
        self.can2hideIdsSet = set(self.hideIdsLineEdit_can2.text().split(" "))

    def linhideIdsTextChanged(self):
        self.linhideIdsSet.clear()
        self.linhideIdsSet = set(self.hideIdsLineEdit_lin.text().split(" "))
    # 初始化，加载解码表和ID标签字典
    def init(self):
        self.loadTableFromFile(self.decodedMessagesTableWidget, "save/decodedPackets.csv")
        self.loadTableFromFile(self.idLabelDictTable, "save/labelDict.csv")
        for row in range(self.idLabelDictTable.rowCount()):
            self.idLabelDict[str(self.idLabelDictTable.item(row, 0).text())] = \
                str(self.idLabelDictTable.item(row, 1).text())
        self.isInited = True

    # 清空ID标签字典
    def clearLabelDict(self):
        self.idLabelDictTable.setRowCount(0)
        self.saveTableToFile(self.idLabelDictTable, "save/labelDict.csv")

    # 保存表格到文件
    def saveTableToFile(self, table, path):
        if path is None:
            path, _ = QFileDialog.getSaveFileName(self, 'Save File', './save', 'CSV(*.csv)')
        if path != '':
            with open(str(path), 'w', newline='') as stream:
                writer = csv.writer(stream)
                for row in range(table.rowCount()-1, -1, -1):
                    rowData = []
                    for column in range(table.columnCount()):
                        item = table.item(row, column)
                        if item is not None:
                            tempItem = item.text()
                            if self.exportDecodedListInMillisecTimestamp and column == 0:
                                timeSplit = item.text().split('.')
                                sec = timeSplit[0]
                                ms = timeSplit[1][0:3]
                                tempItem = sec + ms
                            rowData.append(str(tempItem))
                        else:
                            rowData.append('')
                    writer.writerow(rowData)

    def fileLoaderCallback(self, rowData):
        self.mainTablePopulatorCallback(rowData)

    def can1TablePopulatorCallback(self, rowData):
        # 过滤显示和隐藏ID
        if self.showOnlyIdsCheckBox_can1.isChecked():
            if str(rowData[1]) not in self.can1showOnlyIdsSet:
                return
        if self.hideIdsCheckBox_can1.isChecked():
            if str(rowData[1]) in self.can1hideIdsSet:
                return

        newId = str(rowData[1])

        row = 0
        if self.groupModeCheckBox.isChecked():
            if newId in self.idDict.keys():
                row = self.idDict[newId]
            else:
                row = self.can1MessageTableWidget.rowCount()
                self.can1MessageTableWidget.insertRow(row)
        else:
            self.can1MessageTableWidget.insertRow(row)

        if self.can1MessageTableWidget.isRowHidden(row):
            self.can1MessageTableWidget.setRowHidden(row, False)

        for i in range(self.can1MessageTableWidget.columnCount()):
            if i < len(rowData):
                data = str(rowData[i])
                item = self.can1MessageTableWidget.item(row, i)
                newItem = QTableWidgetItem(data)
                if item:
                    if item.text() != data:
                        if self.highlightNewDataCheckBox_can1.isChecked() and \
                                self.groupModeCheckBox.isChecked() and \
                                i > 4:
                            newItem.setBackground(QColor(104, 37, 98))
                else:
                    if self.highlightNewDataCheckBox_can1.isChecked() and \
                            self.groupModeCheckBox.isChecked() and \
                            i > 4:
                        newItem.setBackground(QColor(104, 37, 98))
            else:
                newItem = QTableWidgetItem()
            self.can1MessageTableWidget.setItem(row, i, newItem)

        isFamiliar = False

        if self.highlightNewIdCheckBox_can1.isChecked():
            if newId not in self.idDict.keys():
                for j in range(3):
                    self.can1MessageTableWidget.item(row, j).setBackground(QColor(52, 44, 124))

        self.idDict[newId] = row

        if newId in self.idLabelDict.keys():
            value = newId + " (" + self.idLabelDict[newId] + ")"
            self.can1MessageTableWidget.setItem(row, 1, QTableWidgetItem(value))
            isFamiliar = True

        for i in range(self.can1MessageTableWidget.columnCount()):
            if (isFamiliar or (newId.find("(") >= 0)) and i < 3:
                self.can1MessageTableWidget.item(row, i).setBackground(QColor(53, 81, 52))

            self.can1MessageTableWidget.item(row, i).setTextAlignment(Qt.AlignVCenter | Qt.AlignHCenter)

        self.receivedPackets = self.receivedPackets + 1
        self.packageCounterLabel.setText(str(self.receivedPackets))

     # 主表填充回调
    def can2TablePopulatorCallback(self, rowData):
        # 过滤显示和隐藏ID
        if self.showOnlyIdsCheckBox_can2.isChecked():
            if str(rowData[1]) not in self.can2showOnlyIdsSet:
                return
        if self.hideIdsCheckBox_can2.isChecked():
            if str(rowData[1]) in self.can2hideIdsSet:
                return

        newId = str(rowData[1])

        row = 0
        if self.groupModeCheckBox.isChecked():
            if newId in self.idDict.keys():
                row = self.idDict[newId]
            else:
                row = self.can2MessageTableWidget.rowCount()
                self.can2MessageTableWidget.insertRow(row)
        else:
            self.can2MessageTableWidget.insertRow(row)

        if self.can2MessageTableWidget.isRowHidden(row):
            self.can2MessageTableWidget.setRowHidden(row, False)

        for i in range(self.can2MessageTableWidget.columnCount()):
            if i < len(rowData):
                data = str(rowData[i])
                item = self.can2MessageTableWidget.item(row, i)
                newItem = QTableWidgetItem(data)
                if item:
                    if item.text() != data:
                        if self.highlightNewDataCheckBox_can2.isChecked() and \
                                self.groupModeCheckBox.isChecked() and \
                                i > 4:
                            newItem.setBackground(QColor(104, 37, 98))
                else:
                    if self.highlightNewDataCheckBox_can2.isChecked() and \
                            self.groupModeCheckBox.isChecked() and \
                            i > 4:
                        newItem.setBackground(QColor(104, 37, 98))
            else:
                newItem = QTableWidgetItem()
            self.can2MessageTableWidget.setItem(row, i, newItem)

        isFamiliar = False

        if self.highlightNewIdCheckBox_can2.isChecked():
            if newId not in self.idDict.keys():
                for j in range(3):
                    self.can2MessageTableWidget.item(row, j).setBackground(QColor(52, 44, 124))

        self.idDict[newId] = row

        if newId in self.idLabelDict.keys():
            value = newId + " (" + self.idLabelDict[newId] + ")"
            self.can2MessageTableWidget.setItem(row, 1, QTableWidgetItem(value))
            isFamiliar = True

        for i in range(self.can2MessageTableWidget.columnCount()):
            if (isFamiliar or (newId.find("(") >= 0)) and i < 3:
                self.can2MessageTableWidget.item(row, i).setBackground(QColor(53, 81, 52))

            self.can2MessageTableWidget.item(row, i).setTextAlignment(Qt.AlignVCenter | Qt.AlignHCenter)

        self.receivedPackets = self.receivedPackets + 1
        self.packageCounterLabel.setText(str(self.receivedPackets))

     # 主表填充回调
    def linTablePopulatorCallback(self, rowData):
        # 过滤显示和隐藏ID
        if self.showOnlyIdsCheckBox_lin.isChecked():
            if str(rowData[1]) not in self.linshowOnlyIdsSet:
                return
        if self.hideIdsCheckBox_lin.isChecked():
            if str(rowData[1]) in self.linhideIdsSet:
                return

        newId = str(rowData[1])

        row = 0
        if self.groupModeCheckBox.isChecked():
            if newId in self.idDict.keys():
                row = self.idDict[newId]
            else:
                row = self.linMessageTableWidget.rowCount()
                self.linMessageTableWidget.insertRow(row)
        else:
            self.linMessageTableWidget.insertRow(row)

        if self.linMessageTableWidget.isRowHidden(row):
            self.linMessageTableWidget.setRowHidden(row, False)

        for i in range(self.linMessageTableWidget.columnCount()):
            if i < len(rowData):
                data = str(rowData[i])
                item = self.linMessageTableWidget.item(row, i)
                newItem = QTableWidgetItem(data)
                if item:
                    if item.text() != data:
                        if self.highlightNewDataCheckBox_lin.isChecked() and \
                                self.groupModeCheckBox.isChecked() and \
                                i > 4:
                            newItem.setBackground(QColor(104, 37, 98))
                else:
                    if self.highlightNewDataCheckBox_lin.isChecked() and \
                            self.groupModeCheckBox.isChecked() and \
                            i > 4:
                        newItem.setBackground(QColor(104, 37, 98))
            else:
                newItem = QTableWidgetItem()
            self.linMessageTableWidget.setItem(row, i, newItem)

        isFamiliar = False

        if self.highlightNewIdCheckBox_lin.isChecked():
            if newId not in self.idDict.keys():
                for j in range(3):
                    self.linMessageTableWidget.item(row, j).setBackground(QColor(52, 44, 124))

        self.idDict[newId] = row

        if newId in self.idLabelDict.keys():
            value = newId + " (" + self.idLabelDict[newId] + ")"
            self.linMessageTableWidget.setItem(row, 1, QTableWidgetItem(value))
            isFamiliar = True

        for i in range(self.linMessageTableWidget.columnCount()):
            if (isFamiliar or (newId.find("(") >= 0)) and i < 3:
                self.linMessageTableWidget.item(row, i).setBackground(QColor(53, 81, 52))

            self.linMessageTableWidget.item(row, i).setTextAlignment(Qt.AlignVCenter | Qt.AlignHCenter)

        self.receivedPackets = self.receivedPackets + 1
        self.packageCounterLabel.setText(str(self.receivedPackets))

    # 主表填充回调
    def mainTablePopulatorCallback(self, rowData):
        if rowData[1][0] == 'A':
            rowData[1] = rowData[1][1:]
            self.can1TablePopulatorCallback(rowData)
        elif rowData[1][0] == 'B':
            rowData[1] = rowData[1][1:]
            self.can2TablePopulatorCallback(rowData)
        elif rowData[1][0] == 'L':
            rowData[1] = rowData[1][1:]
            self.linTablePopulatorCallback(rowData)
        

    # 从文件加载表格
    def loadTableFromFile(self, table, path):
        if path is None:
            path, _ = QFileDialog.getOpenFileName(self, 'Open File', './save', 'CSV(*.csv)')
        if path != '':
            if table == self.can2MessageTableWidget:
                self.fileLoaderThread.start()
                self.fileLoaderThread.enable(path, self.playbackDelaySpinBox.value())
                self.abortSessionLoadingButton.setEnabled(True)
                return True
            try:
                with open(str(path), 'r',encoding='utf-8') as stream:
                    for rowData in csv.reader(stream):
                        row = table.rowCount()
                        table.insertRow(row)
                        for i in range(len(rowData)):
                            if len(rowData[i]):
                                item = QTableWidgetItem(str(rowData[i]))
                                if not (table == self.decodedMessagesTableWidget and i == 0):
                                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignHCenter)
                                table.setItem(row, i, item)
            except OSError:
                print("file not found: " + path)

    # 从文件加载会话
    def loadSessionFromFile(self):
        if self.autoclearCheckBox.isChecked():
            self.idDict.clear()
            self.can2MessageTableWidget.setRowCount(0)
        self.loadTableFromFile(self.can2MessageTableWidget, None)

    # 保存会话到文件
    def saveSessionToFile(self):
        self.saveTableToFile(self.can2MessageTableWidget, None)

    # 单元格点击回调
    def cellWasClicked(self):
        self.saveIdToDictLineEdit_can2.setText(self.can2MessageTableWidget.item(self.can2MessageTableWidget.currentRow(), 1).text())

    # 保存ID标签到字典
    def saveIdLabelToDictCallback(self):
        if (not self.saveIdToDictLineEdit_can2.text()) or (not self.saveLabelToDictLineEdit_can2.text()):
            return
        newRow = self.idLabelDictTable.rowCount()
        self.idLabelDictTable.insertRow(newRow)
        widgetItem = QTableWidgetItem()
        widgetItem.setTextAlignment(Qt.AlignVCenter | Qt.AlignHCenter)
        widgetItem.setText(self.saveIdToDictLineEdit_can2.text())
        self.idLabelDictTable.setItem(newRow, 0, QTableWidgetItem(widgetItem))
        widgetItem.setText(self.saveLabelToDictLineEdit_can2.text())
        self.idLabelDictTable.setItem(newRow, 1, QTableWidgetItem(widgetItem))
        self.idLabelDict[str(self.saveIdToDictLineEdit_can2.text())] = str(self.saveLabelToDictLineEdit_can2.text())
        self.saveIdToDictLineEdit_can2.setText('')
        self.saveLabelToDictLineEdit_can2.setText('')
        self.saveTableToFile(self.idLabelDictTable, "save/labelDict.csv")

    # 开始嗅探
    def startSniffing(self):
        if self.autoclearCheckBox.isChecked():
            self.idDict.clear()
            self.can1MessageTableWidget.setRowCount(0)
            self.can2MessageTableWidget.setRowCount(0)
            self.linMessageTableWidget.setRowCount(0)
        self.startSniffingButton.setEnabled(False)
        self.stopSniffingButton.setEnabled(True)
        self.sendTxTableButton.setEnabled(True)
        self.activeChannelComboBox.setEnabled(False)

        if self.activeChannelComboBox.isEnabled():
            txBuf = [0x42, self.activeChannelComboBox.currentIndex()]   # TX FORWARDER
            self.serialWriterThread.write(txBuf)
            txBuf = [0x41, 1 << self.activeChannelComboBox.currentIndex()]  # RX FORWARDER
            self.serialWriterThread.write(txBuf)

        self.startTime = time.time()

    # 停止嗅探
    def stopSniffing(self):
        self.startSniffingButton.setEnabled(True)
        self.stopSniffingButton.setEnabled(False)
        self.sendTxTableButton.setEnabled(False)
        self.activeChannelComboBox.setEnabled(True)
        self.setRadioButton(self.rxDataRadioButton, 0)

    # 串口数据包接收回调
    def serialPacketReceiverCallback(self, packet, time):
        if self.startSniffingButton.isEnabled():
            return
        packetSplit = packet[:-1].split(',')

        if len(packetSplit) != 4:
            print("wrong packet!" + packet)
            self.snifferMsgPlainTextEdit.document().setPlainText(packet)
            return

        rowData = [str(time - self.startTime)[:7]]  # 时间戳
        rowData += packetSplit[0:3]  # IDE, RTR, EXT
        DLC = len(packetSplit[3]) // 2
        rowData.append(str("{:02X}".format(DLC)))  # DLC
        if DLC > 0:
            rowData += [packetSplit[3][i:i + 2] for i in range(0, len(packetSplit[3]), 2)]  # 数据

        self.mainTablePopulatorCallback(rowData)

    # 串口连接
    def serialPortConnect(self):
        try:
            self.serialController.port = self.portSelectorComboBox.currentText()
            self.serialController.baudrate = 250000
            self.serialController.open()
            self.serialReaderThread.start()
            self.serialWriterThread.start()
            self.serialConnectedCheckBox.setChecked(True)
            self.portDisconnectButton.setEnabled(True)
            self.portConnectButton.setEnabled(False)
            self.startSniffingButton.setEnabled(True)
            self.stopSniffingButton.setEnabled(False)
        except serial.SerialException as e:
            print('Error opening port: ' + str(e))

    # 串口断开
    def serialPortDisconnect(self):
        if self.stopSniffingButton.isEnabled():
            self.stopSniffing()
        try:
            self.serialReaderThread.stop()
            self.serialWriterThread.stop()
            self.portDisconnectButton.setEnabled(False)
            self.portConnectButton.setEnabled(True)
            self.startSniffingButton.setEnabled(False)
            self.serialConnectedCheckBox.setChecked(False)
            self.serialController.close()
        except serial.SerialException as e:
            print('Error closing port: ' + str(e))

    # 扫描串口
    def scanPorts(self):
        self.portSelectorComboBox.clear()
        comPorts = serial.tools.list_ports.comports()
        nameList = list(port.device for port in comPorts)
        for name in nameList:
            self.portSelectorComboBox.addItem(name)


# 异常钩子
def exception_hook(exctype, value, traceback):
    print(exctype, value, traceback)
    sys._excepthook(exctype, value, traceback)
    sys.exit(1)

# 主函数入口
def main():
    # 重定向异常钩子
    sys._excepthook = sys.excepthook
    sys.excepthook = exception_hook

    # 创建应用
    app = QApplication(sys.argv)
    gui = canSnifferGUI()

    # 应用暗色主题
    qtmodern.styles.dark(app)
    darked_gui = qtmodern.windows.ModernWindow(gui)

    # 添加窗口缩放控件
    layout = QVBoxLayout()
    sizegrip = QSizeGrip(darked_gui)
    sizegrip.setMaximumSize(30, 30)
    layout.addWidget(sizegrip, 50, Qt.AlignBottom | Qt.AlignRight)
    darked_gui.setLayout(layout)

    # 启动应用
    darked_gui.show()
    app.exec_()


if __name__ == "__main__":
    main()
