# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'move_win_.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1205, 808)
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.frame_8 = QtWidgets.QFrame(Form)
        self.frame_8.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_8.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_8.setObjectName("frame_8")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.frame_8)
        self.horizontalLayout_10.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.frame_7 = QtWidgets.QFrame(self.frame_8)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame_7.sizePolicy().hasHeightForWidth())
        self.frame_7.setSizePolicy(sizePolicy)
        self.frame_7.setMinimumSize(QtCore.QSize(0, 0))
        self.frame_7.setMaximumSize(QtCore.QSize(600, 510))
        self.frame_7.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_7.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_7.setObjectName("frame_7")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.frame_7)
        self.verticalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.frFile = QtWidgets.QFrame(self.frame_7)
        self.frFile.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frFile.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frFile.setObjectName("frFile")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout(self.frFile)
        self.horizontalLayout_9.setContentsMargins(9, 9, 9, 9)
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.leDir = QtWidgets.QLineEdit(self.frFile)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.leDir.sizePolicy().hasHeightForWidth())
        self.leDir.setSizePolicy(sizePolicy)
        self.leDir.setObjectName("leDir")
        self.horizontalLayout_9.addWidget(self.leDir)
        self.cmbFile = QtWidgets.QComboBox(self.frFile)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cmbFile.sizePolicy().hasHeightForWidth())
        self.cmbFile.setSizePolicy(sizePolicy)
        self.cmbFile.setObjectName("cmbFile")
        self.horizontalLayout_9.addWidget(self.cmbFile)
        self.cmbTab = QtWidgets.QComboBox(self.frFile)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cmbTab.sizePolicy().hasHeightForWidth())
        self.cmbTab.setSizePolicy(sizePolicy)
        self.cmbTab.setObjectName("cmbTab")
        self.horizontalLayout_9.addWidget(self.cmbTab)
        self.verticalLayout_3.addWidget(self.frFile)
        self.frame_6 = QtWidgets.QFrame(self.frame_7)
        self.frame_6.setMaximumSize(QtCore.QSize(16777215, 120))
        self.frame_6.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_6.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_6.setObjectName("frame_6")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.frame_6)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.frFond = QtWidgets.QFrame(self.frame_6)
        self.frFond.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frFond.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frFond.setLineWidth(5)
        self.frFond.setObjectName("frFond")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.frFond)
        self.verticalLayout_2.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.pbFond = QtWidgets.QPushButton(self.frFond)
        self.pbFond.setObjectName("pbFond")
        self.verticalLayout_2.addWidget(self.pbFond)
        self.frame_3 = QtWidgets.QFrame(self.frFond)
        self.frame_3.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.frame_3)
        self.horizontalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.clbXFond = QtWidgets.QCommandLinkButton(self.frame_3)
        self.clbXFond.setMinimumSize(QtCore.QSize(20, 20))
        self.clbXFond.setMaximumSize(QtCore.QSize(20, 16777215))
        self.clbXFond.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("x.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbXFond.setIcon(icon)
        self.clbXFond.setIconSize(QtCore.QSize(10, 10))
        self.clbXFond.setObjectName("clbXFond")
        self.horizontalLayout_5.addWidget(self.clbXFond)
        self.leFond = QtWidgets.QLineEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.leFond.sizePolicy().hasHeightForWidth())
        self.leFond.setSizePolicy(sizePolicy)
        self.leFond.setObjectName("leFond")
        self.horizontalLayout_5.addWidget(self.leFond)
        self.verticalLayout_2.addWidget(self.frame_3)
        self.cmbFond = QtWidgets.QComboBox(self.frFond)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cmbFond.sizePolicy().hasHeightForWidth())
        self.cmbFond.setSizePolicy(sizePolicy)
        self.cmbFond.setObjectName("cmbFond")
        self.verticalLayout_2.addWidget(self.cmbFond)
        self.horizontalLayout_2.addWidget(self.frFond)
        self.frAgent = QtWidgets.QFrame(self.frame_6)
        self.frAgent.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frAgent.setFrameShadow(QtWidgets.QFrame.Plain)
        self.frAgent.setLineWidth(5)
        self.frAgent.setMidLineWidth(0)
        self.frAgent.setObjectName("frAgent")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.frAgent)
        self.verticalLayout_4.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.pbAgent = QtWidgets.QPushButton(self.frAgent)
        self.pbAgent.setObjectName("pbAgent")
        self.verticalLayout_4.addWidget(self.pbAgent)
        self.frame_4 = QtWidgets.QFrame(self.frAgent)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame_4.sizePolicy().hasHeightForWidth())
        self.frame_4.setSizePolicy(sizePolicy)
        self.frame_4.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_4.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_4.setObjectName("frame_4")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout(self.frame_4)
        self.horizontalLayout_6.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.clbXAgent = QtWidgets.QCommandLinkButton(self.frame_4)
        self.clbXAgent.setMinimumSize(QtCore.QSize(12, 12))
        self.clbXAgent.setMaximumSize(QtCore.QSize(20, 16777215))
        self.clbXAgent.setText("")
        self.clbXAgent.setIcon(icon)
        self.clbXAgent.setIconSize(QtCore.QSize(10, 10))
        self.clbXAgent.setObjectName("clbXAgent")
        self.horizontalLayout_6.addWidget(self.clbXAgent)
        self.leAgent = QtWidgets.QLineEdit(self.frame_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.leAgent.sizePolicy().hasHeightForWidth())
        self.leAgent.setSizePolicy(sizePolicy)
        self.leAgent.setObjectName("leAgent")
        self.horizontalLayout_6.addWidget(self.leAgent)
        self.verticalLayout_4.addWidget(self.frame_4)
        self.cmbAgent = QtWidgets.QComboBox(self.frAgent)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cmbAgent.sizePolicy().hasHeightForWidth())
        self.cmbAgent.setSizePolicy(sizePolicy)
        self.cmbAgent.setObjectName("cmbAgent")
        self.verticalLayout_4.addWidget(self.cmbAgent)
        self.horizontalLayout_2.addWidget(self.frAgent)
        self.frSigner = QtWidgets.QFrame(self.frame_6)
        self.frSigner.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frSigner.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frSigner.setObjectName("frSigner")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.frSigner)
        self.verticalLayout_5.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.pbSigner = QtWidgets.QPushButton(self.frSigner)
        self.pbSigner.setObjectName("pbSigner")
        self.verticalLayout_5.addWidget(self.pbSigner)
        self.frame_5 = QtWidgets.QFrame(self.frSigner)
        self.frame_5.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_5.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_5.setObjectName("frame_5")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.frame_5)
        self.horizontalLayout_7.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.clbXSigner = QtWidgets.QCommandLinkButton(self.frame_5)
        self.clbXSigner.setMinimumSize(QtCore.QSize(20, 20))
        self.clbXSigner.setMaximumSize(QtCore.QSize(20, 16777215))
        self.clbXSigner.setText("")
        self.clbXSigner.setIcon(icon)
        self.clbXSigner.setIconSize(QtCore.QSize(10, 10))
        self.clbXSigner.setObjectName("clbXSigner")
        self.horizontalLayout_7.addWidget(self.clbXSigner)
        self.leSigner = QtWidgets.QLineEdit(self.frame_5)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.leSigner.sizePolicy().hasHeightForWidth())
        self.leSigner.setSizePolicy(sizePolicy)
        self.leSigner.setObjectName("leSigner")
        self.horizontalLayout_7.addWidget(self.leSigner)
        self.verticalLayout_5.addWidget(self.frame_5)
        self.cmbSigner = QtWidgets.QComboBox(self.frSigner)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cmbSigner.sizePolicy().hasHeightForWidth())
        self.cmbSigner.setSizePolicy(sizePolicy)
        self.cmbSigner.setObjectName("cmbSigner")
        self.verticalLayout_5.addWidget(self.cmbSigner)
        self.horizontalLayout_2.addWidget(self.frSigner)
        self.verticalLayout_3.addWidget(self.frame_6)
        self.frAll = QtWidgets.QFrame(self.frame_7)
        self.frAll.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frAll.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frAll.setObjectName("frAll")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.frAll)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.frame = QtWidgets.QFrame(self.frAll)
        self.frame.setMinimumSize(QtCore.QSize(0, 0))
        self.frame.setMaximumSize(QtCore.QSize(320, 16777215))
        self.frame.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.frame)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.chbClientOnly = QtWidgets.QCheckBox(self.frame)
        self.chbClientOnly.setObjectName("chbClientOnly")
        self.verticalLayout.addWidget(self.chbClientOnly)
        self.chbSocium = QtWidgets.QCheckBox(self.frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbSocium.sizePolicy().hasHeightForWidth())
        self.chbSocium.setSizePolicy(sizePolicy)
        self.chbSocium.setObjectName("chbSocium")
        self.verticalLayout.addWidget(self.chbSocium)
        self.frSuff = QtWidgets.QFrame(self.frame)
        self.frSuff.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frSuff.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frSuff.setLineWidth(0)
        self.frSuff.setObjectName("frSuff")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.frSuff)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.chbSuff = QtWidgets.QCheckBox(self.frSuff)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbSuff.sizePolicy().hasHeightForWidth())
        self.chbSuff.setSizePolicy(sizePolicy)
        self.chbSuff.setObjectName("chbSuff")
        self.horizontalLayout.addWidget(self.chbSuff)
        self.leSuff = QtWidgets.QLineEdit(self.frSuff)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.leSuff.sizePolicy().hasHeightForWidth())
        self.leSuff.setSizePolicy(sizePolicy)
        self.leSuff.setMinimumSize(QtCore.QSize(0, 0))
        self.leSuff.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.leSuff.setObjectName("leSuff")
        self.horizontalLayout.addWidget(self.leSuff)
        self.verticalLayout.addWidget(self.frSuff)
        self.chbOurStat = QtWidgets.QCheckBox(self.frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbOurStat.sizePolicy().hasHeightForWidth())
        self.chbOurStat.setSizePolicy(sizePolicy)
        self.chbOurStat.setObjectName("chbOurStat")
        self.verticalLayout.addWidget(self.chbOurStat)
        self.chbFondStat = QtWidgets.QCheckBox(self.frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbFondStat.sizePolicy().hasHeightForWidth())
        self.chbFondStat.setSizePolicy(sizePolicy)
        self.chbFondStat.setObjectName("chbFondStat")
        self.verticalLayout.addWidget(self.chbFondStat)
        self.chbArhivON = QtWidgets.QCheckBox(self.frame)
        self.chbArhivON.setObjectName("chbArhivON")
        self.verticalLayout.addWidget(self.chbArhivON)
        self.chbArhivOFF = QtWidgets.QCheckBox(self.frame)
        self.chbArhivOFF.setObjectName("chbArhivOFF")
        self.verticalLayout.addWidget(self.chbArhivOFF)
        self.horizontalLayout_4.addWidget(self.frame)
        self.frImportInf = QtWidgets.QFrame(self.frAll)
        self.frImportInf.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frImportInf.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frImportInf.setObjectName("frImportInf")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout(self.frImportInf)
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.lbmportInf = QtWidgets.QLabel(self.frImportInf)
        self.lbmportInf.setText("")
        self.lbmportInf.setObjectName("lbmportInf")
        self.verticalLayout_11.addWidget(self.lbmportInf)
        self.cmbGenderType = QtWidgets.QComboBox(self.frImportInf)
        self.cmbGenderType.setObjectName("cmbGenderType")
        self.verticalLayout_11.addWidget(self.cmbGenderType)
        self.horizontalLayout_4.addWidget(self.frImportInf)
        self.frMoveInf = QtWidgets.QFrame(self.frAll)
        self.frMoveInf.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frMoveInf.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frMoveInf.setObjectName("frMoveInf")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout(self.frMoveInf)
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.lbMoveInf = QtWidgets.QLabel(self.frMoveInf)
        self.lbMoveInf.setText("")
        self.lbMoveInf.setObjectName("lbMoveInf")
        self.verticalLayout_12.addWidget(self.lbMoveInf)
        self.horizontalLayout_4.addWidget(self.frMoveInf)
        self.frPasportInf = QtWidgets.QFrame(self.frAll)
        self.frPasportInf.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frPasportInf.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frPasportInf.setObjectName("frPasportInf")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout(self.frPasportInf)
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.lbPasportInf = QtWidgets.QLabel(self.frPasportInf)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lbPasportInf.sizePolicy().hasHeightForWidth())
        self.lbPasportInf.setSizePolicy(sizePolicy)
        self.lbPasportInf.setText("")
        self.lbPasportInf.setObjectName("lbPasportInf")
        self.verticalLayout_9.addWidget(self.lbPasportInf)
        self.frDateFrom = QtWidgets.QFrame(self.frPasportInf)
        self.frDateFrom.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frDateFrom.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frDateFrom.setLineWidth(0)
        self.frDateFrom.setObjectName("frDateFrom")
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout(self.frDateFrom)
        self.horizontalLayout_12.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.chbDateFrom = QtWidgets.QCheckBox(self.frDateFrom)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbDateFrom.sizePolicy().hasHeightForWidth())
        self.chbDateFrom.setSizePolicy(sizePolicy)
        self.chbDateFrom.setMinimumSize(QtCore.QSize(110, 0))
        self.chbDateFrom.setObjectName("chbDateFrom")
        self.horizontalLayout_12.addWidget(self.chbDateFrom)
        self.deDateFrom = QtWidgets.QDateEdit(self.frDateFrom)
        self.deDateFrom.setEnabled(False)
        self.deDateFrom.setDateTime(QtCore.QDateTime(QtCore.QDate(2018, 1, 1), QtCore.QTime(0, 0, 0)))
        self.deDateFrom.setMinimumDateTime(QtCore.QDateTime(QtCore.QDate(1752, 9, 14), QtCore.QTime(0, 0, 0)))
        self.deDateFrom.setCalendarPopup(True)
        self.deDateFrom.setObjectName("deDateFrom")
        self.horizontalLayout_12.addWidget(self.deDateFrom)
        self.verticalLayout_9.addWidget(self.frDateFrom)
        self.frDateTo = QtWidgets.QFrame(self.frPasportInf)
        self.frDateTo.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frDateTo.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frDateTo.setLineWidth(0)
        self.frDateTo.setObjectName("frDateTo")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout(self.frDateTo)
        self.horizontalLayout_13.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.chbDateTo = QtWidgets.QCheckBox(self.frDateTo)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.chbDateTo.sizePolicy().hasHeightForWidth())
        self.chbDateTo.setSizePolicy(sizePolicy)
        self.chbDateTo.setMinimumSize(QtCore.QSize(110, 0))
        self.chbDateTo.setObjectName("chbDateTo")
        self.horizontalLayout_13.addWidget(self.chbDateTo)
        self.deDateTo = QtWidgets.QDateEdit(self.frDateTo)
        self.deDateTo.setEnabled(False)
        self.deDateTo.setDateTime(QtCore.QDateTime(QtCore.QDate(2018, 1, 1), QtCore.QTime(0, 0, 0)))
        self.deDateTo.setCalendarPopup(True)
        self.deDateTo.setObjectName("deDateTo")
        self.horizontalLayout_13.addWidget(self.deDateTo)
        self.verticalLayout_9.addWidget(self.frDateTo)
        self.horizontalLayout_4.addWidget(self.frPasportInf)
        self.verticalLayout_3.addWidget(self.frAll)
        self.frSQLcl = QtWidgets.QFrame(self.frame_7)
        self.frSQLcl.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frSQLcl.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frSQLcl.setLineWidth(0)
        self.frSQLcl.setObjectName("frSQLcl")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.frSQLcl)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.lbSQLcl = QtWidgets.QLabel(self.frSQLcl)
        self.lbSQLcl.setObjectName("lbSQLcl")
        self.horizontalLayout_3.addWidget(self.lbSQLcl)
        self.leSQLcl = QtWidgets.QLineEdit(self.frSQLcl)
        self.leSQLcl.setObjectName("leSQLcl")
        self.horizontalLayout_3.addWidget(self.leSQLcl)
        self.verticalLayout_3.addWidget(self.frSQLcl)
        self.frSQLco = QtWidgets.QFrame(self.frame_7)
        self.frSQLco.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frSQLco.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frSQLco.setLineWidth(0)
        self.frSQLco.setObjectName("frSQLco")
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout(self.frSQLco)
        self.horizontalLayout_8.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.lbSQLco = QtWidgets.QLabel(self.frSQLco)
        self.lbSQLco.setObjectName("lbSQLco")
        self.horizontalLayout_8.addWidget(self.lbSQLco)
        self.leSQLco = QtWidgets.QLineEdit(self.frSQLco)
        self.leSQLco.setObjectName("leSQLco")
        self.horizontalLayout_8.addWidget(self.leSQLco)
        self.verticalLayout_3.addWidget(self.frSQLco)
        self.frImport = QtWidgets.QFrame(self.frame_7)
        self.frImport.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frImport.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frImport.setObjectName("frImport")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout(self.frImport)
        self.verticalLayout_7.setContentsMargins(-1, 0, -1, 0)
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.frame_9 = QtWidgets.QFrame(self.frImport)
        self.frame_9.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_9.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_9.setObjectName("frame_9")
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout(self.frame_9)
        self.horizontalLayout_11.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.pbRefreshImport = QtWidgets.QPushButton(self.frame_9)
        self.pbRefreshImport.setMinimumSize(QtCore.QSize(100, 0))
        self.pbRefreshImport.setObjectName("pbRefreshImport")
        self.horizontalLayout_11.addWidget(self.pbRefreshImport)
        spacerItem = QtWidgets.QSpacerItem(304, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_11.addItem(spacerItem)
        self.clbImport = QtWidgets.QCommandLinkButton(self.frame_9)
        self.clbImport.setMaximumSize(QtCore.QSize(45, 45))
        self.clbImport.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("import.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbImport.setIcon(icon1)
        self.clbImport.setIconSize(QtCore.QSize(30, 30))
        self.clbImport.setObjectName("clbImport")
        self.horizontalLayout_11.addWidget(self.clbImport)
        spacerItem1 = QtWidgets.QSpacerItem(304, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_11.addItem(spacerItem1)
        self.pbImport = QtWidgets.QPushButton(self.frame_9)
        self.pbImport.setMinimumSize(QtCore.QSize(100, 0))
        self.pbImport.setObjectName("pbImport")
        self.horizontalLayout_11.addWidget(self.pbImport)
        self.verticalLayout_7.addWidget(self.frame_9)
        self.verticalLayout_3.addWidget(self.frImport)
        self.frPasport = QtWidgets.QFrame(self.frame_7)
        self.frPasport.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frPasport.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frPasport.setObjectName("frPasport")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout(self.frPasport)
        self.verticalLayout_10.setContentsMargins(-1, 0, -1, 0)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.frame_10 = QtWidgets.QFrame(self.frPasport)
        self.frame_10.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame_10.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_10.setObjectName("frame_10")
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout(self.frame_10)
        self.horizontalLayout_14.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.pbIdGenerate4PasportCheck = QtWidgets.QPushButton(self.frame_10)
        self.pbIdGenerate4PasportCheck.setMinimumSize(QtCore.QSize(100, 0))
        self.pbIdGenerate4PasportCheck.setObjectName("pbIdGenerate4PasportCheck")
        self.horizontalLayout_14.addWidget(self.pbIdGenerate4PasportCheck)
        spacerItem2 = QtWidgets.QSpacerItem(304, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_14.addItem(spacerItem2)
        self.clbPasport = QtWidgets.QCommandLinkButton(self.frame_10)
        self.clbPasport.setMaximumSize(QtCore.QSize(45, 45))
        self.clbPasport.setText("")
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("orel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbPasport.setIcon(icon2)
        self.clbPasport.setIconSize(QtCore.QSize(30, 30))
        self.clbPasport.setObjectName("clbPasport")
        self.horizontalLayout_14.addWidget(self.clbPasport)
        spacerItem3 = QtWidgets.QSpacerItem(304, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_14.addItem(spacerItem3)
        self.pbPasportCheck = QtWidgets.QPushButton(self.frame_10)
        self.pbPasportCheck.setMinimumSize(QtCore.QSize(100, 0))
        self.pbPasportCheck.setObjectName("pbPasportCheck")
        self.horizontalLayout_14.addWidget(self.pbPasportCheck)
        self.verticalLayout_10.addWidget(self.frame_10)
        self.verticalLayout_3.addWidget(self.frPasport)
        self.frMove = QtWidgets.QFrame(self.frame_7)
        self.frMove.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frMove.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frMove.setObjectName("frMove")
        self.gridLayout = QtWidgets.QGridLayout(self.frMove)
        self.gridLayout.setContentsMargins(-1, 0, -1, 0)
        self.gridLayout.setObjectName("gridLayout")
        spacerItem4 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem4, 1, 1, 1, 1)
        self.pbRefresh = QtWidgets.QPushButton(self.frMove)
        self.pbRefresh.setMinimumSize(QtCore.QSize(100, 0))
        self.pbRefresh.setObjectName("pbRefresh")
        self.gridLayout.addWidget(self.pbRefresh, 1, 0, 1, 1)
        self.pbMove = QtWidgets.QPushButton(self.frMove)
        self.pbMove.setMinimumSize(QtCore.QSize(100, 0))
        self.pbMove.setObjectName("pbMove")
        self.gridLayout.addWidget(self.pbMove, 1, 4, 1, 1)
        self.clbMove = QtWidgets.QCommandLinkButton(self.frMove)
        self.clbMove.setMaximumSize(QtCore.QSize(45, 45))
        self.clbMove.setText("")
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("move.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbMove.setIcon(icon3)
        self.clbMove.setIconSize(QtCore.QSize(30, 30))
        self.clbMove.setObjectName("clbMove")
        self.gridLayout.addWidget(self.clbMove, 1, 2, 1, 1)
        spacerItem5 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem5, 1, 3, 1, 1)
        self.verticalLayout_3.addWidget(self.frMove)
        self.progressBar = QtWidgets.QProgressBar(self.frame_7)
        self.progressBar.setMaximum(0)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_3.addWidget(self.progressBar)
        self.horizontalLayout_10.addWidget(self.frame_7)
        self.frCfgFile = QtWidgets.QFrame(self.frame_8)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frCfgFile.sizePolicy().hasHeightForWidth())
        self.frCfgFile.setSizePolicy(sizePolicy)
        self.frCfgFile.setMinimumSize(QtCore.QSize(400, 0))
        self.frCfgFile.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frCfgFile.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frCfgFile.setObjectName("frCfgFile")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.frCfgFile)
        self.verticalLayout_8.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.cmbCfgFile = QtWidgets.QComboBox(self.frCfgFile)
        self.cmbCfgFile.setObjectName("cmbCfgFile")
        self.verticalLayout_8.addWidget(self.cmbCfgFile)
        self.tableWidget = QtWidgets.QTableWidget(self.frCfgFile)
        self.tableWidget.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.verticalLayout_8.addWidget(self.tableWidget)
        self.horizontalLayout_10.addWidget(self.frCfgFile)
        self.verticalLayout_6.addWidget(self.frame_8)
        self.twAllExcels = QtWidgets.QTableWidget(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.twAllExcels.sizePolicy().hasHeightForWidth())
        self.twAllExcels.setSizePolicy(sizePolicy)
        self.twAllExcels.setObjectName("twAllExcels")
        self.twAllExcels.setColumnCount(0)
        self.twAllExcels.setRowCount(0)
        self.verticalLayout_6.addWidget(self.twAllExcels)
        self.twParsingResult = QtWidgets.QTableWidget(Form)
        self.twParsingResult.setMinimumSize(QtCore.QSize(0, 124))
        self.twParsingResult.setMaximumSize(QtCore.QSize(16777215, 124))
        self.twParsingResult.setObjectName("twParsingResult")
        self.twParsingResult.setColumnCount(0)
        self.twParsingResult.setRowCount(0)
        self.verticalLayout_6.addWidget(self.twParsingResult)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.leDir.setText(_translate("Form", "/home/da3/Загрузки/перенос/"))
        self.pbFond.setText(_translate("Form", "Фонд"))
        self.pbAgent.setText(_translate("Form", "Агент"))
        self.pbSigner.setText(_translate("Form", "Подписант"))
        self.chbClientOnly.setText(_translate("Form", "перенести только клиента"))
        self.chbSocium.setText(_translate("Form", "Сбросить номер Социума"))
        self.chbSuff.setText(_translate("Form", "Суффикс"))
        self.leSuff.setText(_translate("Form", "07-АА00-208"))
        self.chbOurStat.setText(_translate("Form", "Сбросить внутренние статусы"))
        self.chbFondStat.setText(_translate("Form", "Сбросить статусы Фонда"))
        self.chbArhivON.setText(_translate("Form", "Поставить флаг \"Архивный\""))
        self.chbArhivOFF.setText(_translate("Form", "Сбросить флаг \"Архивный\""))
        self.chbDateFrom.setText(_translate("Form", "Выборка от"))
        self.chbDateTo.setText(_translate("Form", "Выборка до"))
        self.lbSQLcl.setText(_translate("Form", "client     "))
        self.lbSQLco.setText(_translate("Form", "contract"))
        self.pbRefreshImport.setText(_translate("Form", "Обновить"))
        self.pbImport.setText(_translate("Form", "Импорт"))
        self.pbIdGenerate4PasportCheck.setText(_translate("Form", "Генерация"))
        self.pbPasportCheck.setText(_translate("Form", "Проверить"))
        self.pbRefresh.setText(_translate("Form", "Обновить"))
        self.pbMove.setText(_translate("Form", "Перенос"))
        self.progressBar.setFormat(_translate("Form", "%v из %m"))

