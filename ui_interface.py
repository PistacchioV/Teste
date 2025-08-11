# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'new_interface.ui'
##
## Created by: Qt User Interface Compiler version 6.6.3
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QComboBox, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QProgressBar,
    QPushButton, QSizePolicy, QSpacerItem, QVBoxLayout,
    QWidget, QGridLayout)

from Custom_Widgets.QCustomQStackedWidget import QCustomQStackedWidget
from Custom_Widgets.QCustomSlideMenu import QCustomSlideMenu

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
                MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(1009, 619)
        font = QFont()
        font.setPointSize(10)
        MainWindow.setFont(font)
        MainWindow.setStyleSheet(u"")

        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.centralwidget.setMinimumSize(QSize(971, 619))
        self.horizontalLayout = QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setSpacing(0)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.horizontalLayout.setContentsMargins(10, 10, 10, 10)

        # ========= CRIE OS WIDGETS ANTES DE ADICIONAR AO LAYOUT =========
        self.leftMenu = QCustomSlideMenu(self.centralwidget)
        self.leftMenu.setObjectName(u"leftMenu")
        self.centerMenu = QCustomSlideMenu(self.centralwidget)
        self.centerMenu.setObjectName(u"centerMenu")
        self.mainBody = QWidget(self.centralwidget)
        self.mainBody.setObjectName(u"mainBody")
        # Adicione ao layout na ordem correta para ficar lado a lado
        self.horizontalLayout.addWidget(self.leftMenu, 0, Qt.AlignmentFlag.AlignLeft)
        self.horizontalLayout.addWidget(self.centerMenu)
        self.horizontalLayout.addWidget(self.mainBody)
        # ========= leftMenu content (igual ao original) =========
        self.verticalLayout = QVBoxLayout(self.leftMenu)
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)

        # ------------ leftMenu Top (menuBtn) ---------------
        self.widget = QWidget(self.leftMenu)
        self.widget.setObjectName(u"widget")
        self.widget.setMinimumSize(QSize(46, 42))
        self.verticalLayout_2 = QVBoxLayout(self.widget)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.verticalLayout_2.setContentsMargins(5, 5, 0, 5)
        self.menuBtn = QPushButton(self.widget)
        self.menuBtn.setObjectName(u"menuBtn")
        icon = QIcon()
        icon.addFile(u":/material_design/icons/material_design/menu.png", QSize(), QIcon.Normal, QIcon.Off)
        self.menuBtn.setIcon(icon)
        self.verticalLayout_2.addWidget(self.menuBtn)
        self.verticalLayout.addWidget(self.widget, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        # ------------ leftMenu Main Buttons ---------------
        self.widget_3 = QWidget(self.leftMenu)
        self.widget_3.setObjectName(u"widget_3")
        self.widget_3.setMinimumSize(QSize(123, 136))
        self.verticalLayout_4 = QVBoxLayout(self.widget_3)
        self.verticalLayout_4.setSpacing(5)
        self.verticalLayout_4.setObjectName(u"verticalLayout_4")
        self.verticalLayout_4.setContentsMargins(5, 5, 0, 5)

        # ----------- INÍCIO Página Home -----------
        self.homeBtn = QPushButton(self.widget_3)
        self.homeBtn.setObjectName(u"homeBtn")
        icon1 = QIcon()
        icon1.addFile(u":/feather/icons/feather/home.png", QSize(), QIcon.Normal, QIcon.Off)
        self.homeBtn.setIcon(icon1)
        self.verticalLayout_4.addWidget(self.homeBtn)

        # ----------- INÍCIO Página NDF -----------
        self.ndfBtn = QPushButton(self.widget_3)
        self.ndfBtn.setObjectName(u"ndfBtn")
        icon2 = QIcon()
        icon2.addFile(u":/font_awesome/brands/nfc-symbol.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfBtn.setIcon(icon2)
        self.verticalLayout_4.addWidget(self.ndfBtn)
        
        # ----------- INÍCIO Página Option -----------
        self.optBtn = QPushButton(self.widget_3)
        self.optBtn.setObjectName(u"optBtn")
        icon3 = QIcon()
        icon3.addFile(u":icons/font_awesome/brands/opera.png", QSize(), QIcon.Normal, QIcon.Off)
        self.optBtn.setIcon(icon3)
        self.verticalLayout_4.addWidget(self.optBtn)

        # ----------- INÍCIO Página Swap -----------
        self.swaptBtn = QPushButton(self.widget_3)
        self.swaptBtn.setObjectName(u"swaptBtn")
        icon4 = QIcon()
        icon4.addFile("Qss/icons/000080/font_awesome/brands/stripe-s.png", QSize(), QIcon.Normal, QIcon.Off)
        self.swaptBtn.setIcon(icon4)
        self.verticalLayout_4.addWidget(self.swaptBtn)

        # ----------- INÍCIO Página counterparty -----------
        self.counterpartyBtn = QPushButton(self.widget_3)
        self.counterpartyBtn.setObjectName(u"counterpartyBtn")
        icon5 = QIcon()
        icon5.addFile("Qss/icons/000080/font_awesome/brands/creative-commons-by.png", QSize(), QIcon.Normal, QIcon.Off)
        self.counterpartyBtn.setIcon(icon5)
        self.verticalLayout_4.addWidget(self.counterpartyBtn)

        # ----------- INÍCIO Página Commodities -----------
        self.commodiBtn = QPushButton(self.widget_3)
        self.commodiBtn.setObjectName(u"commodiBtn")
        icon6 = QIcon()
        icon6.addFile("Qss/icons/000080/font_awesome/brands/audible.png", QSize(), QIcon.Normal, QIcon.Off)
        self.commodiBtn.setIcon(icon6)
        self.verticalLayout_4.addWidget(self.commodiBtn)

        # ----------- INÍCIO Página Metrics -----------
        self.metricsBtn = QPushButton(self.widget_3)
        self.metricsBtn.setObjectName(u"metricsBtn")
        icon_metrics = QIcon()
        icon_metrics.addFile("Qss/icons/000080/feather/trending-up.png", QSize(), QIcon.Normal, QIcon.Off)
        self.metricsBtn.setIcon(icon_metrics)
        self.metricsBtn.setText(QCoreApplication.translate("MainWindow", u"Metrics", None))
        self.verticalLayout_4.addWidget(self.metricsBtn)

        # ----------- INÍCIO Página Confirmations -----------
        self.confirmationsBtn = QPushButton(self.widget_3)
        self.confirmationsBtn.setObjectName(u"confirmationsBtn")
        icon_confirm = QIcon()
        icon_confirm.addFile("Qss/icons/000080/font_awesome/regular/file-lines.png", QSize(), QIcon.Normal, QIcon.Off)
        self.confirmationsBtn.setIcon(icon_confirm)
        self.confirmationsBtn.setText(QCoreApplication.translate("MainWindow", u"Confirmations", None))
        self.verticalLayout_4.addWidget(self.confirmationsBtn)

        # ----------- INÍCIO Página Intrag -----------
        self.intragBtn = QPushButton(self.widget_3)
        self.intragBtn.setObjectName(u"intragBtn")
        iconIntrag = QIcon()
        iconIntrag.addFile("Qss/icons/000080/font_awesome/brands/itau.png", QSize(), QIcon.Normal, QIcon.Off)
        self.intragBtn.setIcon(iconIntrag)
        self.verticalLayout_4.addWidget(self.intragBtn)

        self.verticalLayout.addWidget(self.widget_3, 0, Qt.AlignmentFlag.AlignTop)
        self.verticalSpacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.verticalLayout.addItem(self.verticalSpacer)

        # ------------ leftMenu Bottom (settings/info/help) ---------------
        self.widget_2 = QWidget(self.leftMenu)
        self.widget_2.setObjectName(u"widget_2")
        self.widget_2.setMinimumSize(QSize(113, 105))
        self.verticalLayout_3 = QVBoxLayout(self.widget_2)
        self.verticalLayout_3.setSpacing(5)
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.verticalLayout_3.setContentsMargins(5, 5, 0, 5)
        self.settingsBtn = QPushButton(self.widget_2)
        self.settingsBtn.setObjectName(u"settingsBtn")
        icon7 = QIcon()
        icon7.addFile(u":/feather/icons/feather/settings.png", QSize(), QIcon.Normal, QIcon.Off)
        self.settingsBtn.setIcon(icon7)
        self.verticalLayout_3.addWidget(self.settingsBtn)
        self.infoBtn = QPushButton(self.widget_2)
        self.infoBtn.setObjectName(u"infoBtn")
        icon8 = QIcon()
        icon8.addFile(u":/feather/icons/feather/info.png", QSize(), QIcon.Normal, QIcon.Off)
        self.infoBtn.setIcon(icon8)
        self.verticalLayout_3.addWidget(self.infoBtn)
        self.helpBtn = QPushButton(self.widget_2)
        self.helpBtn.setObjectName(u"helpBtn")
        icon9 = QIcon()
        icon9.addFile(u":/feather/icons/feather/help-circle.png", QSize(), QIcon.Normal, QIcon.Off)
        self.helpBtn.setIcon(icon9)
        self.verticalLayout_3.addWidget(self.helpBtn)
        self.verticalLayout.addWidget(self.widget_2, 0, Qt.AlignmentFlag.AlignBottom)

       # ========= centerMenu content (ajustado) =========
        self.verticalLayout_5 = QVBoxLayout(self.centerMenu)
        self.verticalLayout_5.setObjectName(u"verticalLayout_5")
        self.verticalLayout_5.setContentsMargins(6, 6, 6, 6)        

        # Barra de botões de topo do centerMenu
        self.widget_4 = QWidget(self.centerMenu)        
        self.widget_4.setObjectName(u"widget_4")

        # Crie os botões e label normalmente
        self.returnCenterMenuBtn = QPushButton(self.widget_4)
        self.returnCenterMenuBtn.setObjectName(u"returnCenterMenuBtn")
        icon_return = QIcon()
        icon_return.addFile(u":/feather/icons/feather/rotate-ccw.png", QSize(), QIcon.Normal, QIcon.Off)
        self.returnCenterMenuBtn.setIcon(icon_return)
        self.returnCenterMenuBtn.setVisible(False)

        self.closeCenterMenuBtn = QPushButton(self.widget_4)
        self.closeCenterMenuBtn.setObjectName(u"closeCenterMenuBtn")
        icon_close = QIcon()
        icon_close.addFile(u":/feather/icons/feather/x-circle.png", QSize(), QIcon.Normal, QIcon.Off)
        self.closeCenterMenuBtn.setIcon(icon_close)

        self.label = QLabel(self.widget_4)
        self.label.setObjectName(u"label")

        # --------- INSIRA AQUI O BLOCO DO LAYOUT ---------
        self.verticalLayout_header = QVBoxLayout(self.widget_4)
        self.verticalLayout_header.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_btns = QHBoxLayout()
        self.horizontalLayout_btns.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_btns.addWidget(self.returnCenterMenuBtn)
        self.horizontalLayout_btns.addWidget(self.closeCenterMenuBtn)
        self.verticalLayout_header.addLayout(self.horizontalLayout_btns)
        self.verticalLayout_header.addWidget(self.label)
        # -------------------------------------------------

        self.verticalLayout_5.addWidget(self.widget_4)


        self.centerMenuPages = QCustomQStackedWidget(self.centerMenu)
        self.centerMenuPages.setObjectName(u"centerMenuPages")


        # ----- Página Settings -----
        self.settingsPage = QWidget()
        self.settingsPage.setObjectName(u"settingsPage")
        self.verticalLayout_6 = QVBoxLayout(self.settingsPage)
        self.verticalLayout_6.setObjectName(u"verticalLayout_6")
        self.verticalSpacer_2 = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.verticalLayout_6.addItem(self.verticalSpacer_2)
        self.widget_5 = QWidget(self.settingsPage)
        self.widget_5.setObjectName(u"widget_5")
        self.verticalLayout_7 = QVBoxLayout(self.widget_5)
        self.verticalLayout_7.setObjectName(u"verticalLayout_7")
        self.label_2 = QLabel(self.widget_5)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_7.addWidget(self.label_2)
        self.frame = QFrame(self.widget_5)
        self.frame.setObjectName(u"frame")
        self.frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.frame.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_3 = QHBoxLayout(self.frame)
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.label_3 = QLabel(self.frame)
        self.label_3.setObjectName(u"label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.themeList = QComboBox(self.frame)

        self.themeList.setObjectName(u"themeList")
        self.horizontalLayout_3.addWidget(self.themeList)
        self.verticalLayout_7.addWidget(self.frame)
        self.verticalLayout_6.addWidget(self.widget_5, 0, Qt.AlignmentFlag.AlignVCenter)
        self.verticalSpacer_3 = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.verticalLayout_6.addItem(self.verticalSpacer_3)
        self.centerMenuPages.addWidget(self.settingsPage)

        # ----- Página Information -----
        self.informationPage = QWidget()
        self.informationPage.setObjectName(u"informationPage")
        self.verticalLayout_8 = QVBoxLayout(self.informationPage)
        self.verticalLayout_8.setObjectName(u"verticalLayout_8")
        self.label_4 = QLabel(self.informationPage)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_8.addWidget(self.label_4, 0, Qt.AlignmentFlag.AlignVCenter)
        self.centerMenuPages.addWidget(self.informationPage)

        # ----- Página Help -----
        self.helpPage = QWidget()
        self.helpPage.setObjectName(u"helpPage")
        self.verticalLayout_9 = QVBoxLayout(self.helpPage)
        self.verticalLayout_9.setObjectName(u"verticalLayout_9")
        self.label_5 = QLabel(self.helpPage)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_9.addWidget(self.label_5, 0, Qt.AlignmentFlag.AlignVCenter)
        self.centerMenuPages.addWidget(self.helpPage)

        # ----- Sub Página NDF (NDF) -----
        self.ndfSubPage = QWidget()
        self.ndfSubPage.setObjectName(u"ndfSubPage")
        self.verticalLayout_ndf = QVBoxLayout(self.ndfSubPage)
        self.verticalLayout_ndf.setObjectName(u"verticalLayout_ndf")
        self.ndfSearchBtn = QPushButton(self.ndfSubPage)        
        self.ndfSearchBtn.setObjectName(u"ndfSearchBtn")
        self.ndfSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        icon_search = QIcon()       
        icon_search.addFile("Qss/icons/000080/feather/search.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfSearchBtn.setIcon(icon_search)
        self.verticalLayout_ndf.addWidget(self.ndfSearchBtn)
        self.ndfRegisterBtn = QPushButton(self.ndfSubPage)
        self.ndfRegisterBtn.setObjectName(u"ndfRegisterBtn")
        self.ndfRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register ", None))
        icon_reg = QIcon()
        icon_reg.addFile("Qss/icons/000080/feather/save.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfRegisterBtn.setIcon(icon_reg)
        self.verticalLayout_ndf.addWidget(self.ndfRegisterBtn)
        self.ndfSettleBtn = QPushButton(self.ndfSubPage)
        self.ndfSettleBtn.setObjectName(u"ndfSettleBtn")
        self.ndfSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement ", None))
        icon_settle = QIcon()        
        icon_settle.addFile(u"Qss/icons/000080/font_awesome/regular/money-bill-1.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfSettleBtn.setIcon(icon_settle)
        self.verticalLayout_ndf.addWidget(self.ndfSettleBtn)        
        self.label_ndf = QLabel(self.ndfSubPage)
        self.label_ndf.setObjectName(u"label_ndf")
        self.label_ndf.setText("Subpágina NDF")
        self.label_ndf.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_ndf.addWidget(self.label_ndf)
        self.centerMenuPages.addWidget(self.ndfSubPage)

         # ----- Sub Sub Página NDF -----
        self.ndfSubSubPage = QWidget()
        self.ndfSubSubPage.setObjectName(u"ndfSubSubPage")
        self.verticalLayout_ndfSubSub = QVBoxLayout(self.ndfSubSubPage)
        self.verticalLayout_ndfSubSub.setObjectName(u"verticalLayout_ndfSubSub")
        self.ndfNewBtn = QPushButton(self.ndfSubSubPage)
        self.ndfNewBtn.setObjectName(u"ndfNewBtn")
        self.ndfNewBtn.setText(QCoreApplication.translate("MainWindow", u"New Deals", None))
        icon_new = QIcon()
        icon_new.addFile(u":/feather/icons/feather/plus.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfNewBtn.setIcon(icon_new)  # Reuse icon_new or create another if desired
        self.verticalLayout_ndfSubSub.addWidget(self.ndfNewBtn)
        self.ndfUnwindBtn = QPushButton(self.ndfSubSubPage)
        self.ndfUnwindBtn.setObjectName(u"ndfUnwindBtn")
        self.ndfUnwindBtn.setText(QCoreApplication.translate("MainWindow", u"Unwinds", None))
        icon_unwind = QIcon()
        icon_unwind.addFile(u":/feather/icons/feather/repeat.png", QSize(), QIcon.Normal, QIcon.Off)
        self.ndfUnwindBtn.setIcon(icon_unwind)  # Reuse icon_unwind
        self.verticalLayout_ndfSubSub.addWidget(self.ndfUnwindBtn)
        self.label_ndfSubSub = QLabel(self.ndfSubSubPage)
        self.label_ndfSubSub.setObjectName(u"label_ndfSubSub")
        self.label_ndfSubSub.setText("SubSubpágina NDF")
        self.label_ndfSubSub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_ndfSubSub.addWidget(self.label_ndfSubSub)
        self.centerMenuPages.addWidget(self.ndfSubSubPage)

        # ----- Sub Página Option -----
        self.opcaoSubPage = QWidget()
        self.opcaoSubPage.setObjectName(u"opcaoSubPage")
        self.verticalLayout_opcao = QVBoxLayout(self.opcaoSubPage)
        self.verticalLayout_opcao.setObjectName(u"verticalLayout_opcao")        
        self.optSearchBtn = QPushButton(self.opcaoSubPage)       
        self.optSearchBtn.setObjectName(u"optSearchBtn")
        self.optSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.optSearchBtn.setIcon(icon_search)
        self.verticalLayout_opcao.addWidget(self.optSearchBtn)
        self.optRegisterBtn = QPushButton(self.opcaoSubPage)
        self.optRegisterBtn.setObjectName(u"optRegisterBtn")
        self.optRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register ", None))
        self.optRegisterBtn.setIcon(icon_reg)
        self.verticalLayout_opcao.addWidget(self.optRegisterBtn)
        self.optSettleBtn = QPushButton(self.opcaoSubPage)
        self.optSettleBtn.setObjectName(u"optSettleBtn")
        self.optSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement ", None))
        self.optSettleBtn.setIcon(icon_settle)
        self.verticalLayout_opcao.addWidget(self.optSettleBtn)         
        self.label_opcao = QLabel(self.opcaoSubPage)
        self.label_opcao.setObjectName(u"label_opcao")
        self.label_opcao.setText("Subpágina Option")
        self.label_opcao.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_opcao.addWidget(self.label_opcao)
        self.centerMenuPages.addWidget(self.opcaoSubPage)

         # ----- Sub Sub Página Option -----
        self.opcaoSubSubPage = QWidget()
        self.opcaoSubSubPage.setObjectName(u"opcaoSubSubPage")
        self.verticalLayout_opcaoSubSub = QVBoxLayout(self.opcaoSubSubPage)
        self.verticalLayout_opcaoSubSub.setObjectName(u"verticalLayout_opcaoSubSub")
        self.optNewBtn = QPushButton(self.opcaoSubSubPage)
        self.optNewBtn.setObjectName(u"optNewBtn")
        self.optNewBtn.setText(QCoreApplication.translate("MainWindow", u"New Deals", None))        
        self.optNewBtn.setIcon(icon_new)
        self.verticalLayout_opcaoSubSub.addWidget(self.optNewBtn)
        self.optUnwindBtn = QPushButton(self.opcaoSubSubPage)
        self.optUnwindBtn.setObjectName(u"optUnwindBtn")
        self.optUnwindBtn.setText(QCoreApplication.translate("MainWindow", u"Unwinds", None))        
        self.optUnwindBtn.setIcon(icon_unwind)
        self.verticalLayout_opcaoSubSub.addWidget(self.optUnwindBtn)
        self.label_opcaoSubSub = QLabel(self.opcaoSubSubPage)
        self.label_opcaoSubSub.setObjectName(u"label_opcaoSubSub")
        self.label_opcaoSubSub.setText("SubSubpágina Option")
        self.label_opcaoSubSub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_opcaoSubSub.addWidget(self.label_opcaoSubSub)
        self.centerMenuPages.addWidget(self.opcaoSubSubPage)
       
        # ----- Sub Página Swap -----
        self.swapSubPage = QWidget()
        self.swapSubPage.setObjectName(u"swapSubPage")
        self.verticalLayout_swap = QVBoxLayout(self.swapSubPage)
        self.verticalLayout_swap.setObjectName(u"verticalLayout_swap")        
        self.swapSearchBtn = QPushButton(self.swapSubPage)      
        self.swapSearchBtn.setObjectName(u"swapSearchBtn")
        self.swapSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.swapSearchBtn.setIcon(icon_search)
        self.verticalLayout_swap.addWidget(self.swapSearchBtn)
        self.swapRegisterBtn = QPushButton(self.swapSubPage)
        self.swapRegisterBtn.setObjectName(u"swapRegisterBtn")
        self.swapRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register ", None))
        self.swapRegisterBtn.setIcon(icon_reg)
        self.verticalLayout_swap.addWidget(self.swapRegisterBtn)
        self.swapSettleBtn = QPushButton(self.swapSubPage)
        self.swapSettleBtn.setObjectName(u"swapSettleBtn")
        self.swapSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement ", None))
        self.swapSettleBtn.setIcon(icon_settle)
        self.verticalLayout_swap.addWidget(self.swapSettleBtn)          
        self.label_swap = QLabel(self.swapSubPage)
        self.label_swap.setObjectName(u"label_swap")
        self.label_swap.setText("Subpágina Swap")
        self.label_swap.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_swap.addWidget(self.label_swap)
        self.centerMenuPages.addWidget(self.swapSubPage)

         #----- Sub Sub Página Swap -----
        self.swapSubSubPage = QWidget()
        self.swapSubSubPage.setObjectName(u"swapSubSubPage")
        self.verticalLayout_swapSubSub = QVBoxLayout(self.swapSubSubPage)
        self.verticalLayout_swapSubSub.setObjectName(u"verticalLayout_swapSubSub")
        self.swapNewBtn = QPushButton(self.swapSubSubPage)
        self.swapNewBtn.setObjectName(u"swapNewBtn")
        self.swapNewBtn.setText(QCoreApplication.translate("MainWindow", u"New Deals", None))
        self.swapNewBtn.setIcon(icon_new)
        self.verticalLayout_swapSubSub.addWidget(self.swapNewBtn)
        self.swapUnwindBtn = QPushButton(self.swapSubSubPage)
        self.swapUnwindBtn.setObjectName(u"swapUnwindBtn")
        self.swapUnwindBtn.setText(QCoreApplication.translate("MainWindow", u"Unwinds", None))
        self.swapUnwindBtn.setIcon(icon_unwind)
        self.verticalLayout_swapSubSub.addWidget(self.swapUnwindBtn)
        self.label_swapSubSub = QLabel(self.swapSubSubPage)
        self.label_swapSubSub.setObjectName(u"label_swapSubSub")
        self.label_swapSubSub.setText("SubSubpágina Swap")
        self.label_swapSubSub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_swapSubSub.addWidget(self.label_swapSubSub)
        self.centerMenuPages.addWidget(self.swapSubSubPage)

        # ----- Sub Página Comitente -----
        self.counterpartySubPage = QWidget()
        self.counterpartySubPage.setObjectName(u"counterpartySubPage")
        self.verticalLayout_counterparty = QVBoxLayout(self.counterpartySubPage)
        self.verticalLayout_counterparty.setObjectName(u"verticalLayout_counterparty")
        self.counterpartyRegisterBtn = QPushButton(self.counterpartySubPage)
        self.counterpartyRegisterBtn.setObjectName(u"counterpartyRegisterBtn")
        self.counterpartyRegisterBtn.setText("Cadastro")
        self.counterpartyRegisterBtn.setIcon(icon_reg)
        self.verticalLayout_counterparty.addWidget(self.counterpartyRegisterBtn)
        self.counterpartySearchBtn = QPushButton(self.counterpartySubPage)
        self.counterpartySearchBtn.setObjectName(u"counterpartySearchBtn")
        self.counterpartySearchBtn.setText("Search")
        self.counterpartySearchBtn.setIcon(icon_search)
        self.verticalLayout_counterparty.addWidget(self.counterpartySearchBtn)
        self.label_counterparty = QLabel(self.counterpartySubPage)
        self.label_counterparty.setObjectName(u"label_counterparty")
        self.label_counterparty.setText("Subpágina Counterparty")
        self.label_counterparty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_counterparty.addWidget(self.label_counterparty)
        self.centerMenuPages.addWidget(self.counterpartySubPage)

        # ----- Sub Página Commodities -----
        self.commoditiesSubPage = QWidget()
        self.commoditiesSubPage.setObjectName(u"commoditiesSubPage")
        self.verticalLayout_commodities = QVBoxLayout(self.commoditiesSubPage)
        self.verticalLayout_commodities.setObjectName(u"verticalLayout_commodities")
        self.commoditiesRegisterBtn = QPushButton(self.commoditiesSubPage)
        self.commoditiesRegisterBtn.setObjectName(u"commoditiesRegisterBtn")
        self.commoditiesRegisterBtn.setText("Cadastro")
        self.commoditiesRegisterBtn.setIcon(icon_reg)
        self.verticalLayout_commodities.addWidget(self.commoditiesRegisterBtn)
        self.commoditiesSearchBtn = QPushButton(self.commoditiesSubPage)
        self.commoditiesSearchBtn.setObjectName(u"commoditiesSearchBtn")
        self.commoditiesSearchBtn.setText("Search")
        self.commoditiesSearchBtn.setIcon(icon_search)
        self.verticalLayout_commodities.addWidget(self.commoditiesSearchBtn)
        self.label_commodities = QLabel(self.commoditiesSubPage)
        self.label_commodities.setObjectName(u"label_commodities")
        self.label_commodities.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_commodities.addWidget(self.label_commodities)
        self.centerMenuPages.addWidget(self.commoditiesSubPage)

        # ----- Sub Página Metrics -----
        self.metricsSubPage = QWidget()
        self.metricsSubPage.setObjectName(u"metricsSubPage")
        self.verticalLayout_metrics = QVBoxLayout(self.metricsSubPage)
        self.verticalLayout_metrics.setObjectName(u"verticalLayout_metrics")
        self.kpiBtn = QPushButton(self.metricsSubPage)
        self.kpiBtn.setObjectName(u"kpiBtn")
        self.kpiBtn.setText(QCoreApplication.translate("MainWindow", u"KPI", None))
        icon_kpi = QIcon()
        icon_kpi.addFile("Qss/icons/000080/feather/target.png", QSize(), QIcon.Normal, QIcon.Off)
        self.kpiBtn.setIcon(icon_kpi)
        self.verticalLayout_metrics.addWidget(self.kpiBtn)
        self.adhocBtn = QPushButton(self.metricsSubPage)
        self.adhocBtn.setObjectName(u"adhocBtn")
        self.adhocBtn.setText(QCoreApplication.translate("MainWindow", u"Ad-Hoc", None))
        icon_adhoc = QIcon()
        icon_adhoc.addFile("Qss/icons/000080/feather/message-square.png", QSize(), QIcon.Normal, QIcon.Off)
        self.adhocBtn.setIcon(icon_adhoc)
        self.verticalLayout_metrics.addWidget(self.adhocBtn)
        self.label_metrics = QLabel(self.metricsSubPage)
        self.label_metrics.setObjectName(u"label_metrics")
        self.label_metrics.setText("Subpágina Metrics")
        self.label_metrics.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_metrics.addWidget(self.label_metrics)
        self.centerMenuPages.addWidget(self.metricsSubPage)

        # ----- Sub Página Confirmations -----        
    
        self.confirmationsSubPage = QWidget()
        self.confirmationsSubPage.setObjectName(u"confirmationsSubPage")
        self.verticalLayout_conf = QVBoxLayout(self.confirmationsSubPage)
        self.verticalLayout_conf.setObjectName(u"verticalLayout_conf")
        self.NDFConfirmationBtn = QPushButton(self.confirmationsSubPage)
        self.NDFConfirmationBtn.setObjectName(u"NDFConfirmationBtn")
        self.NDFConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"NDF", None))        
        self.NDFConfirmationBtn.setIcon(icon2)
        self.verticalLayout_conf.addWidget(self.NDFConfirmationBtn)
        self.optionConfirmationBtn = QPushButton(self.confirmationsSubPage)
        self.optionConfirmationBtn.setObjectName(u"optionConfirmationBtn")
        self.optionConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"Option", None))       
        self.optionConfirmationBtn.setIcon(icon3)
        self.verticalLayout_conf.addWidget(self.optionConfirmationBtn)
        self.swapConfirmationBtn = QPushButton(self.confirmationsSubPage)
        self.swapConfirmationBtn.setObjectName(u"swapConfirmationBtn")
        self.swapConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"Swap", None))     
        self.swapConfirmationBtn.setIcon(icon4)
        self.verticalLayout_conf.addWidget(self.swapConfirmationBtn)
        self.label_confirmations = QLabel(self.confirmationsSubPage)
        self.label_confirmations.setObjectName(u"label_confirmations")
        self.label_confirmations.setText("Subpágina Confirmations")
        self.label_confirmations.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_conf.addWidget(self.label_confirmations)
        self.centerMenuPages.addWidget(self.confirmationsSubPage)


      

       

        # ----- Sub Página Intrag -----
        self.intragSubPage = QWidget()
        self.intragSubPage.setObjectName(u"intragSubPage")
        self.verticalLayout_intrag = QVBoxLayout(self.intragSubPage)
        self.verticalLayout_intrag.setObjectName(u"verticalLayout_intrag")

        self.intragSearchBtn = QPushButton(self.intragSubPage)
        self.intragSearchBtn.setObjectName(u"intragSearchBtn")
        self.intragSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.verticalLayout_intrag.addWidget(self.intragSearchBtn)

        self.intragBoletaBtn = QPushButton(self.intragSubPage)
        self.intragBoletaBtn.setObjectName(u"intragBoletaBtn")
        self.intragBoletaBtn.setText(QCoreApplication.translate("MainWindow", u"Instruction", None))
        self.verticalLayout_intrag.addWidget(self.intragBoletaBtn)
        icon_boleta= QIcon()       
        icon_boleta.addFile("Qss/icons/Icons/feather/file-plus.png", QSize(), QIcon.Normal, QIcon.Off)
        self.intragBoletaBtn.setIcon(icon_boleta)
        self.label_intrag = QLabel(self.intragSubPage)
        self.label_intrag.setObjectName(u"label_intrag")
        self.label_intrag.setText("Subpágina Intrag")
        self.label_intrag.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_intrag.addWidget(self.label_intrag)

        self.centerMenuPages.addWidget(self.intragSubPage)

        self.verticalLayout_5.addWidget(self.centerMenuPages)      

       # ================= CORPO PRINCIPAL ==================
        self.mainBody = QWidget(self.centralwidget)
        self.mainBody.setObjectName(u"mainBody")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.mainBody.sizePolicy().hasHeightForWidth())
        self.mainBody.setSizePolicy(sizePolicy)
        self.verticalLayout_10 = QVBoxLayout(self.mainBody)
        self.verticalLayout_10.setSpacing(0)
        self.verticalLayout_10.setObjectName(u"verticalLayout_10")
        self.verticalLayout_10.setContentsMargins(0, 0, 0, 0)

        # ---------- Header ----------
        self.header = QWidget(self.mainBody)
        self.header.setObjectName(u"header")
        self.horizontalLayout_7 = QHBoxLayout(self.header)
        self.horizontalLayout_7.setSpacing(5)
        self.horizontalLayout_7.setObjectName(u"horizontalLayout_7")
        self.horizontalLayout_7.setContentsMargins(5, 0, 0, 5)
        self.titleTxt = QLabel(self.header)
        self.titleTxt.setObjectName(u"titleTxt")
        font1 = QFont()
        font1.setPointSize(13)
        font1.setBold(True)
        self.titleTxt.setFont(font1)
        self.horizontalLayout_7.addWidget(self.titleTxt, 0, Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignBottom)
        self.frame_3 = QFrame(self.header)
        self.frame_3.setObjectName(u"frame_3")
        self.frame_3.setFrameShape(QFrame.Shape.StyledPanel)
        self.frame_3.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_6 = QHBoxLayout(self.frame_3)
        self.horizontalLayout_6.setSpacing(5)
        self.horizontalLayout_6.setObjectName(u"horizontalLayout_6")
        self.horizontalLayout_6.setContentsMargins(5, 5, -1, 5)
        self.notificationBtn = QPushButton(self.frame_3)
        self.notificationBtn.setObjectName(u"notificationBtn")
        icon11 = QIcon()
        icon11.addFile(u":/feather/icons/feather/bell.png", QSize(), QIcon.Normal, QIcon.Off)
        self.notificationBtn.setIcon(icon11)
        self.horizontalLayout_6.addWidget(self.notificationBtn)
        self.moreBtn = QPushButton(self.frame_3)
        self.moreBtn.setObjectName(u"moreBtn")
        icon12 = QIcon()
        icon12.addFile(u":/feather/icons/feather/more-horizontal.png", QSize(), QIcon.Normal, QIcon.Off)
        self.moreBtn.setIcon(icon12)
        self.horizontalLayout_6.addWidget(self.moreBtn)
        self.profileBtn = QPushButton(self.frame_3)
        self.profileBtn.setObjectName(u"profileBtn")
        icon13 = QIcon()
        icon13.addFile(u":/feather/icons/feather/user.png", QSize(), QIcon.Normal, QIcon.Off)
        self.profileBtn.setIcon(icon13)
        self.horizontalLayout_6.addWidget(self.profileBtn)
        self.horizontalLayout_7.addWidget(self.frame_3, 0, Qt.AlignmentFlag.AlignHCenter|Qt.AlignmentFlag.AlignBottom)
        self.searchInpCont = QFrame(self.header)
        self.searchInpCont.setObjectName(u"searchInpCont")
        self.searchInpCont.setMaximumSize(QSize(261, 16777215))
        self.searchInpCont.setFrameShape(QFrame.Shape.StyledPanel)
        self.searchInpCont.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_8 = QHBoxLayout(self.searchInpCont)
        self.horizontalLayout_8.setSpacing(0)
        self.horizontalLayout_8.setObjectName(u"horizontalLayout_8")
        self.horizontalLayout_8.setContentsMargins(5, 5, 5, 5)
        self.label_9 = QLabel(self.searchInpCont)
        self.label_9.setObjectName(u"label_9")
        self.label_9.setMinimumSize(QSize(16, 16))
        self.label_9.setMaximumSize(QSize(16, 16))
        self.label_9.setPixmap(QPixmap(u":/material_design/icons/material_design/search.png"))
        self.label_9.setScaledContents(True)
        self.horizontalLayout_8.addWidget(self.label_9)
        self.searchInp = QLineEdit(self.searchInpCont)
        self.searchInp.setObjectName(u"searchInp")
        self.searchInp.setMinimumSize(QSize(100, 0))
        self.horizontalLayout_8.addWidget(self.searchInp)
        self.searchBtn = QPushButton(self.searchInpCont)
        self.searchBtn.setObjectName(u"searchBtn")
        self.horizontalLayout_8.addWidget(self.searchBtn)
        self.horizontalLayout_7.addWidget(self.searchInpCont, 0, Qt.AlignmentFlag.AlignBottom)
        self.frame_4 = QFrame(self.header)
        self.frame_4.setObjectName(u"frame_4")
        self.frame_4.setFrameShape(QFrame.Shape.StyledPanel)
        self.frame_4.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_9 = QHBoxLayout(self.frame_4)
        self.horizontalLayout_9.setSpacing(2)
        self.horizontalLayout_9.setObjectName(u"horizontalLayout_9")
        self.horizontalLayout_9.setContentsMargins(0, 0, 0, 0)
        self.minimizeBtn = QPushButton(self.frame_4)
        self.minimizeBtn.setObjectName(u"minimizeBtn")
        icon14 = QIcon()
        icon14.addFile(u":/feather/icons/feather/window_minimize.png", QSize(), QIcon.Normal, QIcon.Off)
        self.minimizeBtn.setIcon(icon14)
        self.horizontalLayout_9.addWidget(self.minimizeBtn)
        self.restoreBtn = QPushButton(self.frame_4)
        self.restoreBtn.setObjectName(u"restoreBtn")
        icon15 = QIcon()
        icon15.addFile(u":/feather/icons/feather/square.png", QSize(), QIcon.Normal, QIcon.Off)
        self.restoreBtn.setIcon(icon15)
        self.horizontalLayout_9.addWidget(self.restoreBtn)
        self.closeBtn = QPushButton(self.frame_4)
        self.closeBtn.setObjectName(u"closeBtn")
        icon16 = QIcon()
        icon16.addFile(u":/feather/icons/feather/window_close.png", QSize(), QIcon.Normal, QIcon.Off)
        self.closeBtn.setIcon(icon16)
        self.horizontalLayout_9.addWidget(self.closeBtn)
        self.horizontalLayout_7.addWidget(self.frame_4, 0, Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignTop)
        self.verticalLayout_10.addWidget(self.header, 0, Qt.AlignmentFlag.AlignTop)

        # ---------- Conteúdo principal: área central ----------
        self.mainContents = QWidget(self.mainBody)
        self.mainContents.setObjectName(u"mainContents")
        sizePolicy1 = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.mainContents.sizePolicy().hasHeightForWidth())
        self.mainContents.setSizePolicy(sizePolicy1)
        self.horizontalLayout_10 = QHBoxLayout(self.mainContents)
        self.horizontalLayout_10.setSpacing(5)
        self.horizontalLayout_10.setObjectName(u"horizontalLayout_10")
        self.horizontalLayout_10.setContentsMargins(5, 0, 5, 0)

        self.mainPagesCont = QWidget(self.mainContents)
        self.mainPagesCont.setObjectName(u"mainPagesCont")
        self.verticalLayout_11 = QVBoxLayout(self.mainPagesCont)
        self.verticalLayout_11.setSpacing(5)
        self.verticalLayout_11.setObjectName(u"verticalLayout_11")
        self.verticalLayout_11.setContentsMargins(5, 5, 5, 5)
        self.mainPages = QCustomQStackedWidget(self.mainPagesCont)
        self.mainPages.setObjectName(u"mainPages")
       # -------------------------------------------------------------------------
        # Página de Pesquisa de Commodities - Painel central totalmente responsivo
        # -------------------------------------------------------------------------
        # Este widget representa a página de pesquisa de commodities, exibida no painel central do sistema.
        # Ele é responsável por exibir o painel de busca, garantindo que tudo fique responsivo
        # e se ajuste ao tamanho da janela principal, sem limitações de largura.
        self.commoditiesSearchPage = QWidget()
        self.commoditiesSearchPage.setObjectName(u"commoditiesSearchPage")
        sizePolicyComm = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        sizePolicyComm.setHorizontalStretch(0)
        sizePolicyComm.setVerticalStretch(0)
        sizePolicyComm.setHeightForWidth(self.commoditiesSearchPage.sizePolicy().hasHeightForWidth())
        self.commoditiesSearchPage.setSizePolicy(sizePolicyComm)
       

        # Página Home
        self.homePage = QWidget()
        self.homePage.setObjectName(u"homePage")
        self.verticalLayout_12 = QVBoxLayout(self.homePage)
        self.verticalLayout_12.setObjectName(u"verticalLayout_12")
        self.label_8 = QLabel(self.homePage)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_12.addWidget(self.label_8)
        self.mainPages.addWidget(self.homePage)
        # Página NDF
        self.ndfPage = QWidget()
        self.ndfPage.setObjectName(u"ndfPage")
        self.verticalLayout_13 = QVBoxLayout(self.ndfPage)
        self.verticalLayout_13.setObjectName(u"verticalLayout_13")
        self.label_10 = QLabel(self.ndfPage)
        self.label_10.setObjectName(u"label_10")
        self.label_10.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_13.addWidget(self.label_10)
        self.mainPages.addWidget(self.ndfPage)
        # Página Option
        self.optionPage = QWidget()
        self.optionPage.setObjectName(u"optionPage")
        self.horizontalLayout_11 = QHBoxLayout(self.optionPage)
        self.horizontalLayout_11.setObjectName(u"horizontalLayout_11")
        self.label_11 = QLabel(self.optionPage)
        self.label_11.setObjectName(u"label_11")
        self.label_11.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.horizontalLayout_11.addWidget(self.label_11)
        self.mainPages.addWidget(self.optionPage)
        # Página Swap
        self.swapPage = QWidget()
        self.swapPage.setObjectName(u"swapPage")
        self.verticalLayout_14 = QVBoxLayout(self.swapPage)
        self.verticalLayout_14.setObjectName(u"verticalLayout_14")
        self.label_12 = QLabel(self.swapPage)
        self.label_12.setObjectName(u"label_12")
        self.label_12.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_14.addWidget(self.label_12)
        self.mainPages.addWidget(self.swapPage)
        # Página counterpartys
        self.counterpartyPage = QWidget()
        self.counterpartyPage.setObjectName(u"counterpartyPage")
        self.verticalLayout_15 = QVBoxLayout(self.counterpartyPage)
        self.verticalLayout_15.setObjectName(u"verticalLayout_15")
        self.label_13 = QLabel(self.counterpartyPage)
        self.label_13.setObjectName(u"label_13")
        self.label_13.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_15.addWidget(self.label_13)
        self.mainPages.addWidget(self.counterpartyPage)
        # Página Commodities
        self.commoditiesPage = QWidget()
        self.commoditiesPage.setObjectName(u"commoditiesPage")
        self.verticalLayout_16 = QVBoxLayout(self.commoditiesPage)
        self.verticalLayout_16.setObjectName(u"verticalLayout_16")
        self.label_14 = QLabel(self.commoditiesPage)
        self.label_14.setObjectName(u"label_14")
        self.label_14.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_16.addWidget(self.label_14)
        self.mainPages.addWidget(self.commoditiesPage)
        # Página Metrics
        self.metricsPage = QWidget()
        self.metricsPage.setObjectName(u"metricsPage")
        self.verticalLayout_metricsPage = QVBoxLayout(self.metricsPage)
        self.verticalLayout_metricsPage.setObjectName(u"verticalLayout_metricsPage")
        self.label_metricsPage = QLabel(self.metricsPage)
        self.label_metricsPage.setObjectName(u"label_metricsPage")
        self.label_metricsPage.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_metricsPage.addWidget(self.label_metricsPage)
        self.mainPages.addWidget(self.metricsPage)
        # Página Confirmations
        self.confirmationsPage = QWidget()
        self.confirmationsPage.setObjectName(u"confirmationsPage")
        self.verticalLayout_confirmationsPage = QVBoxLayout(self.confirmationsPage)
        self.verticalLayout_confirmationsPage.setObjectName(u"verticalLayout_confirmationsPage")
        self.label_confirmationsPage = QLabel(self.confirmationsPage)
        self.label_confirmationsPage.setObjectName(u"label_confirmationsPage")
        self.label_confirmationsPage.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_confirmationsPage.addWidget(self.label_confirmationsPage)
        self.mainPages.addWidget(self.confirmationsPage)
        # Página Intrag
        self.intragPage = QWidget()
        self.intragPage.setObjectName(u"intragPage")
        self.verticalLayout_intragPage = QVBoxLayout(self.intragPage)
        self.verticalLayout_intragPage.setObjectName(u"verticalLayout_intragPage")
        self.label_intragPage = QLabel(self.intragPage)
        self.label_intragPage.setObjectName(u"label_intragPage")
        self.label_intragPage.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_intragPage.setText(QCoreApplication.translate("MainWindow", u"Página Intrag", None))
        self.verticalLayout_intragPage.addWidget(self.label_intragPage)
        self.mainPages.addWidget(self.intragPage)
        
        self.verticalLayout_11.addWidget(self.mainPages)
        self.horizontalLayout_10.addWidget(self.mainPagesCont)

        # ----- Menu lateral direito (deslizante) -----
        self.rightMenu = QCustomSlideMenu(self.mainContents)
        self.rightMenu.setObjectName(u"rightMenu")
        self.rightMenu.setMinimumSize(QSize(200, 0))
        self.verticalLayout_17 = QVBoxLayout(self.rightMenu)
        self.verticalLayout_17.setSpacing(5)
        self.verticalLayout_17.setObjectName(u"verticalLayout_17")
        self.verticalLayout_17.setContentsMargins(5, 5, 5, 5)
        self.widget_6 = QWidget(self.rightMenu)
        self.widget_6.setObjectName(u"widget_6")
        self.horizontalLayout_12 = QHBoxLayout(self.widget_6)
        self.horizontalLayout_12.setObjectName(u"horizontalLayout_12")
        self.label_15 = QLabel(self.widget_6)
        self.label_15.setObjectName(u"label_15")
        self.horizontalLayout_12.addWidget(self.label_15)
        self.closeRightMenuBtn = QPushButton(self.widget_6)
        self.closeRightMenuBtn.setObjectName(u"closeRightMenuBtn")
        self.closeRightMenuBtn.setIcon(icon_close)
        self.horizontalLayout_12.addWidget(self.closeRightMenuBtn)
        self.verticalLayout_17.addWidget(self.widget_6)
        self.rightMenuPages = QCustomQStackedWidget(self.rightMenu)
        self.rightMenuPages.setObjectName(u"rightMenuPages")
        self.notificationsPage = QWidget()
        self.notificationsPage.setObjectName(u"notificationsPage")
        self.verticalLayout_18 = QVBoxLayout(self.notificationsPage)
        self.verticalLayout_18.setObjectName(u"verticalLayout_18")
        self.label_16 = QLabel(self.notificationsPage)
        self.label_16.setObjectName(u"label_16")
        self.label_16.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_18.addWidget(self.label_16)
        self.rightMenuPages.addWidget(self.notificationsPage)
        self.morePage = QWidget()
        self.morePage.setObjectName(u"morePage")
        self.verticalLayout_19 = QVBoxLayout(self.morePage)
        self.verticalLayout_19.setObjectName(u"verticalLayout_19")
        self.label_17 = QLabel(self.morePage)
        self.label_17.setObjectName(u"label_17")
        self.label_17.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_19.addWidget(self.label_17)
        self.rightMenuPages.addWidget(self.morePage)
        self.profilePage = QWidget()
        self.profilePage.setObjectName(u"profilePage")
        self.verticalLayout_20 = QVBoxLayout(self.profilePage)
        self.verticalLayout_20.setObjectName(u"verticalLayout_20")
        self.label_18 = QLabel(self.profilePage)
        self.label_18.setObjectName(u"label_18")
        self.label_18.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout_20.addWidget(self.label_18)
        self.rightMenuPages.addWidget(self.profilePage)
        self.verticalLayout_17.addWidget(self.rightMenuPages)
        self.horizontalLayout_10.addWidget(self.rightMenu)

        self.verticalLayout_10.addWidget(self.mainContents)

        # ---------- Footer ----------
        self.footer = QWidget(self.mainBody)
        self.footer.setObjectName(u"footer")
        self.horizontalLayout_4 = QHBoxLayout(self.footer)
        self.horizontalLayout_4.setSpacing(5)
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.horizontalLayout_4.setContentsMargins(5, 5, 0, 0)
        self.label_19 = QLabel(self.footer)
        self.label_19.setObjectName(u"label_19")
        self.horizontalLayout_4.addWidget(self.label_19, 0, Qt.AlignmentFlag.AlignLeft)
        self.frame_2 = QFrame(self.footer)
        self.frame_2.setObjectName(u"frame_2")
        self.frame_2.setFrameShape(QFrame.Shape.StyledPanel)
        self.frame_2.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_5 = QHBoxLayout(self.frame_2)
        self.horizontalLayout_5.setObjectName(u"horizontalLayout_5")
        self.label_20 = QLabel(self.frame_2)
        self.label_20.setObjectName(u"label_20")
        self.horizontalLayout_5.addWidget(self.label_20)
        self.activityProgress = QProgressBar(self.frame_2)
        self.activityProgress.setObjectName(u"activityProgress")
        self.activityProgress.setMaximumSize(QSize(16777215, 10))
        self.activityProgress.setValue(24)
        self.activityProgress.setTextVisible(False)
        self.horizontalLayout_5.addWidget(self.activityProgress)
        self.horizontalLayout_4.addWidget(self.frame_2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.sizeGrip = QFrame(self.footer)
        self.sizeGrip.setObjectName(u"sizeGrip")
        self.sizeGrip.setMinimumSize(QSize(15, 15))
        self.sizeGrip.setMaximumSize(QSize(15, 15))
        self.sizeGrip.setFrameShape(QFrame.Shape.StyledPanel)
        self.sizeGrip.setFrameShadow(QFrame.Shadow.Raised)
        self.horizontalLayout_4.addWidget(self.sizeGrip, 0, Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignBottom)
        self.verticalLayout_10.addWidget(self.footer, 0, Qt.AlignmentFlag.AlignBottom)
        self.horizontalLayout.addWidget(self.mainBody)
        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        self.centerMenuPages.setCurrentIndex(0)  
           # -------------------------------------------------------------------------
        # Página de Pesquisa de Commodities - Painel central totalmente responsivo
        # -------------------------------------------------------------------------
        # Este widget representa a área de busca de commodities, exibida no painel central do sistema.
        # Ele é responsável por exibir o painel de busca, garantindo responsividade
        # e adaptação ao tamanho da janela principal, sem limitações de largura.
        self.commoditiesSearchPage = QWidget()
        self.commoditiesSearchPage.setObjectName(u"commoditiesSearchPage")
        # SizePolicy expansivo para crescimento junto com a área central
        self.commoditiesSearchPage.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding))

        # -------------------------------------------------------------------------
        # Layout vertical principal da página de commodities
        # -------------------------------------------------------------------------
        # Organiza os elementos de cima para baixo, aplicando margens e espaçamento
        # para um visual agradável e evitando elementos colados nas bordas.
        self.verticalLayout_commoditiesSearchPage = QVBoxLayout(self.commoditiesSearchPage)
        self.verticalLayout_commoditiesSearchPage.setObjectName(u"verticalLayout_commoditiesSearchPage")
        self.verticalLayout_commoditiesSearchPage.setContentsMargins(20, 20, 20, 20)
        self.verticalLayout_commoditiesSearchPage.setSpacing(15)

        # -------------------------------------------------------------------------
        # Painel de busca de commodities - Widget expansivo e centralizado
        # -------------------------------------------------------------------------
        # Este widget contém os campos de filtro e botões de busca.
        # Ele é filho direto da página de commodities e ocupa o máximo de espaço possível.
        self.searchCommoditiesWidget = QWidget(self.commoditiesSearchPage)
        self.searchCommoditiesWidget.setObjectName(u"searchCommoditiesWidget")
        # SizePolicy expansivo: garante expansão máxima em largura/altura
        self.searchCommoditiesWidget.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding))
        # Estilo visual: cor, borda e padding. Não usar max-width nem margens CSS!
        self.searchCommoditiesWidget.setStyleSheet("""
            background: #fff;
            border: 1px solid #d0d8e0;
            border-radius: 7px;
            padding: 18px 22px;
        """)
        # O painel de busca inicia oculto, só aparece quando o usuário clica em "Search" no menu.
        self.searchCommoditiesWidget.hide()

        # -------------------------------------------------------------------------
        # Layout em grade para filtros e botões do bloco de pesquisa
        # -------------------------------------------------------------------------
        # Organiza os campos de filtro e botões em linhas e colunas, facilitando a disposição responsiva.
        # Customizado para menor espaçamento entre linhas (labels e combobox), como pedido.
        self.grid_searchCommodities = QGridLayout(self.searchCommoditiesWidget)
        self.grid_searchCommodities.setObjectName(u"grid_searchCommodities")
        self.grid_searchCommodities.setHorizontalSpacing(8)  # Diminui o espaçamento horizontal
        self.grid_searchCommodities.setVerticalSpacing(6)    # Diminui o espaçamento vertical

        # -------------------------------------------------------------------------
        # Label de cabeçalho azul escuro - título do painel de busca
        # -------------------------------------------------------------------------
        # Exibe o título da área de consulta, destacado em azul escuro (#000080), alinhado ao centro.
        # O texto também é centralizado via alinhamento e CSS.
        self.label_searchHeader = QLabel(self.searchCommoditiesWidget)
        self.label_searchHeader.setObjectName(u"label_searchHeader")
        self.label_searchHeader.setText("Consulta ao Cadastro de Commodities")
        self.label_searchHeader.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_searchHeader.setStyleSheet(
            "background-color: #000060; color: white; font-size: 15px; font-weight: bold; padding: 7px 8px; border-radius: 3px;"
        )
        self.label_searchHeader.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.grid_searchCommodities.addWidget(self.label_searchHeader, 0, 0, 1, 2)

        # -------------------------------------------------------------------------
        # ComboBox Código - Campo inserido no topo do formulário de busca
        # -------------------------------------------------------------------------
        # Permite ao usuário filtrar pelo código da commodity. Preenchido pelo backend.
        self.label_codigo = QLabel(self.searchCommoditiesWidget)
        self.label_codigo.setText("Código:")
        self.label_codigo.setObjectName(u"label_codigo")
        self.combo_codigo = QComboBox(self.searchCommoditiesWidget)
        self.combo_codigo.setObjectName(u"combo_codigo")
        self.combo_codigo.setMinimumWidth(210)
        self.combo_codigo.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_codigo.addItem("Selecione o código")
        self.grid_searchCommodities.addWidget(self.label_codigo, 1, 0)
        self.grid_searchCommodities.addWidget(self.combo_codigo, 1, 1)

        # -------------------------------------------------------------------------
        # ComboBox de Bolsa - Permite ao usuário selecionar a bolsa de negociação
        # -------------------------------------------------------------------------
        self.label_bolsa = QLabel(self.searchCommoditiesWidget)
        self.label_bolsa.setText("Bolsa:")
        self.label_bolsa.setObjectName(u"label_bolsa")
        self.combo_bolsa = QComboBox(self.searchCommoditiesWidget)
        self.combo_bolsa.setObjectName(u"combo_bolsa")
        self.combo_bolsa.setMinimumWidth(210)
        self.combo_bolsa.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_bolsa.addItem("Selecione a bolsa")  # Preenchido via backend ou manual
        self.grid_searchCommodities.addWidget(self.label_bolsa, 2, 0)
        self.grid_searchCommodities.addWidget(self.combo_bolsa, 2, 1)

        # -------------------------------------------------------------------------
        # ComboBox de Commodity - Permite ao usuário selecionar a commodity
        # -------------------------------------------------------------------------
        self.label_commodity = QLabel(self.searchCommoditiesWidget)
        self.label_commodity.setText("Mercadoria :")
        self.label_commodity.setObjectName(u"label_commodity")
        self.combo_commodity = QComboBox(self.searchCommoditiesWidget)
        self.combo_commodity.setObjectName(u"combo_commodity")
        self.combo_commodity.setMinimumWidth(210)
        self.combo_commodity.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_commodity.addItem("Selecione a Mercadoria")
        self.grid_searchCommodities.addWidget(self.label_commodity, 3, 0)
        self.grid_searchCommodities.addWidget(self.combo_commodity, 3, 1)

        # -------------------------------------------------------------------------
        # ComboBox de mês de vencimento - Permite ao usuário selecionar o mês
        # -------------------------------------------------------------------------
        self.label_mes = QLabel(self.searchCommoditiesWidget)
        self.label_mes.setText("Mês de vencimento:")
        self.label_mes.setObjectName(u"label_mes")
        self.combo_mes = QComboBox(self.searchCommoditiesWidget)
        self.combo_mes.setObjectName(u"combo_mes")
        self.combo_mes.setMinimumWidth(210)
        self.combo_mes.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_mes.addItem("Selecione o mês")
        self.grid_searchCommodities.addWidget(self.label_mes, 4, 0)
        self.grid_searchCommodities.addWidget(self.combo_mes, 4, 1)

        # -------------------------------------------------------------------------
        # ComboBox de ano de vencimento - Permite ao usuário selecionar o ano
        # -------------------------------------------------------------------------
        self.label_ano = QLabel(self.searchCommoditiesWidget)
        self.label_ano.setText("Ano Vencimento :")
        self.label_ano.setObjectName(u"label_ano")
        self.combo_ano = QComboBox(self.searchCommoditiesWidget)
        self.combo_ano.setObjectName(u"combo_ano")
        self.combo_ano.setMinimumWidth(210)
        self.combo_ano.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_ano.addItem("Selecione o ano")
        self.grid_searchCommodities.addWidget(self.label_ano, 5, 0)
        self.grid_searchCommodities.addWidget(self.combo_ano, 5, 1)

        # -------------------------------------------------------------------------
        # ComboBox Status - Campo inserido abaixo dos outros filtros
        # -------------------------------------------------------------------------
        # Permite filtrar pelo status do cadastro da commodity.
        self.label_status = QLabel(self.searchCommoditiesWidget)
        self.label_status.setText("Status:")
        self.label_status.setObjectName(u"label_status")
        self.combo_status = QComboBox(self.searchCommoditiesWidget)
        self.combo_status.setObjectName(u"combo_status")
        self.combo_status.setMinimumWidth(210)
        self.combo_status.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))
        self.combo_status.addItem("Selecione o status")
        self.grid_searchCommodities.addWidget(self.label_status, 6, 0)
        self.grid_searchCommodities.addWidget(self.combo_status, 6, 1)

        # -------------------------------------------------------------------------
        # Botões de ação - Pesquisa e Limpar Campos
        # -------------------------------------------------------------------------
        # Botão "Pesquisar" aciona a consulta das cotações conforme os filtros selecionados.
        # Estilizado com fundo #000060, texto branco e negrito como solicitado.
        self.searchFilterCommoditiesBtn = QPushButton(self.searchCommoditiesWidget)
        self.searchFilterCommoditiesBtn.setText("Pesquisar")
        self.searchFilterCommoditiesBtn.setObjectName(u"searchFilterCommoditiesBtn")
        self.searchFilterCommoditiesBtn.setStyleSheet("""
            min-width: 120px; 
            font-weight: bold; 
            background-color: #000060; 
            color: white;
            border-radius: 5px;
            padding: 7px 0px;
        """)
        self.searchFilterCommoditiesBtn.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))

        # Botão "Limpar Campos" reseta todos os filtros do painel de busca.
        # Mesmo estilo do botão pesquisar, conforme padrão visual solicitado.
        self.deletFilterCommoditiesBtn = QPushButton(self.searchCommoditiesWidget)
        self.deletFilterCommoditiesBtn.setText("Limpar Campos")
        self.deletFilterCommoditiesBtn.setObjectName(u"deletFilterCommoditiesBtn")
        self.deletFilterCommoditiesBtn.setStyleSheet("""
            min-width: 120px; 
            font-weight: bold; 
            background-color: #000060; 
            color: white;
            border-radius: 5px;
            padding: 7px 0px;
        """)
        self.deletFilterCommoditiesBtn.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed))

        self.grid_searchCommodities.addWidget(self.searchFilterCommoditiesBtn, 7, 0)
        self.grid_searchCommodities.addWidget(self.deletFilterCommoditiesBtn, 7, 1)

        # -------------------------------------------------------------------------
        # Adiciona o painel de busca ao layout vertical principal da página de commodities
        # -------------------------------------------------------------------------
        # O painel de busca recebe stretch=1, garantindo que ocupe o máximo de espaço
        # possível na área central da página, ajustando-se proporcionalmente ao redimensionamento da janela.
        self.verticalLayout_commoditiesSearchPage.addWidget(
            self.searchCommoditiesWidget,
            stretch=1
        )

        # -------------------------------------------------------------------------
        # Adiciona commoditiesSearchPage ao QStackedWidget principal de páginas (mainPages)
        # -------------------------------------------------------------------------
        # Permite navegação entre as diferentes páginas centrais do sistema (ex.: Home, NDF, Option, etc)
        self.mainPages.addWidget(self.commoditiesSearchPage)

                # -------------------------------------------------------------------------
        # Área de Resultados da Busca - Tabela de Commodities e Botões de Ação
        # -------------------------------------------------------------------------
        # Esta área fica logo abaixo (ou ao lado) do painel de busca. Ela é responsável por exibir:
        # - A tabela com os resultados da busca de commodities, preenchida dinamicamente pelo backend.
        # - Os botões de ação: "Back" (voltar para os filtros) e "Export" (exportar resultados).
        # Por padrão, esta área está oculta e só aparece após uma busca bem-sucedida.
        self.widget_commoditiesTableArea = QWidget(self.commoditiesSearchPage)
        self.widget_commoditiesTableArea.setObjectName("widget_commoditiesTableArea")
        self.widget_commoditiesTableArea.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding))
        self.widget_commoditiesTableArea.hide()  # Só mostra quando há resultado

        # Layout vertical para organizar a tabela e os botões em blocos separados
        self.vbox_commoditiesTableArea = QVBoxLayout(self.widget_commoditiesTableArea)
        self.vbox_commoditiesTableArea.setObjectName("vbox_commoditiesTableArea")
        self.vbox_commoditiesTableArea.setContentsMargins(0, 0, 0, 0)
        self.vbox_commoditiesTableArea.setSpacing(10)

        # ---------- Tabela de Resultados (dinâmica) ----------
        # Widget placeholder para a tabela de resultados (QTableWidget), 
        # que será preenchida pelo controller conforme os dados retornados da busca.
        # Por padrão, começa vazio e invisível.
        self.tabela_resultados = None  # Será instanciado dinamicamente no controller

        # ---------- Barra de Botões (Back e Export) ----------
        # Layout horizontal para manter os botões lado a lado, alinhados à direita.
        self.hbox_commoditiesTableBtns = QHBoxLayout()
        self.hbox_commoditiesTableBtns.setObjectName("hbox_commoditiesTableBtns")
        self.hbox_commoditiesTableBtns.setContentsMargins(0, 0, 0, 0)
        self.hbox_commoditiesTableBtns.setSpacing(10)

        # Botão "Back" - retorna para o painel de filtros
        self.backSearchCommoditiesBtn = QPushButton(self.widget_commoditiesTableArea)
        self.backSearchCommoditiesBtn.setObjectName("backSearchCommoditiesBtn")
        self.backSearchCommoditiesBtn.setText("Back")
        self.backSearchCommoditiesBtn.setStyleSheet(
            "QPushButton { margin: 5px 10px 5px 0px; background-color: #000060; color: #fff; border-radius: 6px; font-weight: bold; font-size: 14px; padding: 8px 16px; }"
            "QPushButton:hover { background-color: #7fffd4; color: #000060; }"
        )

        # Botão "Export" - exporta o resultado em CSV, apenas ícone, mesmo azul do Back
        self.exportCommoditiesSearchBtn = QPushButton(self.widget_commoditiesTableArea)
        self.exportCommoditiesSearchBtn.setObjectName("exportCommoditiesSearchBtn")
        icon_exportComm = QIcon()
        icon_exportComm.addFile(u":/feather/icons/feather/download.png", QSize(), QIcon.Normal, QIcon.Off)
        self.exportCommoditiesSearchBtn.setIcon(icon_exportComm)
        self.exportCommoditiesSearchBtn.setToolTip("Export search results to CSV")
        self.exportCommoditiesSearchBtn.setText("")
        self.exportCommoditiesSearchBtn.setStyleSheet(
            "QPushButton { margin: 5px 0px 5px 10px; background-color: #000060; color: #fff; border-radius: 6px; font-weight: bold; font-size: 14px; padding: 8px; }"
            "QPushButton:hover { background-color: #7fffd4; color: #000060; }"
        )

        # Adiciona os botões ao layout horizontal (lado a lado)
        self.hbox_commoditiesTableBtns.addWidget(self.backSearchCommoditiesBtn)
        self.hbox_commoditiesTableBtns.addWidget(self.exportCommoditiesSearchBtn)
        self.hbox_commoditiesTableBtns.addStretch(1)  # Empurra os botões para a esquerda

        # Adiciona os widgets ao layout vertical da área de resultados
        # (a tabela será adicionada dinamicamente pelo controller)
        self.vbox_commoditiesTableArea.addLayout(self.hbox_commoditiesTableBtns)

        # Por fim, adiciona a área de resultados (painel de tabela) ao layout principal da página
        self.verticalLayout_commoditiesSearchPage.addWidget(
            self.widget_commoditiesTableArea,
            stretch=10  # Deixa a tabela ocupar o máximo possível do espaço disponível
        )
         
       

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.menuBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Side Menu", None))
        self.menuBtn.setText("")
        self.homeBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Go Home", None))
        self.homeBtn.setText(QCoreApplication.translate("MainWindow", u"Home", None))
        self.ndfBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View NDF", None))
        self.ndfBtn.setText(QCoreApplication.translate("MainWindow", u"NDF", None))
        self.optBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Options", None))
        self.optBtn.setText(QCoreApplication.translate("MainWindow", u"Option", None))
        self.swaptBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Swaps", None))
        self.swaptBtn.setText(QCoreApplication.translate("MainWindow", u"Swap", None))
        self.counterpartyBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Counterparty", None))
        self.counterpartyBtn.setText(QCoreApplication.translate("MainWindow", u"Counterparty", None))
        self.commodiBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Commoditiess", None))
        self.commodiBtn.setText(QCoreApplication.translate("MainWindow", u"Commodities", None))
        self.metricsBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Metrics", None))
        self.metricsBtn.setText(QCoreApplication.translate("MainWindow", u"Metrics", None))
        self.confirmationsBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Confirmations", None))
        self.confirmationsBtn.setText(QCoreApplication.translate("MainWindow", u"Confirmations", None))
        self.settingsBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Go to settings", None))
        self.settingsBtn.setText(QCoreApplication.translate("MainWindow", u"Settings", None))
        self.infoBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View information", None))
        self.infoBtn.setText(QCoreApplication.translate("MainWindow", u"Information", None))
        self.helpBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Get help", None))
        self.helpBtn.setText(QCoreApplication.translate("MainWindow", u"Help", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"Center Menu", None))
        self.closeCenterMenuBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Close menu", None))
        self.closeCenterMenuBtn.setText("")
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"Settings", None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u"Theme", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"Information", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"Help", None))
        self.label_ndf.setText(QCoreApplication.translate("MainWindow", u"Subpágina NDF", None))
        self.label_opcao.setText(QCoreApplication.translate("MainWindow", u"Subpágina Option", None))
        self.label_swap.setText(QCoreApplication.translate("MainWindow", u"Subpágina Swap", None))
        self.label_counterparty.setText(QCoreApplication.translate("MainWindow", u"Subpágina Counterparty", None))
        self.label_metrics.setText(QCoreApplication.translate("MainWindow", u"Subpágina Metrics", None))
        self.label_confirmations.setText(QCoreApplication.translate("MainWindow", u"Subpágina Confirmations", None))
        self.kpiBtn.setText(QCoreApplication.translate("MainWindow", u"KPI", None))
        self.adhocBtn.setText(QCoreApplication.translate("MainWindow", u"Ad-Hoc", None))
        self.ndfSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.ndfRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register", None))
        self.ndfSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement", None))
        self.optSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.optRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register", None))
        self.optSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement", None))
        self.swapSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.swapRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register", None))
        self.swapSettleBtn.setText(QCoreApplication.translate("MainWindow", u"Settlement", None))
        self.counterpartyRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register", None))
        self.counterpartySearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.commoditiesRegisterBtn.setText(QCoreApplication.translate("MainWindow", u"Register", None))
        self.commoditiesSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.metricsBtn.setText(QCoreApplication.translate("MainWindow", u"Metrics", None))
        self.confirmationsBtn.setText(QCoreApplication.translate("MainWindow", u"Confirmations", None))
        self.kpiBtn.setText(QCoreApplication.translate("MainWindow", u"KPI", None))
        self.adhocBtn.setText(QCoreApplication.translate("MainWindow", u"Ad-Hoc", None))
        self.NDFConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"NDF", None))
        self.optionConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"Option", None))
        self.swapConfirmationBtn.setText(QCoreApplication.translate("MainWindow", u"Swap", None))
        self.label_metrics.setText(QCoreApplication.translate("MainWindow", u"Subpágina Metrics", None))
        self.label_confirmations.setText(QCoreApplication.translate("MainWindow", u"Subpágina Confirmations", None))
        # INTRAG
        self.intragBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View Intrag", None))
        self.intragBtn.setText(QCoreApplication.translate("MainWindow", u"Intrag", None))
        self.intragSearchBtn.setText(QCoreApplication.translate("MainWindow", u"Search", None))
        self.intragBoletaBtn.setText(QCoreApplication.translate("MainWindow", u"Instruction", None))
        self.titleTxt.setText(QCoreApplication.translate("MainWindow", u"OTC SYSTEM", None))
        self.notificationBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View notifications", None))
        self.notificationBtn.setText("")
        self.moreBtn.setToolTip(QCoreApplication.translate("MainWindow", u"View more", None))
        self.moreBtn.setText("")
        self.profileBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Go to profile", None))
        self.profileBtn.setText("")
        self.label_9.setText("")
        self.searchInp.setPlaceholderText(QCoreApplication.translate("MainWindow", u"Search...", None))
        self.searchBtn.setText(QCoreApplication.translate("MainWindow", u"Ctrl+K", None))
        self.minimizeBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Minimize window", None))
        self.minimizeBtn.setText("")
        self.restoreBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Restore window", None))
        self.restoreBtn.setText("")
        self.closeBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Close app", None))
        self.closeBtn.setText("")
        self.label_8.setText(QCoreApplication.translate("MainWindow", u"Home Page", None))
        self.label_10.setText(QCoreApplication.translate("MainWindow", u"Página NDF", None))
        self.label_11.setText(QCoreApplication.translate("MainWindow", u"Página Option", None))
        self.label_12.setText(QCoreApplication.translate("MainWindow", u"Página Swap", None))
        self.label_13.setText(QCoreApplication.translate("MainWindow", u"Página Counterparty", None))
        self.label_14.setText(QCoreApplication.translate("MainWindow", u"Página commodities", None))
        self.label_metricsPage.setText(QCoreApplication.translate("MainWindow", u"Página Metrics", None))
        self.label_confirmationsPage.setText(QCoreApplication.translate("MainWindow", u"Página Confirmations", None))
        self.label_15.setText(QCoreApplication.translate("MainWindow", u"Right Menu", None))
        self.closeRightMenuBtn.setToolTip(QCoreApplication.translate("MainWindow", u"Close menu", None))
        self.closeRightMenuBtn.setText("")
        self.label_16.setText(QCoreApplication.translate("MainWindow", u"Notifications", None))
        self.label_17.setText(QCoreApplication.translate("MainWindow", u"More", None))
        self.label_18.setText(QCoreApplication.translate("MainWindow", u"Profile", None))
        self.label_19.setText(QCoreApplication.translate("MainWindow", u"", None))
        self.label_20.setText(QCoreApplication.translate("MainWindow", u"Theme Progress", None))