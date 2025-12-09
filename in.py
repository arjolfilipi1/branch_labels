from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QMessageBox,QDialog,QFormLayout,QDialogButtonBox,QListWidget,
    QListWidgetItem,QSpinBox,QTextEdit,QScrollArea,QCheckBox
)
from PyQt5.QtCore import Qt, QStringListModel, QSettings,QLocale
from PyQt5 import QtGui
import os,sys
from PyQt5.QtPrintSupport import QPrinterInfo
import pyodbc
from PIL import Image,ImageWin
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
# from pylibdmtx.pylibdmtx import encode
from reportlab.lib.utils import ImageReader
from io import BytesIO
import tempfile,time
import qrcode
import win32print,win32ui,win32api
# Convert mm to pixels
import fitz  # PyMuPDF
from datetime import datetime


locale = QLocale(QLocale.English)
QLocale.setDefault(locale)

    
# Application settings
SETTINGS = QSettings("Forschner", "Label dege")



class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("App Settings")
        self.resize(430, 600)
        # Create a scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        # Create a widget to hold the layout
        content_widget = QWidget()
        layout = QFormLayout(content_widget)
        
        # use qr
        self.checkbox = QCheckBox("Use qr", self)
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.checkbox)
        
        # use qr
        self.check2box = QCheckBox("", self)
        check2box_layout = QHBoxLayout()
        check2box_layout.addWidget(self.check2box)
        
        # Database path
        self.db_path_edit = QLineEdit()
        self.db_path_edit.setPlaceholderText("Vendos Perdoruesin e xPPS")
        db_layout = QHBoxLayout()
        db_layout.addWidget(self.db_path_edit)
        
        # Images path
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Vendos passwordin e xPPS")
        password_layout = QHBoxLayout()
        password_layout.addWidget(self.password_edit)
        
        
        
        # sql2
        self.sql = """
                SELECT STRU.STKOMP AS XVK_MODULE 
                FROM BIDBD220.STRU STRU 
                WHERE STRU.STWKNR = '000' 
                  AND STRU.STFIRM = '1' 
                  AND STRU.STBGNR = ?
            """
        self.con = "Driver={{IBM i Access ODBC Driver}};System=192.168.100.35;UID={};PWD={};DBQ=QGPL;"
        self.sql_edit = QTextEdit()
        self.sql_edit.setPlaceholderText(self.sql)
        slq_layout = QHBoxLayout()
        slq_layout.addWidget(self.sql_edit)
        
        self.con_edit = QTextEdit()
        self.con_edit.setPlaceholderText(self.con)
        con_layout = QHBoxLayout()
        con_layout.addWidget(self.con_edit)
        
        # sql2
        self.sql2 = """
    SELECT STKOMP AS P_N, STEIGW AS Q_C FROM BIDBD220.STRU WHERE BIDBD220.STRU.STBGNR= '{}'
AND (( BIDBD220.STRU.STGUVO = 0 AND BIDBD220.STRU.STGUBI = 0)
OR (BIDBD220.STRU.STGUVO <= {} AND BIDBD220.STRU.STGUBI = 0) 
OR (BIDBD220.STRU.STGUBI >= {} AND BIDBD220.STRU.STGUVO = 0)
 )
 ;
    """
        self.con2 = """SELECT LGBS.LSTENR, LGBS.LSLANR, LGBS.LSLGBE
FROM BIDBD220.LGBS LGBS
WHERE (LGBS.LSFIRM='1') AND (LGBS.LSWKNR='361')
AND (LGBS.LSLGBE>0)
AND LSLANR IN ('3G','3H') AND LSTENR IN ({})"""
        
        self.sql2_edit = QTextEdit()
        self.sql2_edit.setPlaceholderText(self.sql2)
        slq2_layout = QHBoxLayout()
        slq2_layout.addWidget(self.sql2_edit)
        
        self.con2_edit = QTextEdit()
        self.con2_edit.setPlaceholderText(self.con2)
        con2_layout = QHBoxLayout()
        con2_layout.addWidget(self.con2_edit)
        
        
        # Add to form
        layout.addRow("Perdor qr code:", checkbox_layout)
        layout.addRow("Listo sipas parapergatitjes:", check2box_layout)
        layout.addRow("xPPS user:", db_layout)
        layout.addRow("Password:", password_layout)
       
        layout.addRow("SQL e xpps:", slq_layout)
        layout.addRow("Conn e xpps:", con_layout)
        layout.addRow("SQL e Materialeve:", slq2_layout)
        layout.addRow("SQL e sasise ne magazine:", con2_layout)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        
        # Set the content widget to scroll area
        scroll.setWidget(content_widget)
        
        # Set scroll area as the main layout
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll)
        
        # Load current settings
        self.load_settings()

    
    def load_settings(self):
        if SETTINGS.value("use_qr", 'false') != 'false':
            self.checkbox.setChecked(True)
        if SETTINGS.value("sort", 'false') != 'false':
            self.check2box.setChecked(True)
        self.db_path_edit.setText(SETTINGS.value("xpps/user", "FILIPI"))
        self.password_edit.setText(SETTINGS.value("xpps/password", "a110033"))
       
        self.sql_edit.setPlainText(SETTINGS.value("xpps/sql", self.sql))
        self.con_edit.setPlainText(SETTINGS.value("xpps/con", self.con))
        self.sql2_edit.setPlainText(SETTINGS.value("komax/sql", self.sql2))
        self.con2_edit.setPlainText(SETTINGS.value("komax/con", self.con2))
    
    def save_settings(self):
        SETTINGS.setValue("use_qr", self.checkbox.isChecked())
        SETTINGS.setValue("sort", self.check2box.isChecked())
        SETTINGS.setValue("xpps/user", self.db_path_edit.text())
        SETTINGS.setValue("xpps/password", self.password_edit.text())
       
        SETTINGS.setValue("xpps/sql", self.sql_edit.toPlainText())
        SETTINGS.setValue("xpps/con", self.con_edit.toPlainText())
        SETTINGS.setValue("bom/sql", self.sql2_edit.toPlainText())
        SETTINGS.setValue("komax/con", self.con2_edit.toPlainText())

class MyApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Printimi i Labelave te degeve Daimler Wörth")
        self.setup_ui()
        self.apply_styles()
        self.manual_select = 0
        self.tb = []
    def get_settings(self):
        self.sql = SETTINGS.value("xpps/sql", self.sql)
        self.con = SETTINGS.value("xpps/con", self.con)
        self.SQL_TEMPLATE_VODICE = SETTINGS.value("bom/sql", self.SQL_TEMPLATE_VODICE)
        self.con2 = SETTINGS.value("komax/con", self.con2)
        self.sort = SETTINGS.value("sort", False) == 'true'
    def setup_ui(self):
        self.sort = SETTINGS.value("sort", False) == 'true'
        self.SQL_TEMPLATE_VODICE = """
SELECT STKOMP AS P_N, STEIGW AS Q_C FROM BIDBD220.STRU WHERE BIDBD220.STRU.STBGNR= '{}'
AND (( BIDBD220.STRU.STGUVO = 0 AND BIDBD220.STRU.STGUBI = 0)
OR (BIDBD220.STRU.STGUVO <= {} AND BIDBD220.STRU.STGUBI = 0) 
OR (BIDBD220.STRU.STGUBI >= {} AND BIDBD220.STRU.STGUVO = 0)
 )
 ;
"""
        self.con = "Driver={{IBM i Access ODBC Driver}};System=192.168.100.35;UID={};PWD={};DBQ=QGPL;"
        self.con2 = """SELECT LGBS.LSTENR, LGBS.LSLANR, LGBS.LSLGBE
FROM BIDBD220.LGBS LGBS
WHERE (LGBS.LSFIRM='1') AND (LGBS.LSWKNR='361')
AND (LGBS.LSLGBE>0)
AND LSLANR IN ('3G','3H') AND LSTENR IN ({})"""
        self.sql = """
                SELECT STRU.STKOMP AS XVK_MODULE 
                FROM BIDBD220.STRU STRU 
                WHERE STRU.STWKNR = '000' 
                  AND STRU.STFIRM = '1' 
                  AND STRU.STBGNR = ?
            """
        self.get_settings()
        self.connection = None
        self.auf = None
        self.qty = None
        self.xvk = ""
        self.modul =""
        self.to_print_all = []
        self.deget = {}
        
        self.setWindowIcon(QtGui.QIcon('logo.ico'))
        # Layouts
        main_layout = QVBoxLayout()
        top_bar = QHBoxLayout()
        form_layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        status_layout = QVBoxLayout()
        
        
        # Top bar with Settings button
        self.settings_button = QPushButton("Parametra")
        icon = QtGui.QIcon('s.png')
        self.settings_button.setIcon(icon)
        self.settings_button.clicked.connect(self.open_settings)
        top_bar.addStretch()
        top_bar.addWidget(self.settings_button)

        # Input Field
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("34477595")
        self.input_field.returnPressed.connect(self.onScan)
        form_layout.addWidget(QLabel("Vendos XVK:"))
        form_layout.addWidget(self.input_field)
        
        
        # module Field
        self.module_field = QLineEdit()
        self.module_field.setPlaceholderText("444000108010")
        form_layout.addWidget(QLabel("Vendos modulin:"))
        form_layout.addWidget(self.module_field)
        


        # Buttons

        self.reset_button = QPushButton("Reset")
        self.reset_button.setToolTip("Fshij te dhenat e meparshme")
        self.reset_button.clicked.connect(self.reset)
        
        button_layout.addWidget(self.reset_button)

        

        # scroll = QScrollArea()
        # scroll.setWidgetResizable(True)
        self.status = QLabel()
        self.status.setFixedHeight(22)
        # scroll.setWidget(self.status)
        status_layout.addWidget(self.status)
        # status_layout.setSizeConstraint(QtGui.QLayout.SetFixedSize)
        

        # List Widget
        self.list_widget = QListWidget()
        # self.list_widget.itemClicked.connect(self.module_selected)
        form_layout.addWidget(QLabel("Modulet:"))
        form_layout.addWidget(self.list_widget)
        
        # Assemble layouts
        main_layout.addLayout(top_bar)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        main_layout.addLayout(status_layout)

        self.setLayout(main_layout)
    def get_conn(self):
        if self.connection:
            return self.connection
        else:
            uid = SETTINGS.value("xpps/user", "FILIPI")
            pw = SETTINGS.value("xpps/password", "a110033")
             # Connection string
            connection_string = (
 
                self.con.format(uid,pw)
            
                )
            conn = pyodbc.connect(connection_string)
            return conn
    def onScan(self):
        text = self.input_field.text().strip()
        print(text[:-8])
        try:
            text = text[:-8]
            text = int(text)
            self.status.setText("Porosia u lexua")
        except Exception as e:
            self.status.setText(str(e))
            return None

        if text:
            self.status.setText("Duke u lidhur me XPPS")
            

            conn = self.get_conn()
            sql = self.sql
            try:
                self.list_widget.clear()
                dega= {}
                m_d= {}
                # Connect
                
                cursor = conn.cursor()

                # Execute query with parameter
                cursor.execute(sql, (text))

                # Fetch and print results
                rows = cursor.fetchall()
                module_list = []
                for row in rows:
                    self.auf = row.P_N
                    self.qty = row.Q_T
                if self.auf:
                    mat_list = {}
                    date = int(datetime.today().strftime('%Y%m%d'))
                    sql = self.SQL_TEMPLATE_VODICE.format(self.auf,date,date)
                    cursor.execute(sql)
                    rows = cursor.fetchall()
                    for row in rows:
                        if row.Q_C%1==0:
                            pn = row.P_N
                            if pn not in mat_list:
                                mat_list[pn] = int(row.Q_C)
                            else:
                                mat_list[pn] += int(row.Q_C)
                    li = [ "'"+str(x)+ "'," for x in mat_list.keys() ]
                    con2 = self.con2.format("".join(li)[:-1])
                    # print(con2)
                    cursor.execute(con2)
                    rows = cursor.fetchall()
                    mat_ava ={}
                    for row in rows:
                        if (row.LSTENR) not in mat_ava:
                            mat_ava[row.LSTENR] = [row.LSLGBE,row.LSLANR]
                        else:
                            mat_ava[row.LSTENR] = [mat_ava[row.LSTENR][0] + row.LSLGBE, mat_ava[row.LSTENR][1] + "," +row.LSLANR if row.LSLANR not in mat_ava[row.LSTENR][1] else mat_ava[row.LSTENR][1]]
                    print(mat_ava)
            except pyodbc.Error as e:
                QMessageBox.warning(self, "Database error XPPS", str(e))

    def apply_styles(self):
        self.setStyleSheet("""
        QSpinBox {
                background-color: white;
                color: #d40000;
                border: 2px solid #d40000;
                border-radius: 5px;
                padding: 5px;
                font-size: 16px;
                min-width: 100px;
            }
        QWidget {
            background-color: white;
            font-family: Arial;
            font-size: 14px;
        }

        QLabel {
            color: #b00000;
            font-weight: bold;
            text-align: right;
            text-align: right;
        }

        QPushButton {
            background-color: #b00000;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
        }

        QPushButton:hover {
            background-color: #d00000;
        }

        QPushButton:pressed {
            background-color: #900000;
        }

        QLineEdit, QComboBox {
            border: 1px solid #b00000;
            border-radius: 4px;
            padding: 4px;
            background-color: #fff0f0;
        }

        QListWidget {
            border: 1px solid #b00000;
            border-radius: 6px;
            background-color: #fff4f4;
            padding: 4px;
        }

        QListWidget::item {
            padding: 6px;
            border-bottom: 1px solid #ffd6d6;
        }

        QListWidget::item:selected {
            background-color: #ffcccc;
            color: #900000;
            font-weight: bold;
            border: 1px solid #b00000;
        }

        QScrollBar:vertical {
            border: none;
            background: #fff0f0;
            width: 10px;
            margin: 2px 0 2px 0;
        }

        QScrollBar::handle:vertical {
            background: #d00000;
            min-height: 20px;
            border-radius: 5px;
        }

        QScrollBar::handle:vertical:hover {
            background: #b00000;
        }

        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0;
        }
    """)

    def open_settings(self):
        text = self.input_field.text().strip()
        if text == "A-110033B":
            dialog = SettingsDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                dialog.save_settings()
                self.get_settings()
                
        else:
            self.status.setText("Vetem persona te autorizuar")

    def submit(self):
        text = self.input_field.text().strip()
        self.xvk = text
        if text:
            uid = SETTINGS.value("xpps/user", "FILIPI")
            pw = SETTINGS.value("xpps/password", "a110033")

            # Connection string
            connection_string = (
 
        self.con.format(uid,pw)
            
            )

            # SQL Query (using parameterized query to avoid SQL injection)
            sql = self.sql

            try:
                self.list_widget.clear()
                dega= {}
                m_d= {}
                # Connect
                conn = pyodbc.connect(connection_string)
                cursor = conn.cursor()

                # Execute query with parameter
                cursor.execute(sql, ('4441' + text,))

                # Fetch and print results
                rows = cursor.fetchall()
                module_list = []
                for row in rows:
                    module_list.append(row.XVK_MODULE)
                
                connection_string2 = self.con2
                text01 = ', '.join("'"+str(id)+"'" for id in module_list)
                # SQL query with Unicode character Č (U+010D) and formatted IN clause
                sql2 = self.SQL_TEMPLATE_VODICE.format(text01)

                try:
                    # Connect to the database
                    conn2 = pyodbc.connect(connection_string2)
                    cursor2 = conn2.cursor()

                    # Execute the query
                    cursor2.execute(sql2)

                    # Fetch and print results
                    for row in cursor2.fetchall():
                        print(row)
                        if row.TB not in self.tb:
                            self.tb.append(row.TB)
                        if "pre-a" in str(row.Vend_pune):
                            rend = 0
                        else:
                            rend = 1
                        dega[row.Nga_dega] = rend
                        dega[row.Tek_dega] = rend
                        if row.Moduli in m_d:
                            m_d[row.Moduli][row.Nga_dega] = 1
                            m_d[row.Moduli][row.Tek_dega] = 1
                        else:
                            m_d[row.Moduli] = {}
                            m_d[row.Moduli][row.Nga_dega] = 1
                            m_d[row.Moduli][row.Tek_dega] = 1
                    if self.sort:
                        sorted_items_asc = sorted(dega.items(), key=lambda item: item[1],reverse = True)
                        self.to_print_all = [k[0] for k in sorted_items_asc]
                    else:
                        self.to_print_all = sorted(dega.keys())
                    self.deget = m_d
                    for key in sorted(module_list):
                        self.list_widget.addItem(QListWidgetItem(key))
                    self.status.setText(f"U ngarkuar modulet per XVK{str(text)}")
                except pyodbc.Error as e:
                    QMessageBox.warning(self, "Database error Pozicionet", str(e))
                    
                finally:
                    if 'conn2' in locals():
                        conn2.close()
            except pyodbc.Error as e:
                QMessageBox.warning(self, "Database error XPPS", str(e))
                
            finally:
                if 'conn' in locals():
                    conn.close()
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem plotesoni XVK")
            # option = self.dropdown.currentText()
            # QMessageBox.information(self, "Submitted", f"You entered: {text}\nOption: {option}")

    def reset(self):
        self.manual_select = 0
        self.xvk = ""
        self.modul =""
        self.copies_input.setValue(0)
        self.to_print_all = []
        self.module_field.setText("")
        self.deget = {}
        self.input_field.clear()
        self.list_widget.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.resize(450, 650)
    window.show()
    sys.exit(app.exec_())