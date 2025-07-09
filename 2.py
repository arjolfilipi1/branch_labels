from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QMessageBox,QDialog,QFormLayout,QDialogButtonBox,QListWidget,
    QListWidgetItem,QSpinBox,QTextEdit,QScrollArea
)
from PyQt5.QtCore import Qt, QStringListModel, QSettings
from PyQt5 import QtGui
import os,sys
import pyodbc
from PIL import Image, ImageDraw, ImageFont

import io
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image

# Convert mm to pixels
def mm_to_px(mm,dpi):
    return int((mm / 25.4) * dpi)
    
def save_pdf(data_list, output_pdf,nr_copies):
    for i in range(nr_copies):
        print(i,data_list)
    
# Application settings
SETTINGS = QSettings("Forschner", "Label dege")

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        
        # Create a scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        # Create a widget to hold the layout
        content_widget = QWidget()
        layout = QFormLayout(content_widget)
        
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
        
        # dpi
        self.dpi_edit = QLineEdit()
        self.dpi_edit.setPlaceholderText("300")
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(self.dpi_edit)
        
        # sql2
        self.sql2 = """
    SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli, vodiče.BIS AS Tek_dega
    FROM (KABELY 
    INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
    INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
    WHERE KABELY.Forsch_Nr_kabelu IN ({}) AND vodiče.MAT <> 'Wellrohr';
    """
        self.sql2_edit = QTextEdit()
        self.sql2_edit.setPlaceholderText(self.sql2)
        slq2_layout = QHBoxLayout()
        slq2_layout.addWidget(self.sql2_edit)
        
        # Add to form
        layout.addRow("xPPS user:", db_layout)
        layout.addRow("Password:", password_layout)
        layout.addRow("Vendos dpi e printerit:", dpi_layout)
        layout.addRow("SQL e fijeve:", slq2_layout)
        
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
        self.db_path_edit.setText(SETTINGS.value("xpps/user", "FILIPI"))
        self.password_edit.setText(SETTINGS.value("xpps/password", "a110033"))
        self.dpi_edit.setText(SETTINGS.value("app/dpi", "300"))
        self.sql2_edit.setPlainText(SETTINGS.value("komax/sql", self.sql2))
    
    def save_settings(self):
        SETTINGS.setValue("xpps/user", self.db_path_edit.text())
        SETTINGS.setValue("xpps/password", self.password_edit.text())
        SETTINGS.setValue("app/dpi", self.dpi_edit.text())
        SETTINGS.setValue("komax/sql", self.sql2_edit.toPlainText())

class MyApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Printimi i Labelave te degeve Daimler Wörth")
        self.setup_ui()
        self.apply_styles()
        self.manual_select = 0
    def setup_ui(self):
        
        self.SQL_TEMPLATE_VODICE = """
SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli, vodiče.BIS AS Tek_dega
FROM (KABELY 
INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
WHERE KABELY.Forsch_Nr_kabelu IN ({}) AND vodiče.MAT <> 'Wellrohr';
"""
        self.SQL_TEMPLATE_VODICE = SETTINGS.value("komax/sql", self.SQL_TEMPLATE_VODICE)
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
        button_layout2 = QHBoxLayout()

        # Top bar with Settings button
        self.settings_button = QPushButton("Parametra")
        self.settings_button.clicked.connect(self.open_settings)
        top_bar.addStretch()
        top_bar.addWidget(self.settings_button)

        # Input Field
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("34477595")
        form_layout.addWidget(QLabel("Vendos XVK:"))
        form_layout.addWidget(self.input_field)
        
        # Number input (QSpinBox)
        self.copies_input = QSpinBox()
        self.copies_input.setMinimum(1)  # Minimum 1 copy
        self.copies_input.setMaximum(100)  # Maximum 100 copies
        self.copies_input.setValue(1)  # Default value
        
        form_layout.addWidget(QLabel("Nr. i kopjeve:"))
        form_layout.addWidget(self.copies_input)
        
        # module Field
        self.module_field = QLineEdit()
        self.module_field.setPlaceholderText("444000108010")
        form_layout.addWidget(QLabel("Vendos modulin:"))
        form_layout.addWidget(self.module_field)
        
        # Dropdown
        # self.dropdown = QComboBox()
        # self.dropdown.addItems(["Option 1", "Option 2", "Option 3"])
        # form_layout.addWidget(QLabel("Choose an option:"))
        # form_layout.addWidget(self.dropdown)

        # Buttons
        self.submit_button = QPushButton("Kerko modulet")
        self.submit_button.setToolTip("Shfaq listen e moduleve per kete XVK")
        self.submit_button.clicked.connect(self.submit)

        self.reset_button = QPushButton("Reset")
        self.reset_button.setToolTip("Fshij te dhenat e meparshme")
        self.reset_button.clicked.connect(self.reset)
        
        self.ruaj_button = QPushButton("Ruaj PDF")
        self.ruaj_button.setToolTip("Ruaj labelat ne formatin PDF ne te njejtin folder si aplikacioni")
        self.ruaj_button.clicked.connect(self.save_pdf_button)
        
        self.modul_button = QPushButton("Printo per modulin")
        self.modul_button.setToolTip("Printo labela vetem per modulin")
        self.modul_button.clicked.connect(self.save_pdf_modul)
        
        button_layout.addWidget(self.submit_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.ruaj_button)
        button_layout.addWidget(self.modul_button)
        
        self.print_button = QPushButton("Printo")
        self.print_button.setToolTip("Printo labela per XVK")
        self.print_button.clicked.connect(self.print_all)
        
        button_layout2.addWidget(self.print_button)
        
        # one Field
        self.one_field = QLineEdit()
        self.one_field.setPlaceholderText("dega")
        form_layout.addWidget(QLabel("Etikete shtese"))
        form_layout.addWidget(self.one_field)
        self.one_button = QPushButton("Printo nje Etikete")
        self.one_button.setToolTip("Printo vetem nje etikete")
        self.one_button.clicked.connect(self.one)
        form_layout.addWidget(self.one_button)
        # List Widget
        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.module_selected)
        form_layout.addWidget(QLabel("Modulet:"))
        form_layout.addWidget(self.list_widget)
        
        # Assemble layouts
        main_layout.addLayout(top_bar)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        main_layout.addLayout(button_layout2)

        self.setLayout(main_layout)
        
    def save_pdf_button(self):
        if self.xvk and self.to_print_all:
            l = self.to_print_all
            l.append(self.xvk)
            nr_copies =  self.copies_input.value()
            save_pdf(l,self.xvk +'.pdf',nr_copies)
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem plotesoni XVK")
    def one(self):
        text = self.one_field.text().strip()
        l = [text]
        save_pdf(l,'1.pdf',1)
    def save_pdf_modul(self):
        if self.modul and self.deget:
            l = sorted(self.deget[self.modul].keys())
            l.append(self.modul)
            nr_copies =  self.copies_input.value()
            save_pdf(l,self.xvk+"-"+ self.modul +"-" +'.pdf',nr_copies)
        else:
            if self.module_field.text().strip():
                self.dege_modul()
            else:
                QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem Zgjidhni modulin")
    
    def module_selected(self, item):
        self.manual_select += 1
        self.module_field.setText(item.text())
        self.modul = item.text().strip()
        print(f"Clicked item text: {item.text()}")
        
    def dege_modul(self):
        dega ={}
        text01 = "'" + str(self.module_field.text().strip()) + "'"
        # text01 = "'444000801015'"
        if not text01:
            return None
        connection_string2 = (
                    "DSN=KomaxAL_Durres2;"
                    "Driver={SQL Server};"
                    "System=192.168.102.232;"
                    "UID=komax;"
                    "PWD=komax1;"
                )
        # print(text01)
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
                dega[row.Nga_dega] = 1
                dega[row.Tek_dega] = 1
                
                # print(f"Nga_dega: {row.Nga_dega}, Tek_dega: {row.Tek_dega}"
            l = []
            for d in dega.keys():
                l.append(d)
            l.append(str(self.module_field.text().strip()))
            save_pdf(l,str(self.module_field.text().strip()) +'.pdf',nr_copies)
        except pyodbc.Error as e:
            print("Database error:", e)

        finally:
            if 'conn2' in locals():
                conn2.close()
    
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
        dialog = SettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            dialog.save_settings()
    def print_all(self):
        self.submit()
        self.save_pdf_button()
    def submit(self):
        text = self.input_field.text().strip()
        self.xvk = text
        if text:
            uid = SETTINGS.value("xpps/user", "FILIPI")
            pw = SETTINGS.value("xpps/password", "a110033")

            # Connection string
            connection_string = (
                "Driver={IBM i Access ODBC Driver};"
                "System=192.168.100.35;"
                "DBQ=QGPL;"
                f"UID={uid};"
                f"PWD={pw};"
            )

            # SQL Query (using parameterized query to avoid SQL injection)
            sql = """
                SELECT STRU.STKOMP AS XVK_MODULE 
                FROM BIDBD220.STRU STRU 
                WHERE STRU.STWKNR = '000' 
                  AND STRU.STFIRM = '1' 
                  AND STRU.STBGNR = ?
            """

            try:
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
                    # print(f"XVK_MODULE: {row.XVK_MODULE}")
                    module_list.append(row.XVK_MODULE)
                
                connection_string2 = (
                    "DSN=KomaxAL_Durres2;"
                    "Driver={SQL Server};"
                    "System=192.168.102.232;"
                    "UID=komax;"
                    "PWD=komax1;"
                )
                text01 = ', '.join("'"+str(id)+"'" for id in module_list)
                # print(text01)
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
                        dega[row.Nga_dega] = 1
                        dega[row.Tek_dega] = 1
                        if row.Moduli in m_d:
                            m_d[row.Moduli][row.Nga_dega] = 1
                            m_d[row.Moduli][row.Tek_dega] = 1
                        else:
                            m_d[row.Moduli] = {}
                            m_d[row.Moduli][row.Nga_dega] = 1
                            m_d[row.Moduli][row.Tek_dega] = 1
                        # print(f"Nga_dega: {row.Nga_dega}, Tek_dega: {row.Tek_dega}")
                    self.to_print_all = sorted(dega.keys())
                    self.deget = m_d
                    for key in sorted(module_list):
                        self.list_widget.addItem(QListWidgetItem(key))
                except pyodbc.Error as e:
                    print("Database error:", e)

                finally:
                    if 'conn2' in locals():
                        conn2.close()
            except pyodbc.Error as e:
                print("Database error:", e)

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
    window.resize(350, 650)
    window.show()
    sys.exit(app.exec_())