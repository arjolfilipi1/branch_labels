from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QMessageBox,QDialog,QFormLayout,QDialogButtonBox,QListWidget,
    QListWidgetItem
)
from PyQt5.QtCore import Qt, QStringListModel, QSettings
import os,sys
import pyodbc
from PIL import Image, ImageDraw, ImageFont
# from pylibdmtx.pylibdmtx import encode
import io
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image
import qrcode
# Convert mm to pixels
def mm_to_px(mm,dpi):
    return int((mm / 25.4) * dpi)
    
def save_pdf(data_list, output_pdf="labels.pdf"):
    
    # Label dimensions
    label_width = 49 * mm
    label_height = 9 * mm
    
    # Text area dimensions
    text_width = 39 * mm
    text_height = 9 * mm  # Available height for text
    
    # Data matrix dimensions
    dm_width = 10 * mm
    dm_height = 10 * mm
    
    # Create PDF
    c = canvas.Canvas(output_pdf, pagesize=(label_width, label_height))
    
    for index, item in enumerate(data_list):
        # Start new page for each label
        c.setPageSize((label_width, label_height))
        
        # TEXT (Left side)
        font_size = 6
        c.setFont("Helvetica", font_size)
        
        # Adjust font size if needed
        while c.stringWidth(item, "Helvetica", font_size) > text_width - 2*mm and font_size > 3:
            font_size -= 0.5
        
        # Calculate text metrics for vertical centering
        text_width_actual = c.stringWidth(item, "Helvetica", font_size)
        text_height_actual = font_size * 1.2  # Approximate text height
        
        # Calculate horizontal centering for text
        text_x = 3
        
        # Calculate vertical centering - accounts for actual text height
        vertical_offset = (label_height - text_height_actual) / 2
        
        # Draw centered text (both horizontally and vertically)
        c.setFont("Helvetica", font_size)
        c.drawString(text_x, vertical_offset, item)
        
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,  # Adjust this to control the size
            border=4,
        )
        qr.add_data(item.encode('utf-8'))
        qr.make(fit=True)

        # Create PIL Image
        img = qr.make_image(fill_color="black", back_color="white")

        # Prepare image for PDF
        img_buffer = BytesIO()
        img.save(img_buffer, format='PNG', dpi=(300, 300))
        img_buffer.seek(0)

        # BARCODE (Right side) - Centered vertically with text
        dm_x = label_width - dm_width - 1 * mm
        dm_y = (label_height - dm_height) / 2  # Center vertically

        # Add to PDF
        c.drawImage(
            ImageReader(img_buffer),
            dm_x, dm_y,
            width=dm_width,
            height=dm_height,
            preserveAspectRatio=True,
            mask='auto'
        )
        
        c.showPage()
    
    c.save()
    print(f"PDF with perfectly centered labels created: {output_pdf}")
    
# Application settings
SETTINGS = QSettings("Xpps user", "Password")

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        
        layout = QFormLayout(self)
        
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
        
        # Add to form
        layout.addRow("xPPS user:", db_layout)
        layout.addRow("Password:", password_layout)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        
        # Load current settings
        self.load_settings()

    
    def load_settings(self):
        self.db_path_edit.setText(SETTINGS.value("database/path", ""))
        self.password_edit.setText(SETTINGS.value("images/path", ""))
    
    def save_settings(self):
        SETTINGS.setValue("database/path", self.db_path_edit.text())
        SETTINGS.setValue("images/path", self.password_edit.text())

class MyApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Printimi i Labelave te degeve Daimler Wörth")
        self.setup_ui()
        self.apply_styles()

    def setup_ui(self):
        self.xvk = ""
        self.modul =""
        self.to_print_all = []
        self.deget = {}
        # Layouts
        main_layout = QVBoxLayout()
        top_bar = QHBoxLayout()
        form_layout = QVBoxLayout()
        button_layout = QHBoxLayout()

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

        # Dropdown
        # self.dropdown = QComboBox()
        # self.dropdown.addItems(["Option 1", "Option 2", "Option 3"])
        # form_layout.addWidget(QLabel("Choose an option:"))
        # form_layout.addWidget(self.dropdown)

        # Buttons
        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.submit)

        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self.reset)
        
        self.ruaj_button = QPushButton("Ruaj PDF")
        self.ruaj_button.clicked.connect(self.save_pdf_button)
        
        self.modul_button = QPushButton("Krijo PDF per modulin")
        self.modul_button.clicked.connect(self.save_pdf_modul)
        
        button_layout.addWidget(self.submit_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.ruaj_button)
        button_layout.addWidget(self.modul_button)
        
        # List Widget
        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.module_selected)
        form_layout.addWidget(QLabel("Modulet:"))
        form_layout.addWidget(self.list_widget)
        
        # Assemble layouts
        main_layout.addLayout(top_bar)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)
        
    def save_pdf_button(self):
        if self.xvk and self.to_print_all:
            save_pdf(self.to_print_all,self.xvk +'.pdf')
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem plotesoni XVK")
    
    def save_pdf_modul(self):
        if self.modul and self.deget:
            save_pdf(sorted(self.deget[self.modul].keys()),self.xvk+"-"+ self.modul +"-" +'.pdf')
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem Zgjidhni modulin")
    
    def module_selected(self, item):
        self.modul = item.text()
        print(f"Clicked item text: {item.text()}")
        
    def apply_styles(self):
        self.setStyleSheet("""
        QWidget {
            background-color: white;
            font-family: Arial;
            font-size: 14px;
        }

        QLabel {
            color: #b00000;
            font-weight: bold;
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

    def submit(self):
        text = self.input_field.text().strip()
        self.xvk = text
        if text:
            uid = SETTINGS.value("database/path", "")
            pw = SETTINGS.value("images/path", "")

            # Connection string
            connection_string = (
                "Driver={IBM i Access ODBC Driver};"
                "System=192.168.100.35;"
                f"UID={uid};"
                f"PWD={pw};"
                "DBQ=QGPL;"
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
                sql2 = f"""
                    SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli, vodiče.BIS AS Tek_dega
                    FROM (KABELY 
                    INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
                    INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
                    WHERE KABELY.Forsch_Nr_kabelu IN ({text01}) AND vodiče.MAT <> 'Wellrohr';
                """

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
        self.xvk = ""
        self.modul =""
        self.to_print_all = []
        self.deget = {}
        self.input_field.clear()
        self.list_widget.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.resize(350, 650)
    window.show()
    sys.exit(app.exec_())