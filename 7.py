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
from pylibdmtx.pylibdmtx import encode
from reportlab.lib.utils import ImageReader
from io import BytesIO
import tempfile,time
import qrcode
import win32print,win32ui,win32api
# Convert mm to pixels
import fitz  # PyMuPDF


locale = QLocale(QLocale.English)
QLocale.setDefault(locale)

    
# Application settings
SETTINGS = QSettings("Forschner", "Label dege")

class PrintSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Print Settings")
        self.resize(430, 600)
        # Create a scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        # Create a widget to hold the layout
        content_widget = QWidget()
        layout = QFormLayout(content_widget)
        
        self.font_combo = QComboBox()
        font_layout = QHBoxLayout()
        font_layout.addWidget(self.font_combo)
        
        self.printer_combo = QComboBox()
        pc_layout = QHBoxLayout()
        pc_layout.addWidget(self.printer_combo)
        
        # use qr
        self.checkbox = QCheckBox("Use qr", self)
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.checkbox)
        
        # use qr
        self.s_print = QCheckBox("Use simple print", self)
        s_print_layout = QHBoxLayout()
        s_print_layout.addWidget(self.s_print)
        
        # dpi
        self.dpi_edit = QSpinBox()
        self.dpi_edit.setValue(300)
        self.dpi_edit.setMinimum(1)
        self.dpi_edit.setMaximum(1000000)
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(self.dpi_edit)
        
        # min_font
        self.min_font_edit = QSpinBox()
        self.min_font_edit.setValue(6)
        self.min_font_edit.setMinimum(3)
        self.min_font_edit.setMaximum(25)
        min_font_layout = QHBoxLayout()
        min_font_layout.addWidget(self.min_font_edit)
        
        # max_font
        self.max_font_edit = QSpinBox()
        self.max_font_edit.setValue(25)
        self.max_font_edit.setMinimum(3)
        self.max_font_edit.setMaximum(25)
        max_font_layout = QHBoxLayout()
        max_font_layout.addWidget(self.max_font_edit)
        
        # width
        self.width_edit = QSpinBox()
        self.width_edit.setValue(50)
        self.width_edit.setMinimum(1)
        self.width_edit.setMaximum(1000)
        width_layout = QHBoxLayout()
        width_layout.addWidget(self.width_edit)
        
        # height
        self.height_edit = QSpinBox()
        # self.height_edit.setPlaceholderText("10")
        self.height_edit.setValue(10)
        self.height_edit.setMinimum(1)
        self.height_edit.setMaximum(1000)
        height_layout = QHBoxLayout()
        height_layout.addWidget(self.height_edit)
        
        # offset
        self.offset_edit = QSpinBox()
        # self.offset_edit.setPlaceholderText("2")
        self.offset_edit.setValue(2)
        self.offset_edit.setMinimum(1)
        self.offset_edit.setMaximum(100)
        offset_layout = QHBoxLayout()
        offset_layout.addWidget(self.offset_edit)
        
        # Add to form
        layout.addRow("Font:", font_layout)
        layout.addRow("Zgjidh printerin:", pc_layout)
        layout.addRow("Perdor qr code:", checkbox_layout)
        layout.addRow("Perdor Printim automatik:", s_print_layout)
        layout.addRow("dpi e printerit:", dpi_layout)
        layout.addRow("gjeresi e etiketes (mm):", width_layout)
        layout.addRow("madhesia min e textit:", min_font_layout)
        layout.addRow("madhesia max e textit:", max_font_layout)
        layout.addRow("lartesia e etiketes (mm):", height_layout)
        layout.addRow("Fillimi i shkrimit (mm):", offset_layout)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        # layout.addRow(buttons)
        
        # Set the content widget to scroll area
        scroll.setWidget(content_widget)
        
        # Set scroll area as the main layout
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll)
        main_layout.addWidget(buttons)
        self.populate_printers()
        self.populate_fonts()
        # Load current settings
        self.load_settings()
    def populate_fonts(self):
        standard_fonts = [
    'Courier', 'Courier-Bold', 'Courier-Oblique', 'Courier-BoldOblique',
    'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique',
    'Times-Roman', 'Times-Bold', 'Times-Italic', 'Times-BoldItalic',
    'Symbol', 'ZapfDingbats']       
        for font in standard_fonts:
            self.font_combo.addItem(font)
    def populate_printers(self):
        # Get all available printers
        printers = QPrinterInfo.availablePrinters()
        
        # Add printer names to combo box
        for printer in printers:
            self.printer_combo.addItem(printer.printerName())
            
        
            
    def load_settings(self):
        # Load saved printer from settings
        saved_printer = SETTINGS.value("saved_printer", "")
        try:
            if saved_printer:
            # Find the index of the saved printer
                index = self.printer_combo.findText(saved_printer)
                if index >= 0:
                    self.printer_combo.setCurrentIndex(index)
                    
        except:
            pass
        saved_font = SETTINGS.value("saved_font", "Courier-Bold")
        try:
            if saved_font:
            # Find the index of the saved printer
                index = self.font_combo.findText(saved_font)
                if index >= 0:
                    self.font_combo.setCurrentIndex(index)
                    
        except:
            pass
        if SETTINGS.value("simple_print", False) != 'false':
            self.s_print.setChecked(True)
        if SETTINGS.value("use_qr", False) != 'false':
            self.checkbox.setChecked(True)
        self.dpi_edit.setValue(int(SETTINGS.value("app/dpi", 300)))
        self.width_edit.setValue(int(SETTINGS.value("print/width", 49)))
        self.min_font_edit.setValue(int(SETTINGS.value("min_font", 6)))
        self.max_font_edit.setValue(int(SETTINGS.value("max_font", 25)))
        self.height_edit.setValue(int(SETTINGS.value("print/height", 10)))
        self.offset_edit.setValue(int(SETTINGS.value("print/offset", 2)))
    
    def save_settings(self):
        # Get selected printer
        selected_printer = self.printer_combo.currentText()
        selected_font = self.font_combo.currentText()
        index = self.printer_combo.findText(selected_printer)        
        # Save to settings
        SETTINGS.setValue("saved_printer", selected_printer)
        SETTINGS.setValue("saved_font", selected_font)
        SETTINGS.setValue("index_printer", index)
        SETTINGS.setValue("use_qr", self.checkbox.isChecked())
        SETTINGS.setValue("simple_print", self.s_print.isChecked())
        SETTINGS.setValue("app/dpi", str(self.dpi_edit.value()))
        SETTINGS.setValue("min_font", str(self.min_font_edit.value()))
        SETTINGS.setValue("max_font", str(self.max_font_edit.value()))
        SETTINGS.setValue("print/width", str(self.width_edit.value()))
        SETTINGS.setValue("print/height", str(self.height_edit.value()))
        SETTINGS.setValue("print/offset", str(self.offset_edit.value()))

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
        
        # dpi
        self.dpi_edit = QLineEdit()
        self.dpi_edit.setPlaceholderText("300")
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(self.dpi_edit)
        
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
    SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli,vodiče.MODUL, vodiče.BIS AS Tek_dega
    FROM (KABELY 
    INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
    INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
    WHERE KABELY.Forsch_Nr_kabelu IN ({}) AND vodiče.MAT <> 'Wellrohr' ORDER BY vodiče.MODUL;
    """
        self.con2 = "DSN=KomaxAL_Durres2;Driver={SQL Server};System=192.168.102.232;UID=komax;PWD=komax1;"
        
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
        layout.addRow("dpi e printerit:", dpi_layout)
        layout.addRow("SQL e xpps:", slq_layout)
        layout.addRow("Conn e xpps:", con_layout)
        layout.addRow("SQL e fijeve:", slq2_layout)
        layout.addRow("Conn e fijeve:", con2_layout)
        
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
        self.dpi_edit.setText(SETTINGS.value("app/dpi", "300"))
        self.sql_edit.setPlainText(SETTINGS.value("xpps/sql", self.sql))
        self.con_edit.setPlainText(SETTINGS.value("xpps/con", self.con))
        self.sql2_edit.setPlainText(SETTINGS.value("komax/sql", self.sql2))
        self.con2_edit.setPlainText(SETTINGS.value("komax/con", self.con2))
    
    def save_settings(self):
        SETTINGS.setValue("use_qr", self.checkbox.isChecked())
        SETTINGS.setValue("sort", self.check2box.isChecked())
        SETTINGS.setValue("xpps/user", self.db_path_edit.text())
        SETTINGS.setValue("xpps/password", self.password_edit.text())
        SETTINGS.setValue("app/dpi", self.dpi_edit.text())
        SETTINGS.setValue("xpps/sql", self.sql_edit.toPlainText())
        SETTINGS.setValue("xpps/con", self.con_edit.toPlainText())
        SETTINGS.setValue("komax/sql", self.sql2_edit.toPlainText())
        SETTINGS.setValue("komax/con", self.con2_edit.toPlainText())

class MyApp(QWidget):
    def print_pdf_natively(self,pdf_path):
        if self.simple_print:
            default_printer = win32print.GetDefaultPrinter()
            try:
                target_printer = SETTINGS.value("saved_printer", "")
                win32print.SetDefaultPrinter(target_printer)
                win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
            except Exception as e:
                self.status.setText(f"Sending to printing failed: {str(e)}")
            finally:
                win32print.SetDefaultPrinter(default_printer)
        else:
            doc = fitz.open(pdf_path)
            # Convert mm to pixels
            width_px = int((self.label_width / 25.4) * self.dpi)
            height_px = int((self.label_height / 25.4) * self.dpi)
            printer_name = SETTINGS.value("saved_printer", "")
            if not printer_name:
                printer_name = win32print.GetDefaultPrinter()
                
            for i in range(len(doc)):
                try:
                    page = doc[i]
                    pix = page.get_pixmap(dpi=1600)
                    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    self.status.setText(f"Printing page {i + 1}")

                    # Resize image to fixed size
                    resized = image.resize((width_px, height_px), Image.NEAREST)

                    # Create printer DC
                    hdc = win32ui.CreateDC()
                    hdc.CreatePrinterDC(printer_name)

                    # Start printing
                    hdc.StartDoc(f"PDF Page {i+1}")
                    hdc.StartPage()

                    # Draw the image at 0,0 with 50x10mm size (in pixels)
                    dib = ImageWin.Dib(resized)
                    dib.draw(hdc.GetHandleOutput(), (0, 0, width_px, height_px))

                    hdc.EndPage()
                    hdc.EndDoc()
                    hdc.DeleteDC()
                except Exception as e:
                    self.status.setText(f"Sending to printing failed: {str(e)} {i}")
    def print_pdf(self,data_list, output_pdf,nr_copies,use_qr):
        # Label dimensions
        label_width = self.label_width  * mm
        label_height =self.label_height * mm
        
        # Text area dimensions
        text_width = (self.label_width - 10) * mm
        text_height = 9 * mm  # Available height for text
        
        # Data matrix dimensions
        dm_width = 10 * mm
        dm_height = 10 * mm
        # Create temporary PDF file (kept open to prevent deletion)
        temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        temp_pdf_path = temp_pdf.name
        temp_pdf.close()  # Close handle so Sumatra can access it
        try:
            # Create PDF
            c = canvas.Canvas(temp_pdf_path, pagesize=(label_width, label_height))
            temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            temp_pdf.close()
            for i in range(nr_copies):
            
                for index, item in enumerate(data_list):
                    if item == None:
                        continue
                    # Start new page for each label
                    c.setPageSize((label_width, label_height))
                    
                    # TEXT (Left side)
                    font_size = self.max_font
                    c.setFont(SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                    
                    # Adjust font size if needed
                    while c.stringWidth(item, SETTINGS.value("saved_font", "Courier-Bold"), font_size) > ((text_width ) - (self.label_offset*mm*2)) and font_size > self.min_font:
                        font_size -= 0.5
                    
                    # Calculate text metrics for vertical centering
                    text_width_actual = c.stringWidth(item, SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                    text_height_actual = font_size * 0.8  # Approximate text height
                    
                    # Calculate horizontal centering for text
                    text_x = self.label_offset*mm
                    
                    # Calculate vertical centering - accounts for actual text height
                    vertical_offset = (label_height - text_height_actual) / 2
                    
                    # Draw centered text (both horizontally and vertically)
                    c.setFont(SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                    c.drawString(text_x, vertical_offset, item)
                    if use_qr:
                        try:
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
                            img.save(img_buffer, format='PNG', dpi=(self.dpi, self.dpi))
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
                            
                        except Exception as e:
                            self.status.setText(f"Barcode generation failed: {str(e)}")
                    else:
                        try:
                                # Barcode
                            encoded = encode(item.encode('utf-8'))
                            img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
                            img_buffer = BytesIO()
                            img.save(img_buffer, format='PNG')
                            img_buffer.seek(0)
                            
                            c.drawImage(ImageReader(img_buffer), 
                                       49*mm - 10*mm - 1*mm, 
                                       (9*mm - 8*mm)/2, 
                                       width=8*mm, 
                                       height=8*mm,
                                       preserveAspectRatio=True)
                            c.showPage()
                        except Exception as e:
                            self.status.setText(f"Barcode generation failed: {str(e)}")
            c.save()
            
            temp_pdf.close()
            time.sleep(0.5)
            try:
                self.print_pdf_natively(temp_pdf_path)
            # os.startfile(temp_pdf_path, 'print')
            except Exception as e:
                self.status.setText(f"Line 450 Printing failed before: {str(e)}",temp_pdf_path)
        except Exception as e:
            self.status.setText(f"Line 452 Printing failed: {str(e)}")
        finally:
            try:
                os.unlink(temp_pdf.name)
            except:
                pass
    def save_pdf(self,data_list, output_pdf,nr_copies,use_qr):
        # Label dimensions
        label_width = self.label_width  * mm
        label_height =self.label_height * mm
        
        # Text area dimensions
        text_width = (self.label_width - 10) * mm
        text_height = 9 * mm  # Available height for text
        # Data matrix dimensions
        dm_width = 10 * mm
        dm_height = 10 * mm
        
        # Create PDF
        c = canvas.Canvas(output_pdf, pagesize=(label_width, label_height))
        for i in range(nr_copies):
        
            for index, item in enumerate(data_list):
                # Start new page for each label
                c.setPageSize((label_width, label_height))
                
                # TEXT (Left side)
                font_size = self.max_font
                c.setFont(SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                
                # Adjust font size if needed
                
                while c.stringWidth(item, SETTINGS.value("saved_font", "Courier-Bold"), font_size) > ((text_width ) - (self.label_offset*mm*2)) and font_size > self.min_font:
                    font_size -= 0.5
                # Calculate text metrics for vertical centering
                text_width_actual = c.stringWidth(item, SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                text_height_actual = font_size * 0.8   # Approximate text height
                
                # Calculate horizontal centering for text
                text_x = self.label_offset*mm
                
                # Calculate vertical centering - accounts for actual text height
                vertical_offset = (label_height - text_height_actual) / 2
                
                # Draw centered text (both horizontally and vertically)
                c.setFont(SETTINGS.value("saved_font", "Courier-Bold"), font_size)
                c.drawString(text_x, vertical_offset, item)
                
                if use_qr:
                    try:
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
                        img.save(img_buffer, format='PNG', dpi=(self.dpi, self.dpi))
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
                    except Exception as e:
                        self.status.setText(f"Line 536 Printing failed: {str(e)}")
                else:
                    try:
                            # Barcode
                        encoded = encode(item.encode('utf-8'))
                        img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
                        img_buffer = BytesIO()
                        img.save(img_buffer, format='PNG')
                        img_buffer.seek(0)
                        
                        c.drawImage(ImageReader(img_buffer), 
                                   49*mm - 10*mm - 1*mm, 
                                   (9*mm - 8*mm)/2, 
                                   width=dm_width - 1 * mm, 
                                   height=dm_width - 1 * mm,
                                   preserveAspectRatio=True)
                        c.showPage()
                    except Exception as e:
                        self.status.setText(f"Line 554 Printing failed: {str(e)}")
        
        c.save()
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
        self.SQL_TEMPLATE_VODICE = SETTINGS.value("komax/sql", self.SQL_TEMPLATE_VODICE)
        self.con2 = SETTINGS.value("komax/con", self.con2)
        self.use_qr = SETTINGS.value("use_qr", False) == 'true'
        self.dpi = int(SETTINGS.value("app/dpi", 300))
        self.min_font = int(SETTINGS.value("min_font", 6))
        self.max_font = int(SETTINGS.value("max_font", 25))
        self.index_printer = int(SETTINGS.value("index_printer", 0))
        self.label_width = int(SETTINGS.value("print/width", 50))
        self.label_height = int(SETTINGS.value("print/height", 10))
        self.label_offset = int(SETTINGS.value("print/offset", 2))
        self.simple_print = SETTINGS.value("simple_print", False) == 'true'
        self.sort = SETTINGS.value("sort", False) == 'true'
    def setup_ui(self):
        self.min_font = int(SETTINGS.value("min_font", 6))
        self.max_font = int(SETTINGS.value("max_font", 25))
        self.index_printer = int(SETTINGS.value("index_printer", 0))
        self.dpi = int(SETTINGS.value("app/dpi", 300))
        self.label_width = int(SETTINGS.value("print/width", 50))
        self.label_height = int(SETTINGS.value("print/height", 10))
        self.label_offset = int(SETTINGS.value("print/offset", 2))
        self.dpi = int(SETTINGS.value("app/dpi", 300))
        self.use_qr = SETTINGS.value("use_qr", False) == 'true'
        self.sort = SETTINGS.value("sort", False) == 'true'
        self.simple_print = SETTINGS.value("simple_print", False) == 'true'
        
        self.SQL_TEMPLATE_VODICE = """
SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli,vodiče.MODUL, vodiče.BIS AS Tek_dega,KABELY.výkresB0 as TB
FROM (KABELY 
INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
WHERE KABELY.Forsch_Nr_kabelu IN ({}) AND vodiče.MAT <> 'Wellrohr' ORDER BY vodiče.MODUL;
"""
        self.con = "Driver={{IBM i Access ODBC Driver}};System=192.168.100.35;UID={};PWD={};DBQ=QGPL;"
        self.con2 = "DSN=KomaxAL_Durres2;Driver={SQL Server};System=192.168.102.232;UID=komax;PWD=komax1;"
        self.sql = """
                SELECT STRU.STKOMP AS XVK_MODULE 
                FROM BIDBD220.STRU STRU 
                WHERE STRU.STWKNR = '000' 
                  AND STRU.STFIRM = '1' 
                  AND STRU.STBGNR = ?
            """
        self.get_settings()
        
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
        status_layout = QHBoxLayout()
        
        # Top bar with Settings button
        self.psettings_button = QPushButton("Printer")
        icon = QtGui.QIcon('p.png')
        self.psettings_button.setIcon(icon)
        self.psettings_button.clicked.connect(self.open_psettings)
        top_bar.addStretch()
        # top_bar.addWidget(self.psettings_button)
        
        # Top bar with Settings button
        self.settings_button = QPushButton("Parametra")
        icon = QtGui.QIcon('s.png')
        self.settings_button.setIcon(icon)
        self.settings_button.clicked.connect(self.open_settings)
        top_bar.addStretch()
        # top_bar.addWidget(self.settings_button)

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
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.status = QLabel()
        scroll.setWidget(self.status)
        status_layout.addWidget(scroll)
        
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
        main_layout.addLayout(status_layout)

        self.setLayout(main_layout)
        
    def save_pdf_button(self):
        if self.xvk and self.to_print_all:
            l = self.to_print_all
            l.append("XVK" + str(self.xvk))
            nr_copies =  self.copies_input.value()
            self.save_pdf(l,self.xvk +'.pdf',nr_copies,self.use_qr)
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem plotesoni XVK")
    def one(self):
        text = self.one_field.text().strip()
        l = [text]
        self.print_pdf(l,'1.pdf',1,self.use_qr)
    def save_pdf_modul(self):
        if self.modul and self.modul in self.deget:
            l = sorted(self.deget[self.modul].keys())
            l.append(self.modul)
            nr_copies =  self.copies_input.value()
            self.print_pdf(l,self.xvk+"-"+ self.modul +"-" +'.pdf',nr_copies,self.use_qr)
        else:
            if self.module_field.text().strip():
                self.dege_modul()
            else:
                QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem Zgjidhni modulin")
    
    def module_selected(self, item):
        self.manual_select += 1
        self.module_field.setText(item.text())
        self.modul = item.text().strip()
        self.status.setText(f"Perzgjodhet: {item.text()}")
        
    def dege_modul(self):
        nr_copies =  self.copies_input.value()
        # print_pdf(["111"],str(self.module_field.text().strip()) +'.pdf',nr_copies)
        # return None
        dega ={}
        text01 = "'" + str(self.module_field.text().strip()) + "'"
        # text01 = "'444000801015'"
        if not text01:
            return None
        connection_string2 = self.con2
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


            l = []
            for d in dega.keys():
                l.append(d)
            l.append(str(self.module_field.text().strip()))
            self.print_pdf(l,str(self.module_field.text().strip()) +'.pdf',nr_copies,self.use_qr)
        except pyodbc.Error as e:
            QMessageBox.warning(self, "Database error", str(e))
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
    def open_psettings(self):
        text = self.one_field.text().strip()
        if text == "A-12345B":
            dialog = PrintSettingsDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                dialog.save_settings()
                self.get_settings()
        else:
            self.status.setText("Vetem persona te autorizuar")
    def open_settings(self):
        text = self.one_field.text().strip()
        if text == "A-110033B":
            dialog = SettingsDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                dialog.save_settings()
                self.get_settings()
                self.status.setText("Perdorimi i qr code eshte: " + str(self.use_qr))
        else:
            self.status.setText("Vetem persona te autorizuar")
    def print_all(self):
        self.submit()
        if self.xvk and self.to_print_all:
            l = self.to_print_all
            for t in self.tb:
                l.append(t)
            l.append("XVK" + str(self.xvk))
            nr_copies =  self.copies_input.value()
            self.print_pdf(l,self.xvk +'.pdf',nr_copies,self.use_qr)
        else:
            QMessageBox.warning(self, "Mungojne te dhenat", "Ju lutem plotesoni XVK")
        self.tb = []
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
                        if row.TB not in self.tb:
                            self.tb.append(row.TB)
                        dega[row.Nga_dega] = row.MODUL
                        dega[row.Tek_dega] = row.MODUL
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