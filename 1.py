from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, 
                             QTableWidgetItem, QLabel, QMessageBox, QLineEdit,
                             QHeaderView,QStackedWidget,QComboBox,QSpinBox,QShortcut,
                             QDialog,QScrollArea,QFormLayout,QCheckBox,QDialogButtonBox,
                             QTextEdit, QFileDialog,QDoubleSpinBox)
from PyQt5.QtCore import Qt,QSettings,QLocale
from PyQt5 import QtGui
import os,sys
from PyQt5.QtPrintSupport import QPrinterInfo
from openpyxl import load_workbook
SETTINGS = QSettings("Forschner", "Label VOGELE")
locale = QLocale(QLocale.English)
QLocale.setDefault(locale)



class PrintSettingsDialog(QDialog):
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File",
            "",  # Start in current directory
            "All Files (*);;Text Files (*.txt);;JSON Files (*.json)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setToolTip(file_path)  # Show full path on hover
    def load_file(self):
        try:
            # Load from QSettings
            file_path = SETTINGS.value("file_path", "")
            
            # Update UI
            self.file_path = file_path
            if file_path:
                self.file_label.setText(os.path.basename(file_path))
                self.file_label.setToolTip(file_path)
            else:
                self.file_label.setText("No file selected")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load settings: {str(e)}")
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
        
        # Database path
        self.file_label = QLabel("No file selected")
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self.browse_file)
        sqldb_layout = QHBoxLayout()
        sqldb_layout.addWidget(self.file_label)
        sqldb_layout.addWidget(self.browse_btn)
        
        
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
        self.offset_edit.setMinimum(-1)
        self.offset_edit.setMaximum(100)
        offset_layout = QHBoxLayout()
        offset_layout.addWidget(self.offset_edit)
        
        # horizontal offset
        self.h_off_edit = QDoubleSpinBox()
        self.h_off_edit.setValue(0.0)
        self.h_off_edit.setMinimum(-50)
        self.h_off_edit.setMaximum(50)
        self.h_off_edit.setSingleStep(0.1)
        h_offset_layout = QHBoxLayout()
        h_offset_layout.addWidget(self.h_off_edit)
        
        # vertikal offset
        self.v_off_edit = QDoubleSpinBox()
        self.v_off_edit.setValue(0.5)
        self.v_off_edit.setMinimum(-10)
        self.v_off_edit.setMaximum(10)
        self.v_off_edit.setSingleStep(0.1)
        v_offset_layout = QHBoxLayout()
        v_offset_layout.addWidget(self.v_off_edit)
        
        # Add to form
        layout.addRow("Folder:", sqldb_layout)
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
        layout.addRow("Velizje horizontale (mm):", h_offset_layout)
        layout.addRow("Velizje vertikale (mm):", v_offset_layout)
        
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
        self.h_off_edit.setValue(float(SETTINGS.value("app/h_off", 0.0)))
        self.v_off_edit.setValue(float(SETTINGS.value("app/v_off", 0.0)))
        self.dpi_edit.setValue(int(SETTINGS.value("app/dpi", 300)))
        self.width_edit.setValue(int(SETTINGS.value("print/width", 49)))
        self.min_font_edit.setValue(int(SETTINGS.value("min_font", 6)))
        self.max_font_edit.setValue(int(SETTINGS.value("max_font", 25)))
        self.height_edit.setValue(int(SETTINGS.value("print/height", 10)))
        self.offset_edit.setValue(int(SETTINGS.value("print/offset", 2)))
        self.load_file()
    
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
        SETTINGS.setValue("app/h_off", float(self.h_off_edit.value()))
        SETTINGS.setValue("app/v_off", float(self.v_off_edit.value()))
        SETTINGS.setValue("file_path", self.file_path)
        

class Label:
    def __innit__(self):
        self.harness = ""
        self.f1 = ""
        self.p1 = ""
        self.p2 = ""
        self.p3 = ""
        self.size = ""
        self.order = ""
        self.use = use
    def add(self,harness,f1,p1,p2,p3,size,order,use):
        self.harness = harness
        self.f1 = f1
        self.p1 = p1
        self.p2 = p2
        self.p3 = p3
        self.size = size
        self.order = order
        use = "Braiding" if use == "Brading" else use
        self.use = use
    def __str__(self):
        return self.harness
    def __lt__(self,other):
        return self.order < other.order

def get_rend():
        file_path = SETTINGS.value("file_path", "") or '\\\\dc08\\share\\Av Vogele E\\Termoetiketa'
        if file_path:
            file_path = file_path + "\\"  +"Termoetiketa_VogeleDB.xlsx"
            print(file_path)
            workbook = load_workbook(filename=file_path, read_only=True)
            sheet = workbook.active
            i = 0
            res = []
            for row in sheet:
                if i >0 and row[0].value is not None:
                    lab = Label()
                    lab.add(row[0].value,row[1].value,row[2].value,row[3].value,row[4].value,row[5].value,row[6].value,row[7].value)
                    res.append(lab)
                i +=1
            workbook.close()
            return (res)
# get_rend()
class MainPage(QWidget):
    #40535D
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode printing Vogele")
        self.setGeometry(50, 50, 400, 400)
        self.ret = get_rend()
        self.all_pn = []
        print_shortcut = QShortcut(QtGui.QKeySequence("Ctrl+Shift+P"), self)
        print_shortcut.activated.connect(self.p_setlings)
        self.setup_ui()
    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        # First row - image (centered in its own horizontal layout)
        image_row = QHBoxLayout()
        image_label = QLabel()
        pixmap = QtGui.QPixmap("logo_forschner.png")  # Replace with your image path
        pixmap = pixmap.scaledToHeight(40, Qt.SmoothTransformation)
        image_label.setPixmap(pixmap)
        image_label.setAlignment(Qt.AlignCenter)
        image_row.addWidget(image_label)

        # Second row - label (centered)
        title_row = QHBoxLayout()
        skano_label = QLabel("Printo etiketa Vogele")
        skano_label.setAlignment(Qt.AlignCenter)
        skano_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        image_row.addWidget(skano_label)
        main_layout.addLayout(image_row)
        
        bordered_widget = QWidget()
        bordered_widget.setObjectName("bo")
        
        # Apply border styling to the widget
        bordered_widget.setStyleSheet("""
            #bo {
                border: 2px solid #cccccc;
                border-radius: 5px;
                padding: 5px;
                background-color: #ffffff;
            }
        """)
        third_row_layout = QHBoxLayout(bordered_widget)
        third_row_layout.setSpacing(20)  # Space between columns
        
        
        # Column layout
        column_layout = QVBoxLayout()
        column_layout.setSpacing(15)
        # label vendos pn
        pn_label = QLabel("Vendos Part Number")
        pn_label.setAlignment(Qt.AlignCenter)
        column_layout.addWidget(pn_label)
        # "pn" combo
        self.pn_combo = QComboBox()
        self.pn_combo.setEditable(True)
        column_layout.addWidget(self.pn_combo)
        separati_label = QLabel("nr i kopjeve")
        separati_label.setAlignment(Qt.AlignCenter)
        self.pn_combo.lineEdit().editingFinished.connect(self.on_finish)
        self.pn_combo.currentTextChanged.connect(self.on_finish)
        column_layout.addWidget(separati_label)
        p_label = QLabel("PN i perzgjedhur:")
        p_label.setAlignment(Qt.AlignCenter)
        column_layout.addWidget(p_label)
        self.label = QLabel("")
        self.label.setAlignment(Qt.AlignCenter)
        column_layout.addWidget(self.label)
        self.sasi_edit = QSpinBox()
        self.sasi_edit.setLocale(locale)
        self.sasi_edit.setValue(1)
        self.sasi_edit.setMinimum(1)
        self.sasi_edit.setMaximum(1000000)
        column_layout.addWidget(self.sasi_edit)
        
        third_row_layout.addLayout(column_layout,1)
        s_column_layout = QVBoxLayout()
        s_column_layout.setSpacing(15)
        
        # Buttons
        self.Pluging_button = QPushButton("Printo Pluging")
        self.Pluging_button.clicked.connect(self.print_plug)
        s_column_layout.addWidget(self.Pluging_button)
        self.dege_button = QPushButton("Printo Dege")
        self.dege_button.clicked.connect(self.print_dege)
        s_column_layout.addWidget(self.dege_button)
        self.braiding_button = QPushButton("Printo braiding")
        self.braiding_button.clicked.connect(self.print_braiding)
        s_column_layout.addWidget(self.braiding_button)
        self.all_button = QPushButton("Printo te gjitha")
        self.all_button.clicked.connect(self.print_all)
        s_column_layout.addWidget(self.all_button)
        third_row_layout.addLayout(s_column_layout,1)
        main_layout.addWidget(bordered_widget)
        # self.all_button.setEnabled(False)
        self.dege_button.setEnabled(False)
        self.Pluging_button.setEnabled(False)
        self.braiding_button.setEnabled(False)
        self.load_pn()
        self.appl_style()
    def on_finish(self):
        print("change")
        pn = str(self.pn_combo.currentText())
        self.pn = pn
        self.label.setText(pn)
        if pn not in self.all_pn:
            self.all_button.setEnabled(False)
            self.dege_button.setEnabled(False)
            self.Pluging_button.setEnabled(False)
            self.braiding_button.setEnabled(False)
            return None
        else:
            self.all_button.setEnabled(True)
            # print([x.harness for x in filter( lambda x:x.harness == pn and x.use =="Pluging" ,self.ret )])
            if any(filter( lambda x:x.harness == pn and x.use =="Pluging" ,self.ret )) :
                self.Pluging_button.setEnabled(True)
            else:
                self.Pluging_button.setEnabled(False)
            if any(filter( lambda x:x.harness == pn and x.use =="Dege" ,self.ret )):
                self.dege_button.setEnabled(True)
            else:
                self.dege_button.setEnabled(False)
            if any(filter( lambda x:x.harness == pn and x.use =="Braiding" ,self.ret )):
                self.braiding_button.setEnabled(True)
            else:
                self.braiding_button.setEnabled(False)
        uses = [x.use for x in self.ret]
    def appl_style(self):
        self.setStyleSheet("QHBoxLayout {border: 2px solid #40535D; border-radius: 5px;} QVBoxLayout {border: 2px solid #40535D; border-radius: 5px;} QSpinBox { background-color: white; color: #40535D; border: 2px solid #40535D; border-radius: 5px; padding: 5px; font-size: 12px; min-width: 100px; } QDoubleSpinBox { background-color: white; color: #40535D; border: 2px solid #40535D; border-radius: 5px; padding: 5px; font-size: 12px; min-width: 100px; } QWidget { background-color: white; font-family: Arial; font-size: 14px; } QLabel { color: #40535D; font-weight: bold; text-align: center;  } QPushButton { background-color: #1010aa; color: white; border: none; padding: 3px 6px; border-radius: 4px; }QPushButton:disabled { background-color: #0d0d0d; } QPushButton:hover { background-color: #0000ff; } QPushButton:pressed { background-color: #40535D; } QLineEdit, QComboBox { border: 1px solid #40535D; border-radius: 4px; padding: 4px; background-color: #f0f0ff; text-align: center;} QListWidget { border: 1px solid #40535D; border-radius: 6px; background-color: #fff4f4; padding: 4px; } QListWidget::item { padding: 6px; border-bottom: 1px solid #ffd6d6; } QListWidget::item:selected { background-color: #ffcccc; color: #40535D; font-weight: bold; border: 1px solid #40535D; } QScrollBar:vertical { border: 1; background: #99ff99; width: 10px; margin: 2px 0 2px 0; } QScrollBar::handle:vertical { background: #30405D; min-height: 20px; border-radius: 5px; } QScrollBar::handle:vertical:hover { background: #40535D; } QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }")
    def load_pn(self):
        d ={}
        for l in self.ret:
            d[l.harness] = 1
        self.all_pn = [str(x) for x in d.keys()]
        self.pn_combo.addItems(self.all_pn)
        self.pn_combo.setInsertPolicy(QComboBox.NoInsert)
        completer = self.pn_combo.completer()
        completer.setCompletionMode(completer.PopupCompletion)
        completer.setFilterMode(Qt.MatchContains)
        self.pn_combo.setCurrentIndex(-1)
        self.pn_combo.clearEditText()
    def to_print(self,li):
        print(li)
    def print_all(self):
        if self.pn:
            tp =[[x.p1,x.p2,x.p3] for x in filter( lambda x:x.harness == self.pn ,self.ret )]
            to_print(tp)
    def print_plug(self):
        print("plug")
        pass
    def print_dege(self):
        print("dege")
        pass
    def print_braiding(self):
        print("barid")
        pass
    def p_setlings(self):
        dialog = PrintSettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            dialog.save_settings()
            self.get_settings()
    def get_settings(self):
        self.use_qr = SETTINGS.value("use_qr", False) == 'true'
        self.font = SETTINGS.value("saved_font", "Helvetica")
        self.simple = SETTINGS.value("simple_print", False) == 'true'
        self.dpi = int(SETTINGS.value("app/dpi", 300))
        self.min_font = int(SETTINGS.value("min_font", 6))
        self.max_font = int(SETTINGS.value("max_font", 25))
        self.index_printer = int(SETTINGS.value("index_printer", 0))
        self.label_width = int(SETTINGS.value("print/width", 50))
        self.label_height = int(SETTINGS.value("print/height", 10))
        self.label_offset = int(SETTINGS.value("print/offset", 2))
        self.h_off = float(SETTINGS.value("app/h_off", 0))
        self.v_off = float(SETTINGS.value("app/v_off", 0))
# ----------------------------
# Run the Application
# ----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainPage()
    window.show()
    sys.exit(app.exec_())