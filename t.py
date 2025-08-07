import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, 
                             QTableWidgetItem, QLabel, QMessageBox, QLineEdit,
                             QHeaderView,QStackedWidget)

from PyQt5.QtCore import Qt,QSettings
import sqlite3
from PyQt5.QtGui import QClipboard,QPixmap
import pandas as pd
import pyodbc
# Application settings
SETTINGS = QSettings("Forschner", "Label EVOBUS")

class PartInventoryApp(QWidget):
    def __init__(self, db_name="inventory.db",parent = None):
        super().__init__()
        self.parent = parent
        # Database connection
        self.conn = None
        self.db_name = db_name

        self.sql = """
                 SELECT
    BIMOD220.XBAU.XABNU2,
    BIMOD220.XBAU.XAMERK,
    BIMOD220.XBAU.XACDIL,
    BIMOD220.XBAU.XATENR,
    BIMOD220.XBAU.XACOMM,
    BIMOD220.XBAU.XASTAT,
    BIMOD220.XBAU.XAFIRM,
    BIMOD220.XBAU.XAWKNR
FROM
    BIMOD220.XBAU
WHERE
    (
        (BIMOD220.XBAU.XAMERK) IN ('3H','9T','DH','NT','VT')
        AND ((BIMOD220.XBAU.XABNU2) IN {} )
        AND (CAST(BIMOD220.XBAU.XASTAT AS INTEGER) < 90)
        AND ((BIMOD220.XBAU.XAFIRM) = '1')
        AND ((BIMOD220.XBAU.XAWKNR) = '000')
    )
            """
        self.xpps_con = "Driver=IBM i Access ODBC Driver;System=192.168.100.35;UID=FILIPI;PWD=A110033;DBQ=QGPL;"
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        # Main widget and layout
        self.main_widget = QWidget()
        # self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self)
        self.main_widget.setLayout(self.layout)
        
        # Database name input
        self.db_input_layout = QHBoxLayout()
        self.db_label = QLabel("Database Name:")
        self.db_input = QLineEdit(self.db_name)
        self.db_input_layout.addWidget(self.db_label)
        self.db_input_layout.addWidget(self.db_input)
        self.layout.addLayout(self.db_input_layout)
        
        # Table name input
        self.table_input_layout = QHBoxLayout()
        self.table_label = QLabel("Table Name:")
        self.table_input = QLineEdit("Plan")
        self.table_input_layout.addWidget(self.table_label)
        self.table_input_layout.addWidget(self.table_input)
        self.layout.addLayout(self.table_input_layout)
        
        # Buttons
        self.button_layout = QHBoxLayout()
        
        self.paste_button = QPushButton("Merr nga Excel")
        self.paste_button.clicked.connect(self.paste_from_clipboard)
        self.button_layout.addWidget(self.paste_button)
        
        self.load_button = QPushButton("Cutting plan")
        self.load_button.clicked.connect(self.load_vod)
        self.button_layout.addWidget(self.load_button)
        
        self.add_row_button = QPushButton("Shto rresht")
        self.add_row_button.clicked.connect(self.add_empty_row)
        self.button_layout.addWidget(self.add_row_button)
        
        self.save_button = QPushButton("Update plan")
        self.save_button.clicked.connect(self.save_to_sqlite)
        self.button_layout.addWidget(self.save_button)
        
        self.clear_button = QPushButton("Delete plan")
        self.clear_button.clicked.connect(self.clear_table)
        self.button_layout.addWidget(self.clear_button)
        
        self.layout.addLayout(self.button_layout)
        
        # Table widget
        self.table = QTableWidget()
        self.table.setColumnCount(3)  # Part Number, Quantity, and Delete button columns
        self.table.setRowCount(0)
        self.table.setHorizontalHeaderLabels(["BB Number", "UG", "Action"])
        
        # Set column widths
        self.table.setColumnWidth(0, 300)  # Part Number
        self.table.setColumnWidth(1, 100)  # Quantity
        self.table.setColumnWidth(2, 80)   # Delete button
        
        # Allow editing for part number and quantity columns
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)
        
        # Connect cell changed signal
        self.table.cellChanged.connect(self.on_cell_changed)
        
        # Track changes to prevent recursive signals
        self._processing_cell_change = False
        # Make the header resizeable
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        
        self.layout.addWidget(self.table)
        layout.addWidget(self.main_widget)
        # Status bar
        self.parent.statusBar().showMessage("Ready")
        
        # Initialize database
        self.init_db()
        self.refresh_table()
    def refresh_table(self):
        self.clear_table(add = False,delete = False)
        table_name = self.table_input.text().strip()
        if not table_name:
            QMessageBox.warning(self, "Invalid Table Name", "Please enter a valid table name.")
            return
        # self.add_empty_row()  # Start with one empty row
        cursor = self.conn.cursor()
        cursor.execute(f"SELECT part_number,ug FROM {table_name};")
        for row in cursor.fetchall():
            self.table.cellChanged.disconnect(self.on_cell_changed)
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            # Add empty items
            # print(row)
            part_num_item = QTableWidgetItem(row[0])
            qty_item = QTableWidgetItem(row[1])
            self.table.setItem(row_position, 0, part_num_item)
            self.table.setItem(row_position, 1, qty_item)
            # Add delete button
            self.add_delete_button(row_position)
            self.table.cellChanged.connect(self.on_cell_changed)
    def on_cell_changed(self, row, column):
        """Handle changes to table cells"""
        if self._processing_cell_change:
            return
            
        self._processing_cell_change = True
        
        try:
            item = self.table.item(row, column)
            if item:
                new_value = item.text()
                
                if column == 0:  # Part Number column
                    print(f"Part number changed in row {row}: {new_value}")
                    # Add your part number validation logic here
                    if '\n' in new_value:
                        self.delete_row(row)
                        self.paste_from_clipboard(new_value)
                elif column == 1:  # Quantity column
                    print(f"Quantity changed in row {row}: {new_value}")
                    # Validate quantity is a number
                    
                
                # You could add auto-save functionality here if desired
                # self.save_to_sqlite()
                
                self.parent.statusBar().showMessage(f"Perditesim i rreshtit {row + 1}")
                
        finally:
            self._processing_cell_change = False
    def init_db(self):
        """Initialize or connect to SQLite database"""
        try:
            self.db_name = self.db_input.text() or "inventory.db"
            self.conn = sqlite3.connect(self.db_name)
            self.parent.statusBar().showMessage(f"Connected to database: {self.db_name}")
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Nuk mundet te lidhet me databasen:\n{str(e)}")
    
    def paste_from_clipboard(self,text = None):
        """Paste data from clipboard into the table widget"""
        # Disconnect the cellChanged signal to prevent multiple triggers
        self.table.cellChanged.disconnect(self.on_cell_changed)
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text()
        if text:
            clipboard_text = text
        if not clipboard_text.strip():
            QMessageBox.warning(self, "Empty Clipboard", "Ju lutem kopjoni te dhenat nga excel.")
            return
        
        try:
            # Split clipboard text into lines
            lines = clipboard_text.split('\n')
            
            # Filter out empty lines and clean whitespace
            part_numbers = [line.strip() for line in lines if line.strip()]
            
            if not part_numbers:
                QMessageBox.warning(self, "No Data", "Nuk jan kopjuar te dhena te sakta.")
                return
            
            # Get current row count
            current_rows = self.table.rowCount() if not text else self.table.rowCount() - 1 
            
            # Add new rows for each part number
            for i, part_num in enumerate(part_numbers):
                row_position = current_rows + i
                self.table.insertRow(row_position)
                
                # Set part number
                part_num_item = QTableWidgetItem(part_num)
                self.table.setItem(row_position, 0, part_num_item)
                
                qty_item = QTableWidgetItem("")
                self.table.setItem(row_position, 1, qty_item)
                
                # Add delete button
                self.add_delete_button(row_position)
            
            self.parent.statusBar().showMessage(f"U shtuan {len(part_numbers)} BB nr")
            
        except Exception as e:
            QMessageBox.critical(self, "Paste Error", f"Nuk mundet te lexohen te dhenat:\n{str(e)}")
        finally:
            self.table.cellChanged.connect(self.on_cell_changed)
    def add_empty_row(self):
        """Add an empty row at the end of the table"""
        self.table.cellChanged.disconnect(self.on_cell_changed)
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        
        # Add empty items
        part_num_item = QTableWidgetItem("")
        qty_item = QTableWidgetItem("")
        
        self.table.setItem(row_position, 0, part_num_item)
        self.table.setItem(row_position, 1, qty_item)
        
        # Add delete button
        self.add_delete_button(row_position)
        
        # Scroll to the new row
        self.table.scrollToBottom()
        self.table.cellChanged.connect(self.on_cell_changed)
    def add_delete_button(self, row):
        """Add a delete button to the specified row"""
        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(lambda: self.delete_row(row))
        self.table.setCellWidget(row, 2, delete_button)
    
    def delete_row(self, row):
        """Delete the specified row from the table"""
        self.table.removeRow(row)
        
        # Update the delete buttons' connections for remaining rows
        for r in range(row, self.table.rowCount()):
            button = self.table.cellWidget(r, 2)
            if button:
                button.clicked.disconnect()
                button.clicked.connect(lambda _, r=r: self.delete_row(r))
        try:
            # Reconnect if needed
            if not self.conn:
                self.init_db()
            table_name = self.table_input.text().strip()
            cursor = self.conn.cursor()
            part_item = self.table.item(row, 0)
            qty_item = self.table.item(row, 1)
            cursor.execute(f"DELETE FROM {table_name} WHERE part_number ='{part_item}' AND ug = '{qty_item}'")
            self.conn.commit()
            self.parent.statusBar().showMessage(f"Deleted row {row + 1}")
        except Exception as e:
                self.parent.statusBar().showMessage("XPPS", f"Could not save data to database:\n{str(e)}")
    
    def load_vod(self):
        try:
            # Reconnect if needed
            if not self.conn:
                self.init_db()
            
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT XABNU2,XAMERK,XACDIL,XATENR AS Part_number,XACOMM FROM BB_query;")
            bb_data = cursor.fetchall()
            if not bb_data:
                self.parent.statusBar().showMessage("No BB_query data")
                return None
            old_plan =[ "'" + str(r[3]) + "'" for r in bb_data ][:-1]
            old_plan =   ",".join(str(item) for item in old_plan) 
            vod_sql = (self.parent.SQL_TEMPLATE_VODICE.format(old_plan))
            vod_con = pyodbc.connect(self.parent.con2)
            vod_cursor = vod_con.cursor()
            vod_cursor.execute(vod_sql)
            v_data = {}
            cursor.execute("DELETE FROM ldata")
            for row in vod_cursor.fetchall():
                part_num = row.Moduli.strip()
                if part_num in v_data:
                    v_data[part_num].append(row)
                else:
                    v_data[part_num]=[row]
            
            for r in bb_data:
                if r[3].strip() in v_data:
                    v = v_data[r[3].strip()]
                    for vr in v:
                        cursor.execute(f"INSERT OR IGNORE INTO ldata (XABNU2,XAMERK,XACDIL,XATENR,XACOMM,VON,UniqueID,BB_UG,LabelID,výkresB0,Modul) VALUES ('{r[0]}', '{r[1]}' ,'{r[2]}' ,'{r[3]}' ,'{r[4]}' ,'{vr[3]}' ,'{r[0].strip()}{r[1].strip()}{vr[3].strip()}' ,'{r[0].strip()}{r[1].strip()}','{vr[4].strip()}{r[1].strip()}{vr[3].strip()}','{vr[4]}','{vr[1]}')")
                        cursor.execute(f"INSERT OR IGNORE INTO ldata (XABNU2,XAMERK,XACDIL,XATENR,XACOMM,VON,UniqueID,BB_UG,LabelID,výkresB0,Modul) VALUES ('{r[0]}', '{r[1]}' ,'{r[2]}' ,'{r[3]}' ,'{r[4]}' ,'{vr[0]}' ,'{r[0].strip()}{r[1].strip()}{vr[0].strip()}' ,'{r[0].strip()}{r[1].strip()}','{vr[4].strip()}{r[1].strip()}{vr[0].strip()}','{vr[4]}','{vr[1]}')")
            self.conn.commit()
            self.parent.statusBar().showMessage(f"Saved label data for {str(len(bb_data))} parts to table ")
        except Exception as e:
            self.conn.rollback()
            QMessageBox.critical(self, "Save Error", f"Could not save data to database:\n{str(e)}")
    def save_to_sqlite(self):
        """Save table data to SQLite database"""
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "No data to save. Please add parts first.")
            return
        
        table_name = self.table_input.text().strip()
        if not table_name:
            QMessageBox.warning(self, "Invalid Table Name", "Please enter a valid table name.")
            return
        
        try:
            # Reconnect if needed
            if not self.conn:
                self.init_db()
            
            cursor = self.conn.cursor()
            
            # Create table (drop if exists)
            # cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
            
            # Create new table with part_number and ug columns
            # cursor.execute(f"""
                # CREATE TABLE {table_name} (
                    # id INTEGER PRIMARY KEY AUTOINCREMENT,
                    # part_number TEXT,
                    # ug TEXT
                # )
# """)
            
            # Insert data (skip the delete button column)
            added_parts = 0
            bbs = []
            cursor.execute(f"SELECT part_number,ug FROM {table_name};")
            old_plan = [ r[0] for r in cursor.fetchall() ]
            
            for row in range(self.table.rowCount()):
                part_item = self.table.item(row, 0)
                qty_item = self.table.item(row, 1)
                
                part_number = part_item.text().strip() if part_item else ""
                quantity = qty_item.text().strip() if qty_item else ""
                bbs.append(part_number)
                
                if not part_number or part_number in old_plan:
                    continue  # Skip empty part numbers
                
                
                cursor.execute(
                    f"INSERT INTO {table_name} (part_number, ug) VALUES ('{part_number}', '{quantity}')"
                )
                added_parts += 1
            
            self.conn.commit()
            self.parent.statusBar().showMessage(f"Saved {added_parts} parts to table '{table_name}' in {self.db_name}")
            connection_string = self.xpps_con
            sql = self.sql
            try:
                if bbs:
                    bbs_str = "("
                    for b in bbs:
                        bbs_str += "'"+b+"',"
                bbs_str = bbs_str[:-1] + ")"
                cursor.execute("DELETE FROM BB_query")
                xpps_conn = pyodbc.connect(connection_string)
                xpps_cursor = xpps_conn.cursor()
                sql = sql.format(bbs_str)
                print(sql)
                xpps_cursor.execute(sql)
                for r in (xpps_cursor.fetchall()):
                    cursor.execute(f"INSERT OR IGNORE INTO BB_query (XABNU2,XAMERK,XACDIL,XATENR,XACOMM,XASTAT,XAFIRM,XAWKNR) VALUES ('{r[0]}', '{r[1]}' ,'{r[2]}' ,'{r[3]}' ,'{r[4]}' ,'{r[5]}' ,'{r[6]}' ,'{r[7]}')")
                    print(f"INSERT OR IGNORE INTO BB_query (XABNU2,XAMERK,XACDIL,XATENR,XACOMM,XASTAT,XAFIRM,XAWKNR) VALUES ('{r[0]}', '{r[1]}' ,'{r[2]}' ,'{r[3]}' ,'{r[4]}' ,'{r[5]}' ,'{r[6]}' ,'{r[7]}')")
                self.conn.commit()
                self.parent.statusBar().showMessage("Successfully saved " + str(added_parts) + " parts to database")
                load_vod()
            except Exception as e:
                print(str(e))
                QMessageBox.critical(self, "XPPS", f"Could not save data to database BB_query:\n{str(e)}")
                
            
            
        except Exception as e:
            self.conn.rollback()
            QMessageBox.critical(self, "Save Error", f"Could not save data to database:\n{str(e)}")
    
    def clear_table(self,add=True,delete = True):
        """Clear the table widget"""
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Part Number", "Quantity", "Action"])
        self.parent.statusBar().showMessage("Table cleared")
        if add:
            self.add_empty_row()  # Add one empty row back
        if delete:
            try:
                # Reconnect if needed
                if not self.conn:
                    self.init_db()
                
                cursor = self.conn.cursor()
                cursor.execute("DELETE FROM Plan")
                self.conn.commit()
                self.parent.statusBar().showMessage("Plani u fshi")
            except Exception as e:
                print(str(e))
                QMessageBox.critical(self, "Plan", f"Could not delete database Plan:\n{str(e)}")
    def closeEvent(self, event):
        """Clean up when closing the application"""
        if self.conn:
            self.conn.close()
        event.accept()

# ----------------------------
# Main Page (simple welcome screen)
# ----------------------------
class MainPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        # Main vertical layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # First row - image (centered in its own horizontal layout)
        image_row = QHBoxLayout()
        image_label = QLabel()
        pixmap = QPixmap("logo_forschner.png")  # Replace with your image path
        image_label.setPixmap(pixmap)
        image_label.setAlignment(Qt.AlignCenter)
        image_row.addWidget(image_label)
        main_layout.addLayout(image_row)

        # Second row - "Skano BB" label (centered)
        title_row = QHBoxLayout()
        skano_label = QLabel("Skano BB")
        skano_label.setAlignment(Qt.AlignCenter)
        skano_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        title_row.addWidget(skano_label)
        main_layout.addLayout(title_row)

        # Third row - info label (centered)
        info_row = QHBoxLayout()
        info_label = QLabel("BB Number:")
        info_label.setAlignment(Qt.AlignCenter)
        info_row.addWidget(info_label)
        main_layout.addLayout(info_row)

        # Fourth row - text input with stretch on both sides
        input_row = QHBoxLayout()
        input_row.addStretch()  # Add stretch before
        text_input = QLineEdit()
        text_input.setPlaceholderText("Type here...")
        text_input.setFixedWidth(300)  # Set a reasonable width
        input_row.addWidget(text_input)
        input_row.addStretch()  # Add stretch after
        main_layout.addLayout(input_row)

        # Fifth row - button (centered)
        button_row = QHBoxLayout()
        button_row.addStretch()
        button = QPushButton("Submit")
        button.setStyleSheet("""
            QPushButton {
                padding: 8px;
                font-size: 16px;
                min-width: 100px;
            }
        """)
        button_row.addWidget(button)
        button_row.addStretch()
        main_layout.addLayout(button_row)

# ----------------------------
# Main Application Window
# ----------------------------
class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Page App")
        self.setGeometry(100, 100, 800, 600)
        self.setup_ui()

    def setup_ui(self):
        self.SQL_TEMPLATE_VODICE = """
SELECT vodiče.VON AS Nga_dega, KABELY.Forsch_Nr_kabelu AS Moduli,vodiče.MODUL, vodiče.BIS AS Tek_dega,KABELY.výkresB0 AS B0
FROM (KABELY 
INNER JOIN [KABELY NA POZICE] ON KABELY.Forsch_Nr_kabelu = [KABELY NA POZICE].Forsch_Nr_kabelu) 
INNER JOIN vodiče ON [KABELY NA POZICE].Pozice = vodiče.POS
WHERE KABELY.Forsch_Nr_kabelu IN ({});
"""
        self.con2 = "DSN=KomaxAL_Durres2;Driver={SQL Server};System=192.168.102.232;UID=komax;PWD=komax1;"
        # Central Widget and Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Sidebar Navigation
        sidebar = QWidget()
        sidebar.setFixedWidth(150)
        sidebar_layout = QVBoxLayout(sidebar)
        
        self.btn_main = QPushButton("Main Page")
        self.btn_inventory = QPushButton("Inventory")
        
        sidebar_layout.addWidget(self.btn_main)
        sidebar_layout.addWidget(self.btn_inventory)
        sidebar_layout.addStretch()

        # Page Container (Stacked Widget)
        self.pages = QStackedWidget()
        self.pages.addWidget(MainPage())                   # Index 0
        self.p_two = PartInventoryApp("inventory.db",parent = self)
        self.pages.addWidget(self.p_two)  # Index 1

        # Add to main layout
        main_layout.addWidget(sidebar)
        main_layout.addWidget(self.pages)

        # Connect navigation buttons
        self.btn_main.clicked.connect(lambda: self.pages.setCurrentIndex(0))
        self.btn_inventory.clicked.connect(self.page_two)
    def page_two(self):
        self.pages.setCurrentIndex(1)
        self.p_two.refresh_table()
# ----------------------------
# Run the Application
# ----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())