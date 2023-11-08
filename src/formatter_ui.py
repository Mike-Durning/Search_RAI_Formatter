from PyQt6.QtWidgets import QTabWidget, QMainWindow, QWidget, QVBoxLayout, QGridLayout, QInputDialog  # noqa: E501
from PyQt6.QtWidgets import QApplication, QComboBox, QPushButton, QCheckBox, QTextEdit, QHBoxLayout  # noqa: E501
#from PyQt6.QtGui  import QIcon
import sys
from set_up import Config
from excel_macro import excel_macro
from excel_manipulation import ExcelManipulator


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # set_up.py 
        self.file_config = Config()
        self.file_config.instantiate_dicts()
        self.file_config.folders_exist()
        self.file_config.save_file_path_to_json()
        self.file_config.save_client_list_to_json()
        self.file_config.save_settings_to_json()
        self.file_path_dict, self.file_path_error = self.file_config.read_json_file(self.file_config.path_data["config_file_path_json"])
        self.client_list_dict, self.client_list_error = self.file_config.read_json_file(self.file_config.path_data["client_list_json"])
        self.settings_dict, self.settings_error = self.file_config.read_json_file(self.file_config.path_data["settings_json"])
        
        # excel_manipulation.py
        self.excel_manipulator = ExcelManipulator()
                
        self.setWindowTitle("RAI Formatter")
        self.setGeometry(175, 25, 1000, 700)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # Layouts
        main_layout = QVBoxLayout() # Changed to a vertical layout for the entire window
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget) # Add the tab widget to the main layout

        # Create the first tab
        first_tab = QWidget()
        tab_widget.addTab(first_tab, "Action Center") # Add the first tab to the tab widget  # noqa: E501

        button_layout = QGridLayout() # Grid layout for buttons
        
        # Create a QTextEdit widget for the first output
        self.output_text1 = QTextEdit()
        self.output_text1.setReadOnly(True)

        # Create a QTextEdit widget for the second output
        self.output_text2 = QTextEdit()
        self.output_text2.setReadOnly(True)

        # Add QTextEdit widgets directly to the layout
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_text1, 2)  # 2/3 of horizontal space
        output_layout.addWidget(self.output_text2, 1)  # 1/3 of horizontal space

        # Add layouts and widgets to the first tab
        first_tab_layout = QVBoxLayout()
        first_tab_layout.addLayout(button_layout)
        first_tab_layout.addLayout(output_layout)  # Add the horizontal output layout
        first_tab.setLayout(first_tab_layout)  # Set the layout for the first tab

        # Create a container widget for the layouts
        container_widget = QWidget()
        container_widget.setLayout(main_layout)
        self.setCentralWidget(container_widget)

        # Apply dark mode stylesheet
        self.setStyleSheet("""
            QMainWindow {
                background-color: #333;
                color: #FFF;
            }
                  
            QPushButton, QRadioButton, QCheckBox {
                background-color: #444;
                color: #FFF;
                text-align: center;
                border: 1px solid #555;
                padding: 5px;
            }
            
            QComboBox, QComboBox QAbstractItemView::item {
                background-color: #444;
                text-align: center;
                color: #FFF;
            }
            
            QFrame {
                background-color: #444;
                border: 1px solid #555;
            }
            QLineEdit {
                background-color: #555;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
            }
            QTextEdit {
                background-color: #555;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
            }
            QLabel {
                color: #FFF;
            }
            
            QTabWidget {
                background-color: #444;
            }
       
        """)
   
        dropdown_position = {
            # First button, dropdown menu
            "Select Client"           : (0, 0)
        }
        
        add_remove_client_position = {
            # Middle Left Button Row - ADD REMOVE CLIENT
            "Add Client"              : (1, 0),
            "Remove Client"           : (1, 1),
        }
        
        button_positions = {
            
            # Top Button Row | Excluding first
            "Excel Manipulation"      : (0, 1),
            "Excel Macro"             : (0, 2),
           
            
            # Middle Right Button Row
            "Clear Output"            : (1, 2),
            
            # Bottom Toggle Row
            "Toggle Short Search List": (2, 0),
            "Toggle Department Column": (2, 1),
            "Toggle Custom Directory" : (2, 2)
        }

        for text, (row, col) in add_remove_client_position.items():
            button = QPushButton(text)
            button.clicked.connect(lambda _, text=text: self.on_button_click(text))
            button_layout.addWidget(button, row, col)
                    
        for text, (row, col) in dropdown_position.items():
            self.drop_menu = QComboBox()
            button_layout.addWidget(self.drop_menu, row, col)
            if text == "Select Client":
                    self.drop_menu.activated.connect(self.on_drop_menu_execute)

        for text, (row, col) in button_positions.items():
            if text.startswith("Toggle"):
                button = QPushButton(text)
                toggle_button = QCheckBox(text)
                button_layout.addWidget(button, row, col)
                button_layout.addWidget(toggle_button, row, col)
                
                if text == "Toggle Short Search List":
                    toggle_button.stateChanged.connect(self.is_ssl)
                elif text == "Toggle Department Column":
                    toggle_button.stateChanged.connect(self.has_department)
                elif text == "Toggle Custom Directory":
                    toggle_button.stateChanged.connect(self.is_default_path)
            else:
                button = QPushButton(text)
                button_layout.addWidget(button, row, col)
                
                if text == "Excel Manipulation":
                    button.clicked.connect(self.excel_manip)
                elif text == "Excel Macro":
                    button.clicked.connect(self.on_excel_macro)
                elif text == "Clear Output":
                    button.clicked.connect(self.clear_output)

        self.populate_combo_box()
        
    def populate_combo_box(self):
        self.output_text2.clear()
        client_json = self.file_config.print_json(self.file_config.path_data["client_list_json"])
        if client_json:
            self.output_text2.append(client_json[0])
              
        client_list = client_json[1]
        client_list = client_list.values() 
        self.drop_menu.addItems(client_list)
    
    def clear_repop(self):
        self.output_text2.clear()
    
        value = self.drop_menu.currentText()
        
        if not value:
            self.output_text1.append("No client selected in the dropdown.")
            return
        
        self.output_text1.append(f"Selected Client: {value}")
        
        selected_client = self.file_config.select_client_by_value(value)
        
        if selected_client is not None:
            print(f"\nSelected Client: {selected_client}\n") 
            self.file_config.search_list_format_info["client_name"] = selected_client
            
            for key, value in self.file_config.search_list_format_info.items():
                self.output_text2.append(f"{key}: {value}")  
            
            self.output_text2.append("")
            
            for key, value in self.file_config.toggle_states.items():
                self.output_text2.append(f"{key}: {value}")  
                
        else:
            print("\nInvalid option.\n")
    
    def on_drop_menu_execute(self):
        self.populate_combo_box()  
        self.clear_repop()             
        
    def excel_manip(self):
        self.output_text1.append("Excel Manipulation Selected")
        search_list_format_info = self.file_config.search_list_format_info
        
        if search_list_format_info.get("client_name") == "NO CLIENT SELECTED":
            self.output_text1.append("No Client Selected")
        
        else:
            chartsearch_path = self.file_config.path_data["chartsearch_xlsx"]
            toggled_states = self.file_config.toggle_states
            alphabet = self.file_config.alphabet
            concat_excel_file = self.excel_manipulator.pandas_column_rearrange(chartsearch_path, toggled_states)
            wb = self.excel_manipulator.openpyxl_format_workbook(concat_excel_file, alphabet, toggled_states)
            self.file_config.save_xlsx(wb, self.file_config.search_list_format_info)

    def on_excel_macro(self):
        self.output_text1.append("Excel Macro Selected")
        search_list_format_info = self.file_config.search_list_format_info
        
        if search_list_format_info.get("client_name") == "NO CLIENT SELECTED":
            self.output_text1.append("No Client Selected")
        else:
            excel_macro(search_list_format_info)

    def clear_output(self):
            self.output_text1.clear()
            self.populate_combo_box()
            
    def is_ssl(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Short Search List' is ON")
            self.file_config.toggle_states["toggle_short_search_list"] = True

        else:
            self.output_text1.append("'Toggle Short Search List' is OFF")
            self.file_config.toggle_states["toggle_short_search_list"] = False
        self.clear_repop()

    def has_department(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Department' is ON")
            self.file_config.toggle_states["toggle_department"] = True

        else:
            self.output_text1.append("'Toggle Department' is OFF")
            self.file_config.toggle_states["toggle_department"] = False
        self.clear_repop()

    def is_default_path(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Custom Directory is ON")
            custom_dir = self.file_config.manual_return_path()
            self.file_config.search_list_format_info["custom_directory"] = custom_dir
            self.file_config.toggle_states["toggle_custom_directory"] = True

        else:
            self.output_text1.append("'Toggle Custom Directory is OFF")
            self.file_config.toggle_states["toggle_custom_directory"] = False
            self.file_config.search_list_format_info["custom_directory"] = None
        self.clear_repop()

# FOR ADDING AND REMOVE CLIENT FROM JSON         
    def on_button_click(self, button_text):
        if button_text == "Add Client":
            client_name, ok = QInputDialog.getText(self, "Add Client", "Enter Client's Full Name:")
            if ok:
                self.file_config.add_client(client_name)
                self.output_text1.append(f"Added client: {client_name}")
                self.drop_menu.clear() 
            self.populate_combo_box()
        
        elif button_text == "Remove Client":
            client_index, ok = QInputDialog.getText(self, "Remove Client", "Enter Client's Index:")
            if ok:
                self.file_config.delete_client(client_index)
                self.output_text1.append(f"Removed Client at Index: {client_index}")
                self.drop_menu.clear() 
            self.populate_combo_box()


            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())