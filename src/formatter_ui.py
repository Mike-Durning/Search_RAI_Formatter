from PyQt6.QtWidgets import QTabWidget, QMainWindow, QWidget, QVBoxLayout, QGridLayout, QInputDialog  # noqa: E501
from PyQt6.QtWidgets import QApplication, QComboBox, QPushButton, QCheckBox, QTextEdit, QHBoxLayout  # noqa: E501
#from PyQt6.QtGui  import QIcon
import sys
from pathlib import Path
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
            "Toggle Search List"      : (2, 0),
            "Toggle Attempts Column"  : (2, 1),
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
                
                if text == "Toggle Search List":
                    toggle_button.stateChanged.connect(self.is_search_list)
                elif text == "Toggle Attempts Column":
                    toggle_button.stateChanged.connect(self.has_attempt)
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

        self.populate_client_list_menu()
    
    def output_2_update(self): 
        self.output_text2.clear()
        self.search_list_info_formatted_str, self.search_list_info_dict = self.file_config.print_dict_or_json(self.file_config.search_list_format_info)
        self.toggle_states_formatted_str, self.toggle_states_dict = self.file_config.print_dict_or_json(self.file_config.toggle_states)
        self.output_text2.append(self.search_list_info_formatted_str)      
        self.output_text2.append(self.toggle_states_formatted_str)
        return self.search_list_info_dict, self.toggle_states_dict
    
    def final_output_2_update(self): 
        self.output_text2.clear()
        self.excel_dimensions_formatted_str, self.excel_dimensions_dict = self.file_config.print_dict_or_json(self.excel_dimensions)
        self.output_text2.append(self.search_list_info_formatted_str)      
        self.output_text2.append(self.toggle_states_formatted_str)
        self.output_text2.append(self.excel_dimensions_formatted_str)
        return self.search_list_info_dict, self.toggle_states_dict, self.excel_dimensions_dict
        
    def populate_client_list_menu(self): # Inputs client list into Output 2
        
        self.client_list_formatted, client_list_dict, client_list_json_path = self.file_config.print_dict_or_json(Path(self.file_config.path_data["client_list_json"]))
        
        self.output_text2.clear()
        if self.client_list_formatted:
            self.output_text2.append(self.client_list_formatted)
              
        client_dict = client_list_dict
        client_list = client_dict.values() 
        self.drop_menu.addItems(client_list)
    
    def drop_menu_select(self):   
        selected = self.drop_menu.currentText()
        
        if not selected:
            self.output_text1.append("No client selected in the dropdown.")
    
        self.output_text1.append(f"Selected Client: {selected}")
        
        selected_client = self.file_config.select_client_by_value(selected)
        
        if selected_client is not None:
            self.file_config.search_list_format_info["client_name"] = selected_client               
    
    def on_drop_menu_execute(self):
        self.drop_menu_select()
        self.output_2_update()             
        
    def excel_manip(self):
        self.excel_manipulator = ExcelManipulator()
        self.output_text1.append("Excel Manipulation Selected")
        search_list_format_info = self.file_config.search_list_format_info
        
        if search_list_format_info.get("client_name") == "NO CLIENT SELECTED":
            self.output_text1.append("No Client Selected")
        
        else:
            chartsearch_path = Path(self.file_config.path_data["chartsearch_xlsx"])

            concat_excel_file = self.excel_manipulator.pandas_column_clean(chartsearch_path, self.file_config.toggle_states)
            wb, self.excel_dimensions = self.excel_manipulator.openpyxl_format_workbook(concat_excel_file, self.file_config.alphabet)
            
            self.file_config.save_xlsx(wb, self.file_config.search_list_format_info)
            
        self.final_output_2_update()

    def on_excel_macro(self):
        self.output_text1.append("Excel Macro Selected")
        search_list_format_info = self.file_config.search_list_format_info
        
        if search_list_format_info.get("client_name") == "NO CLIENT SELECTED":
            self.output_text1.append("No Client Selected")
        else:
            excel_macro(search_list_format_info)

    def clear_output(self):
            self.output_text1.clear()
            self.populate_client_list_menu()
            
    def is_search_list(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Search List is ON")
            self.file_config.toggle_states["toggle_search_list"] = True

        else:
            self.output_text1.append("'Toggle Search List is OFF")
            self.file_config.toggle_states["toggle_search_list"] = False
        self.output_2_update()

    def has_attempt(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Attempts is ON")
            self.file_config.toggle_states["toggle_attempt"] = True

        else:
            self.output_text1.append("'Toggle Attempts is OFF")
            self.file_config.toggle_states["toggle_attempt"] = False
        self.output_2_update()

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
        self.output_2_update()

# FOR ADDING AND REMOVE CLIENT FROM JSON         
    def on_button_click(self, button_text):
        if button_text == "Add Client":
            client_name, ok = QInputDialog.getText(self, "Add Client", "Enter Client's Full Name:")
            if ok:
                self.file_config.add_client(client_name)
                self.output_text1.append(f"Added client: {client_name}")
                self.drop_menu.clear() 
            self.populate_client_list_menu()
        
        elif button_text == "Remove Client":
            client_index, ok = QInputDialog.getText(self, "Remove Client", "Enter Client's Index:")
            if ok:
                self.file_config.delete_client(client_index)
                self.output_text1.append(f"Removed Client at Index: {client_index}")
                self.drop_menu.clear() 
            self.populate_client_list_menu()
   
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())