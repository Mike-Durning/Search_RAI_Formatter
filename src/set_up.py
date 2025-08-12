from tkinter.filedialog import askdirectory
from datetime import datetime
from pathlib import Path
import json
import os

class Config:
    def __init__(self):
        self.src_folder = Path(__file__).resolve().parent
        self.project_folder = self.src_folder.parent
        
        self.config_folder      = self.project_folder / "config"
        self.assets_folder      = self.project_folder / "assets"
        self.client_drop_folder = self.project_folder / "client_drop_folder"
      
    def instantiate_dicts(self):
         
        self.alphabet = {
            "1": "A",
            "2": "B",
            "3": "C",
            "4": "D",
            "5": "E",
            "6": "F",
            "7": "G",
            "8": "H",
            "9": "I",
            "10": "J",
            "11": "K",
            "12": "L",
            "13": "M",
            "14": "N",
            "15": "O",
            "16": "P",
            "17": "Q",
            "18": "R",
            "19": "S",
            "20": "T",
            "21": "U",
            "22": "V",
            "23": "W",
            "24": "X",
            "25": "Y",
            "26": "Z"
        }
                
        self.path_data = {
            "project_folder"            : str(self.project_folder),
            "config_folder"             : str(self.config_folder),
            "client_drop_folder"        : str(self.client_drop_folder),
            "chartsearch_xlsx"          : str(self.client_drop_folder  / 'ChartSearch.xlsx'),
            "config_file_path_json"     : str(self.config_folder       / "file_path.json"),
            "settings_json"             : str(self.config_folder       / "settings.json"),
            "client_list_json"          : str(self.config_folder       / "client_list.json"),
            "assets_folder"             : str(self.assets_folder),
            "src_folder"                : str(self.src_folder),
            "config_set_default_folder" : str(),
        }
                
        self.search_list_format_info = {
            "client_name"      : "NO CLIENT SELECTED",
            "format"           : "Default - RAI Report",
            "date"             : datetime.now().strftime('%m.%d.%Y'),
            "file_type"        : ".xlsx",
            "custom_directory" : None
        }

        self.toggle_states = {
            "toggle_search_list"      : False,
            "toggle_attempt"          : False,
            "toggle_custom_directory" : False
        }    

    def folders_exist(self):
        self.config_folder.mkdir(exist_ok=True)
        self.assets_folder.mkdir(exist_ok=True)
        self.client_drop_folder.mkdir(exist_ok=True)

    def save_file_path_to_json(self):
        path_data = Path(self.path_data["config_file_path_json"])
        try:
            with open(path_data, "w") as file:
                json.dump(self.path_data, file, indent=4)
            return f"Config data saved to {path_data}"
        except Exception as e:
            return f"Error while saving to JSON: {e}"
        
    def save_settings_to_json(self):
        settings_dict = {}
        settings_json = self.config_folder / "settings.json"
        if not settings_json.exists():
            try:
                with open(settings_json, "w") as file:
                    json.dump(settings_dict, file, indent=4)
                return f"Config data saved to {settings_json}"
            except Exception as e:
                return f"Error while saving to JSON: {e}"
        else:
            return f"{settings_json} already exists. Skipping the saving process."
            
    def save_client_list_to_json(self):
        client_list_dict = {}
        client_list_json = self.config_folder / "client_list.json"
        if not client_list_json.exists():
            try:
                with open(client_list_json, "w") as file:
                    json.dump(client_list_dict, file, indent=4)
                return f"Config data saved to {client_list_json}"
            except Exception as e:
                return f"Error while saving to JSON: {e}"
        else:
            return f"{client_list_json} already exists. Skipping the saving process."

    def manual_return_path(self):
        return_dir = askdirectory(title="Select Folder")
        return return_dir

    def print_dict_or_json(self, dict_or_json_path):
        if isinstance(dict_or_json_path, dict):
            formatted_data = ""
            for key, value in dict_or_json_path.items():
                formatted_data += f"{key}: {value}\n"
            return formatted_data, dict_or_json_path
        elif isinstance(dict_or_json_path, Path):
            if dict_or_json_path.exists():
                with dict_or_json_path.open('r') as json_file:
                    json_dict = json.load(json_file)
                    formatted_data = ""
                    for key, value in json_dict.items():
                        formatted_data += f"{key}: {value}\n"
                    return formatted_data, json_dict, dict_or_json_path
        else:
            return "Dictionary or JSON file does not exist."
            
    def select_client_by_value(self, selected_value):
        clients_json_path = Path(self.path_data["client_list_json"])
        if clients_json_path.exists():
            with clients_json_path.open('r') as json_file:
                clients_data = json.load(json_file)

            for key, client_value in clients_data.items():
                if client_value == selected_value:
                    return client_value
        else:
            print("Clients JSON file does not exist.")
            return None

    def delete_client(self, key_to_delete):
        json_file_path = Path(self.path_data["client_list_json"])
        try:
            with json_file_path.open('r') as json_file:
                client_list = json.load(json_file)
        except FileNotFoundError:
            print("Clients JSON file does not exist.")
            return

        if key_to_delete in client_list:
            del client_list[key_to_delete]
            updated_clients_data = {}
            sorted_keys = sorted(map(int, client_list.keys()))
            for index, old_key in enumerate(sorted_keys):
                new_key = str(index + 1)
                updated_clients_data[new_key] = client_list[str(old_key)]
                
            with json_file_path.open('w') as json_file:
                json.dump(updated_clients_data, json_file, indent=4)
        else:
            print(f"Key {key_to_delete} does not exist in the clients JSON.")
            
    def add_client(self, name_to_add):
        json_file_path = Path(self.path_data["client_list_json"])
        try:
            with json_file_path.open('r') as json_file:
                client_list = json.load(json_file)
        except FileNotFoundError:
            print("Clients JSON file does not exist.")
            return

        if name_to_add in client_list.values():
            print(f"Client with name '{name_to_add}' already exists.")
            return

        new_key = str(len(client_list.keys()) + 1)

        client_list[new_key] = name_to_add

        with json_file_path.open('w') as json_file:
            json.dump(client_list, json_file, indent=4)        
            
    def save_xlsx(self, wb, search_list_format_info):


            todays_date = search_list_format_info["date"]
            file_type = search_list_format_info["file_type"]
            search_list = search_list_format_info["format"]
            client_name = search_list_format_info["client_name"]
            download_directory = search_list_format_info["custom_directory"]
            
            filename_save = f"{client_name} {search_list} {todays_date}{file_type}"

            
            if download_directory is None:
                download_directory = askdirectory(title="Select Client Folder")

            folder_path = Path(download_directory)
                
            file_path = folder_path / filename_save
            wb.save(file_path)
            wb.close()

            os.startfile(file_path)
