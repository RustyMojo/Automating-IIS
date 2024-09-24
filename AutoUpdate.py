import os
import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
import sys
import subprocess
import logging
from datetime import datetime  # Add this import
import tkinter.messagebox as messagebox
import zipfile  # Import the zipfile module



class ServerUpdater:

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')
    appcmd_path = r'C:\Windows\System32\inetsrv\appcmd.exe'  # Full path to appcmd.exe
    end = r"/text:[path='/'].physicalPath"

    def back_to_menu(self):
        # Close the current GUI
        self.root.destroy()
        # Reopen the Menu GUI
        menu = Menu()
    def on_closing(self):

        sys.exit()  # Exit the program

    def __init__(self, menu_instance):
        self.menu_instance = menu_instance


        self.root = tk.Toplevel(self.menu_instance.root)
        self.root.title("Server Updater")
        self.root.geometry("400x300")  # Set a larger window size
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Create a frame to organize the widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10)

        # Create labels and entry fields for file selection
        self.zip_file_label = tk.Label(self.frame, text="Select New Server Folder:")
        self.zip_file_label.pack()

        self.zip_file_entry = tk.Entry(self.frame, width=50)
        self.zip_file_entry.pack()

        self.choose_zip_button = tk.Button(self.frame, text="Browse", command=self.choose_zip_file, bg="blue", fg="white")
        self.choose_zip_button.pack(pady=10)

        self.source_folder_label = tk.Label(self.frame, text="Select Current Folder:")
        self.source_folder_label.pack()

        self.source_folder_entry = tk.Entry(self.frame, width=50)
        self.source_folder_entry.pack()

        self.choose_source_button = tk.Button(self.frame, text="Browse", command=self.choose_source_folder, bg="blue", fg="white")
        self.choose_source_button.pack(pady=10)

        self.process_button = tk.Button(self.frame, text="Extract and Copy", command=self.extract_and_copy, bg="green", fg="white", width=20)
        self.process_button.pack(pady=5)  # Add vertical spacing below the button

        self.process_button = tk.Button(self.frame, text="Back", command=self.back_to_menu, bg="red", fg="white", width=20)
        self.process_button.pack(pady=20)  # Add vertical spacing below the button

        self.menu_instance = menu_instance


        # Variables to store selected paths
        self.zip_file = ""
        self.source_folder = ""
        self.log_window = None  # A reference to the log window
        self.log_text = None
        self.root.mainloop()

    def choose_zip_file(self):
        self.zip_file = filedialog.askopenfilename(title="Select Server Zip")
        self.zip_file_entry.delete(0, tk.END)
        self.zip_file_entry.insert(0, self.zip_file)

    def choose_source_folder(self):
        self.source_folder = filedialog.askdirectory(title="Select Source Folder")
        self.source_folder_entry.delete(0, tk.END)
        self.source_folder_entry.insert(0, self.source_folder)

    def extract_and_copy(self):

        try:

            if not os.path.isfile(self.zip_file):
                self.log_message("Please select a valid ZIP file.")
                return

            if not os.path.isdir(self.source_folder):
                self.log_message("Please select a valid source folder.")
                return
            self.stop_website()

            self.log_window = tk.Toplevel(self.root)
            self.log_window.title("Extraction and Copy Log")

            # Extend the width of the log text widget
            self.log_text = scrolledtext.ScrolledText(self.log_window, wrap=tk.WORD, width=100, height=20)
            self.log_text.pack()

            # Add Close and Undo buttons to the Log GUI
            close_button = tk.Button(self.log_window, text="Close", command=self.close_log_window,  bg="blue", fg="white")
            undo_button = tk.Button(self.log_window, text="Undo", command=self.undo_copy, bg="red", fg="white")

            close_button.pack(side=tk.RIGHT)
            undo_button.pack(side=tk.RIGHT)

            self.log_message("Starting extraction and copy process...\n")


            source_folder_path = os.path.join(self.source_folder)        
            parent_folder = os.path.dirname(source_folder_path)
            backup_folder_path = os.path.join(parent_folder, 'backup')

            if os.path.exists(backup_folder_path) and os.path.isdir(backup_folder_path):
                shutil.rmtree(backup_folder_path)
                self.log_message(f"Deleted existing 'backup' folder: {backup_folder_path}\n")


            source_folder_name = os.path.basename(source_folder_path)
            temp_rename_path = os.path.join(parent_folder, f"backup")

            try:
                os.rename(source_folder_path, temp_rename_path)
                self.log_message(f"Renamed source_folder to '{temp_rename_path}'.\n")
            except Exception as rename_err:
                self.log_message(f"Failed to rename source folder: {str(rename_err)}")

            # Create a destination folder for extraction
            extraction_folder = os.path.join(parent_folder, source_folder_name)

            try:
                # Extract the zip file into the extraction folder while preserving the folder structure
                shutil.unpack_archive(self.zip_file, extract_dir=extraction_folder, format='zip')
                self.log_message(f"Extracted files to: {extraction_folder}\n")
            except Exception as extract_err:
                self.log_message(f"Failed to extract files: {str(extract_err)}")


            source_app_data = os.path.join(temp_rename_path, "App_Data")
            dest_app_data = os.path.join(extraction_folder, "App_Data")

            try:
                if os.path.exists(source_app_data) and os.path.isdir(source_app_data):
                    os.makedirs(dest_app_data, exist_ok=True)
                    shutil.copytree(source_app_data, dest_app_data, dirs_exist_ok=True)
                    self.log_message(f"Copied {source_app_data} to {dest_app_data} (replaced if exists)\n")
            except Exception as copy_err:
                self.log_message(f"Failed to copy App_Data: {str(copy_err)}")

            try:
            # Copy the "logs" folder
                source_logs = os.path.join(temp_rename_path, "Logs")
                dest_logs = os.path.join(extraction_folder, "Logs")
    
                if os.path.exists(source_logs) and os.path.isdir(source_logs):
                    os.makedirs(dest_logs, exist_ok=True)
                    shutil.copytree(source_logs, dest_logs, dirs_exist_ok=True)
                    self.log_message(f"Copied {source_logs} to {dest_logs} (replaced if exists)\n")
            except Exception as copy_logs_err:
                self.log_message(f"Failed to copy 'logs' folder: {str(copy_logs_err)}")

            try:
                # Copy the "images" folder
                source_images = os.path.join(temp_rename_path, "Images")
                dest_images = os.path.join(extraction_folder, "Images")

                if os.path.exists(source_images) and os.path.isdir(source_images):
                    os.makedirs(dest_images, exist_ok=True)
                    shutil.copytree(source_images, dest_images, dirs_exist_ok=True)
                    self.log_message(f"Copied {source_images} to {dest_images} (replaced if exists)\n")
            except Exception as copy_images_err:
                self.log_message(f"Failed to copy 'images' folder: {str(copy_images_err)}")


            self.log_message("Extraction and copying process completed.\n")
            self.root.withdraw()
        except Exception as e:
            print(f"An error occaaurred: {str(e)}")
            
    def log_message(self, message):
        try:
            log_entry = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - INFO: {message}\n"

            # Append the message to the log text widget if it exists
            if self.log_text:
                self.log_text.insert(tk.END, log_entry)
                self.log_text.see(tk.END)

            # Get the path to the parent folder of the source folder
            parent_folder = os.path.dirname(self.source_folder)

            # Log the message to a log file
            log_file_path = os.path.join(parent_folder, "log.txt")

            with open(log_file_path, "a") as log_file:
                if log_entry is not None:
                    log_file.write(log_entry)
        except Exception as e:
            self.log_message(f"An error occurred while logging: {str(e)}")

    def close_log_window(self):
        if self.log_window:
            self.start_website()

            sys.exit()

    def undo_copy(self):
        try:
            if not self.source_folder or not os.path.isdir(self.source_folder):
                self.log_message("Cannot undo: Source folder is not valid.")
                return

            if not self.zip_file or not os.path.isfile(self.zip_file):
                self.log_message("Cannot undo: ZIP file is not valid.")
                return
            # Check if the user confirms the undo operation
            confirmation = messagebox.askyesno("Confirm Undo", "Are you sure you want to undo the operation?")
            if not confirmation:
                return

            source_folder_path = os.path.join(self.source_folder)
            parent_folder = os.path.dirname(source_folder_path)
            backup_folder_path = os.path.join(parent_folder, 'backup')

            if not os.path.exists(backup_folder_path) or not os.path.isdir(backup_folder_path):
                self.log_message("Cannot undo: Backup folder does not exist.")
                return

            extraction_folder = os.path.join(parent_folder, os.path.basename(source_folder_path))

            if os.path.exists(extraction_folder) and os.path.isdir(extraction_folder):
                shutil.rmtree(extraction_folder)
                self.log_message(f"Deleted Updated Version: {extraction_folder}\n")

            temp_rename_path = os.path.join(parent_folder, f"backup")

            try:
                os.rename(temp_rename_path, source_folder_path)
                self.log_message(f"Restored Original Folder: '{source_folder_path}'.\n")
            except Exception as rename_err:
                self.log_message(f"Failed to restore original folder: {str(rename_err)}")

            self.log_message("Undo process completed.")
        except Exception as e:
            self.log_message(f"An error occurred during undo: {str(e)}")

    def start_website(self):
        try:
            pool = self.locate_pool()
            # Construct the command to start the website with /apppool.name:
            command = [self.appcmd_path, 'start', 'apppool', f'/apppool.name:{pool}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)

            # Print the output
            self.log_message(output)

        except Exception as e:
            self.log_message(f"An error occurred while starting the website: {str(e)}")

    def stop_website(self):
        try:
            pool = self.locate_pool()

            # Construct the command to stop the website with /apppool.name:
            command = [self.appcmd_path, 'stop', 'apppool', f'/apppool.name:{pool}']
            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)


            # Print the output
            self.log_message(output)
        except Exception as e:

            self.log_message(f"An error occurred while stopping the website: {str(e)}")
            
    
    def locate_pool(self):
        # Define the appcmd commands
        apppool_command = [self.appcmd_path, 'list', 'app', f'/text:applicationPool']
        physical_path_command = [self.appcmd_path, 'list', 'app', self.end]
        mapping = {}

        try:
            # Get the outputs of appcmd commands
            ap = subprocess.check_output(apppool_command, shell=True, text=True)
            pt = subprocess.check_output(physical_path_command, shell=True, text=True)

            # Split the outputs into lines
            app_pool_lines = ap.strip().split('\n')
            physical_path_lines = pt.strip().split('\n')

            # Create a dictionary to store the mappings
            source_folder_path = os.path.normpath(self.source_folder)        

            # Populate the dictionary with app pool names and physical paths
            for i in range(len(app_pool_lines)):
                app_pool_name = app_pool_lines[i].strip()
                if app_pool_name:
                    physical_path = physical_path_lines[i].strip()
                    mapping[physical_path] = app_pool_name


        except subprocess.CalledProcessError as e:
            print(f"Error running appcmd: {e}")

        corresponding_app_pool = mapping.get(source_folder_path)

        return corresponding_app_pool

class Menu:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("<product> Management")
        self.root.geometry("400x300")  # Set a larger window size

        # Create a frame to organize the widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=20)  # Add vertical spacing around the frame

        # Create a title label
        title_label = tk.Label(self.frame, text="<product> Management", font=("Arial", 16))
        title_label.pack()

        # Add some vertical spacing
        tk.Label(self.frame, text="").pack()
        # Create buttons for Update and New
        update_button_text = "Updates Existing Server Version"
        update_text = tk.Label(self.frame, text=update_button_text, font=("Arial", 10))

        update_button = tk.Button(self.frame, text="Update", command=self.run_updater, bg="blue", fg="white", width=20)
        update_button.pack(pady=10)  # Add vertical spacing below the button

        # Bind the event to show the popup
        update_button.bind("<Enter>", lambda event: self.show_popup(update_text, update_button_text))
        # Bind the event to hide the popup
        update_button.bind("<Leave>", lambda event: update_text.pack_forget())

        new_button_text = "Creates entire new Site and Server"
        new_text = tk.Label(self.frame, text=new_button_text, font=("Arial", 10))

        new_button = tk.Button(self.frame, text="New", command=self.run_new, bg="green", fg="white", width=20)
        new_button.pack(pady=10)  # Add vertical spacing below the button

        # Bind the event to show the popup
        new_button.bind("<Enter>", lambda event: self.show_popup(new_text, new_button_text))
        # Bind the event to hide the popup
        new_button.bind("<Leave>", lambda event: new_text.pack_forget())

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        edit_button_text = "Adds environments to existing Site and Server"
        edit_text = tk.Label(self.frame, text=edit_button_text, font=("Arial", 10))

        edit_button = tk.Button(self.frame, text="Edit", command=self.run_edit, bg="orange", fg="white", width=20)
        edit_button.pack(pady=10)  # Add vertical spacing below the button

        # Bind the event to show the popup
        edit_button.bind("<Enter>", lambda event: self.show_popup(edit_text, edit_button_text))
        # Bind the event to hide the popup
        edit_button.bind("<Leave>", lambda event: edit_text.pack_forget())

        delete_button_text = "Deletes existing Site and Server"
        delete_text = tk.Label(self.frame, text=delete_button_text, font=("Arial", 10))

        delete_button = tk.Button(self.frame, text="Delete", command=self.run_delete, bg="red", fg="white", width=20)
        delete_button.pack(pady=10)  # Add vertical spacing below the button

        # Bind the event to show the popup
        delete_button.bind("<Enter>", lambda event: self.show_popup(delete_text, delete_button_text))
        # Bind the event to hide the popup
        delete_button.bind("<Leave>", lambda event: delete_text.pack_forget())

    def show_popup(self, label, text):
        label.config(text=text)
        label.pack()

    def on_closing(self):

        sys.exit()  # Exit the program


    def run_updater(self):
        # Run the ServerUpdater class and close the current menu GUI
        self.root.withdraw()  # Hide the current menu GUI
        updater = ServerUpdater(self)

    def run_new(self):
        # Implement the functionality for the "New" button here
        # Run the ServerUpdater class and close the current menu GUI
        self.root.withdraw()  # Hide the current menu GUI
        new = newClient(self)

    def run_edit(self):
        # Implement the functionality for the "New" button here
        # Run the ServerUpdater class and close the current menu GUI
        self.root.withdraw()  # Hide the current menu GUI
        new = updateClient(self)

    def run_delete(self):
        # Implement the functionality for the "New" button here
        # Run the ServerUpdater class and close the current menu GUI
        self.root.withdraw()  # Hide the current menu GUI
        new = DeleteClient(self)

class newClient:
    # Set up logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')
    
    # Define constants
    appcmd_path = r'C:\Windows\System32\inetsrv\appcmd.exe'  # Full path to appcmd.exe
    end = r"/text:[path='/'].physicalPath"
    
    def __init__(self, menu_instance):
        self.menu_instance = menu_instance

        # Create a new tkinter Toplevel window
        self.root = tk.Toplevel(self.menu_instance.root)
        self.root.title("New Client")
        self.root.geometry("400x400")

        # Create a frame for widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Create the title label
        title_label = tk.Label(self.frame, text="New Client", font=("Arial", 16))
        title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # Create the entry label
        entry_label = tk.Label(self.frame, text="Enter text:", borderwidth=1)
        entry_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        # Create an entry field for user input
        self.zip_file_entry = tk.Entry(self.frame, width=50)
        self.zip_file_entry.grid(row=1, column=1, padx=10, pady=10)

        # Checkboxes
        self.prod_var = tk.BooleanVar()
        self.prod_checkbox = tk.Checkbutton(self.frame, text="PROD", variable=self.prod_var)
        self.prod_checkbox.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.dev_var = tk.BooleanVar()
        self.dev_checkbox = tk.Checkbutton(self.frame, text="DEV", variable=self.dev_var)
        self.dev_checkbox.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.test_var = tk.BooleanVar()
        self.test_checkbox = tk.Checkbutton(self.frame, text="TEST", variable=self.test_var)
        self.test_checkbox.grid(row=4, column=0, padx=10, pady=5, sticky="w")

        # Create a button to create directories
        self.create_button = tk.Button(self.frame, text="Create", command=self.create_directories, bg="blue", fg="white", width=12)
        self.create_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Create a button to upload a ZIP file
        self.upload_button = tk.Button(self.frame, text="Upload ZIP File", command=self.upload_zip, bg="green", fg="white", width=12)
        self.upload_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Label to display the selected ZIP file
        self.selected_zip_label = tk.Label(self.frame, text="")
        self.selected_zip_label.grid(row=5, column=1, padx=10, pady=10)

        # Create a button to go back
        self.back_button = tk.Button(self.frame, text="Back", command=self.back_to_menu, bg="red", fg="white", width=12)
        self.back_button.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def back_to_menu(self):
        self.root.destroy()
        menu = Menu()

    def on_closing(self):
        sys.exit()

    def create_directories(self):
        try:
            # Ask the user to select a directory for creating subdirectories
            base_dir = filedialog.askdirectory(title="Select Base Directory")

            if not base_dir:
                return  # User canceled, do nothing

            # Create subdirectories based on ticked checkboxes
            subdirs = []
            if self.prod_var.get():
                subdirs.append("PROD")
            if self.dev_var.get():
                subdirs.append("DEV")
            if self.test_var.get():
                subdirs.append("TEST")

            if not subdirs:
                messagebox.showwarning("Warning", "No checkboxes are selected.")
                return  # No subdirectories selected, do nothing

            entry_label_text = self.zip_file_entry.get()  # Get the text from the entry field

            # Create a directory in the same location as base_dir with the specified name
            site_directory = os.path.join(base_dir, entry_label_text)
            os.makedirs(site_directory, exist_ok=True)

            # Create a "<product>" directory within the newly created directory
            red_office_directory_path = os.path.join(site_directory, "<product>")
            os.makedirs(red_office_directory_path, exist_ok=True)

            # Make the app pool
            self.make_appPool(entry_label_text)

            # Make the site
            self.make_site(red_office_directory_path, entry_label_text)

            zip_file_path = self.selected_zip_label.cget("text")
            if zip_file_path and os.path.isfile(zip_file_path):
                with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                    for subdir in subdirs:
                        extract_path = os.path.join(red_office_directory_path, subdir)
                        os.makedirs(extract_path, exist_ok=True)
                        zip_ref.extractall(extract_path)

                        variable_name = f"{entry_label_text}<product>{subdir}"
                        self.make_appPool(variable_name)
                        self.make_app(extract_path, subdir, variable_name)

                messagebox.showinfo("Success", "Export and extraction completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def upload_zip(self):
        try:
            zip_file_path = filedialog.askopenfilename(title="Select ZIP File")
            self.selected_zip_label.config(text=zip_file_path)  # Update the label text with the selected ZIP file path
            # Do something with the selected ZIP file
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def make_appPool(self, apppool):
        try:
            # Construct the command to start the website with /apppool.name:
            command = [self.appcmd_path, 'add', 'apppool', f'/name:{apppool}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the app pool: {str(e)}")


    def make_site(self, physicalPath, apppoolName):
        try:
            siteName = self.zip_file_entry.get()  # Get the text from the entry field

            # Command to add the site
            command = [self.appcmd_path, 'add', 'site', f'/name:{siteName}', '/bindings:http/*:81:', f'/physicalPath:{physicalPath}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)

            # Reassign the app pool for the site
            command = [self.appcmd_path, 'set', 'site', f'/site.name:{siteName}', f"/[path='/'].applicationPool:{apppoolName}"]
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the site: {str(e)}")

    def make_app(self, physicalPath, env, variable_name):
        try:
            # Get the site name from the entry field
            siteName = self.zip_file_entry.get()

            # Command to add the application
            command = [self.appcmd_path, 'add', 'app', f'/site.name:{siteName}', f'/path:/{env}', f'/physicalPath:{physicalPath}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)

            # Set the application pool for the application
            command = [self.appcmd_path, 'set', 'app', f'/app.name:{siteName}/{env}', f'/applicationPool:{variable_name}']
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the application: {str(e)}")


class updateClient:
    # Set up logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')
    
    # Define constants
    appcmd_path = r'C:\Windows\System32\inetsrv\appcmd.exe'  # Full path to appcmd.exe
    end = r"/text:[path='/'].physicalPath"
    
    def __init__(self, menu_instance):
        self.menu_instance = menu_instance

        # Create a new tkinter Toplevel window
        self.root = tk.Toplevel(self.menu_instance.root)
        self.root.title("New Client")
        self.root.geometry("400x400")

        # Create a frame for widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Create the title label
        title_label = tk.Label(self.frame, text="Update Client", font=("Arial", 16))
        title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

                # Create the entry label
        entry_label = tk.Label(self.frame, text="Enter curent IIS Site name:", borderwidth=1)
        entry_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
                        # Create the entry label


        # Create an entry field for user input
        self.zip_file_entry = tk.Entry(self.frame, width=50)
        self.zip_file_entry.grid(row=1, column=1, padx=10, pady=10)

        custom_label = tk.Label(self.frame, text="Custom Name:", borderwidth=1)
        custom_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
                # Create an entry field for user input
        self.custom_name = tk.Entry(self.frame, width=50)
        self.custom_name.grid(row=2, column=1, padx=10, pady=10)


        # Checkboxes
        self.prod_var = tk.BooleanVar()
        self.prod_checkbox = tk.Checkbutton(self.frame, text="PROD", variable=self.prod_var)
        self.prod_checkbox.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.dev_var = tk.BooleanVar()
        self.dev_checkbox = tk.Checkbutton(self.frame, text="DEV", variable=self.dev_var)
        self.dev_checkbox.grid(row=4, column=0, padx=10, pady=5, sticky="w")

        self.test_var = tk.BooleanVar()
        self.test_checkbox = tk.Checkbutton(self.frame, text="TEST", variable=self.test_var)
        self.test_checkbox.grid(row=5, column=0, padx=10, pady=5, sticky="w")





        # Create a button to create directories
        self.create_button = tk.Button(self.frame, text="Create", command=self.create_directories, bg="blue", fg="white", width=12)
        self.create_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Create a button to upload a ZIP file
        self.upload_button = tk.Button(self.frame, text="Upload ZIP File", command=self.upload_zip, bg="green", fg="white", width=12)
        self.upload_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Label to display the selected ZIP file
        self.selected_zip_label = tk.Label(self.frame, text="")
        self.selected_zip_label.grid(row=6, column=1, padx=10, pady=10)

        # Create a button to go back
        self.back_button = tk.Button(self.frame, text="Back", command=self.back_to_menu, bg="red", fg="white", width=12)
        self.back_button.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def back_to_menu(self):
        self.root.destroy()
        menu = Menu()

    def on_closing(self):
        sys.exit()

    def create_directories(self):
        try:
            # Ask the user to select a directory for creating subdirectories
            base_dir = filedialog.askdirectory(title="Select Base Directory")

            if not base_dir:
                return  # User canceled, do nothing

            # Create subdirectories based on ticked checkboxes
            subdirs = []
            if self.prod_var.get():
                subdirs.append("PROD")
            if self.dev_var.get():
                subdirs.append("DEV")
            if self.test_var.get():
                subdirs.append("TEST")
        # Get the custom label text
            custom_text = self.custom_name.get().strip()  # Remove leading/trailing spaces

        # Append the custom label text to the subdirs list if it's not empty
            print(subdirs)
            if custom_text:
                subdirs.append(custom_text)
            if not subdirs:
                messagebox.showwarning("Warning", "No checkboxes are selected.")
                return  # No subdirectories selected, do nothing

            entry_label_text = self.zip_file_entry.get()  # Get the text from the entry field


            zip_file_path = self.selected_zip_label.cget("text")
            if zip_file_path and os.path.isfile(zip_file_path):
                with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                    for subdir in subdirs:
                        extract_path = os.path.join(base_dir, subdir)
                        os.makedirs(extract_path, exist_ok=True)
                        zip_ref.extractall(extract_path)

                        variable_name = f"{entry_label_text}<product>{subdir}"
                        self.make_appPool(variable_name)
                        self.make_app(extract_path, subdir, variable_name)

                messagebox.showinfo("Success", "Export and extraction completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def upload_zip(self):
        try:
            zip_file_path = filedialog.askopenfilename(title="Select ZIP File")
            self.selected_zip_label.config(text=zip_file_path)  # Update the label text with the selected ZIP file path
            # Do something with the selected ZIP file
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def make_appPool(self, apppool):
        try:
            # Construct the command to start the website with /apppool.name:
            command = [self.appcmd_path, 'add', 'apppool', f'/name:{apppool}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the app pool: {str(e)}")



    def make_app(self, physicalPath, env, variable_name):
        try:
            # Get the site name from the entry field
            siteName = self.zip_file_entry.get()

            # Command to add the application
            command = [self.appcmd_path, 'add', 'app', f'/site.name:{siteName}', f'/path:/{env}', f'/physicalPath:{physicalPath}']

            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)

            # Set the application pool for the application
            command = [self.appcmd_path, 'set', 'app', f'/app.name:{siteName}/{env}', f'/applicationPool:{variable_name}']
            output = subprocess.check_output(command, shell=True, text=True)
            print(output)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the application: {str(e)}")

class DeleteClient:
    # Set up logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

    # Define constants
    appcmd_path = r'C:\Windows\System32\inetsrv\appcmd.exe'  # Full path to appcmd.exe

    def __init__(self, menu_instance):
        self.menu_instance = menu_instance

        # Create a new tkinter Toplevel window
        self.root = tk.Toplevel(self.menu_instance.root)
        self.root.title("Delete Client")
        self.root.geometry("400x200")

        # Create a frame for widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Create the title label
        title_label = tk.Label(self.frame, text="Delete Client", font=("Arial", 16))
        title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # Create the entry label
        entry_label = tk.Label(self.frame, text="Enter IIS Site name to delete:", borderwidth=1)
        entry_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        # Create an entry field for user input
        self.site_name_entry = tk.Entry(self.frame, width=50)
        self.site_name_entry.grid(row=1, column=1, padx=10, pady=10)

        # Create a button to delete the site
        self.delete_button = tk.Button(self.frame, text="Delete", command=self.delete_site, bg="red", fg="white", width=12)
        self.delete_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

        # Create a button to go back
        self.back_button = tk.Button(self.frame, text="Back", command=self.back_to_menu, bg="blue", fg="white", width=12)
        self.back_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def back_to_menu(self):
        self.root.destroy()
        menu = Menu()  # Ensure this references your main menu class correctly

    def on_closing(self):
        sys.exit()

    def delete_site(self):
        try:
            site_name = self.site_name_entry.get().strip()  # Get the site name input

            if not site_name:
                messagebox.showwarning("Warning", "Please enter a site name.")
                return  # No site name provided, do nothing

            # Command to delete the website
            self.stop_website()
            command = [self.appcmd_path, 'delete', 'site', f'{site_name}']
            get_pool = self.locate_pool()
            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            logging.info(output)
            messagebox.showinfo("Success", f"Site '{site_name}' deleted successfully!")

            self.delete_apppool(get_pool)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to delete site: {str(e.output)}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while deleting the site: {str(e)}")


    def stop_website(self):
        try:

            pool = self.locate_pool()

            # Construct the command to stop the website with /apppool.name:
            command = [self.appcmd_path, 'stop', 'apppool', f'/apppool.name:{pool}']
            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            logging.info(output)


        except Exception as e:

            print(e)           
    def delete_apppool(self, pool):
        try:

            

            # Construct the command to stop the website with /apppool.name:
            command = [self.appcmd_path, 'delete', 'apppool', f'/apppool.name:{pool}']
            # Run the command and capture the output
            output = subprocess.check_output(command, shell=True, text=True)
            logging.info(output)



        except Exception as e:

            print(e)     
    
    def locate_pool(self):
        site_name = self.site_name_entry.get().strip()  # Get the site name input

        # Define the appcmd commands
        apppool_command = [self.appcmd_path, 'list', 'app', f'{site_name}/', f'/text:applicationPool']

        try:
            # Get the outputs of appcmd commands
            ap = subprocess.check_output(apppool_command, shell=True, text=True)


        except subprocess.CalledProcessError as e:
            print(f"Error running appcmd: {e}")


        return ap

def main():
    # Create an instance of the Menu class
    menu_box = Menu()
    menu_box.root.mainloop()

if __name__ == "__main__":
    main()
