# Report: Server Updater on IIS

## Introduction
This report outlines the functionality, usage, and safety measures of the Server Updater program designed for managing servers on IIS. With a user-friendly graphical interface, the program streamlines the processes of updating, creating, or removing <product> servers, making it accessible even for non-technical users.

## Key Functionalities
The Server Updater offers four primary functions:
1. **Update Existing Server Version:** Replace files of an existing <product> server with a new version from a ZIP archive.
2. **Create New Site and Server:** Set up a new <product> server using files from a provided ZIP archive.
3. **Add Environments to Existing Server:** Incorporate additional environments (e.g., Prod, Test) into an existing server setup.
4. **Delete Existing Site and Server:** Remove an existing <product> server or its environments.

## Detailed Functionality

### 1. Update Existing Server Version
To update an existing server, the program follows these steps:
- **User Input:** Users select the ZIP file containing the new server files and the current server folder.
- **Validation:** The program checks the validity of the provided paths.
- **Website Stoppage:** It stops the IIS website and App pool linked to the server.
- **Backup Creation:** A backup folder is created to store the original files.
- **Source Folder Renaming:** The existing server folder is temporarily renamed.
- **File Extraction and Copying:** New files from the ZIP are extracted and copied over, while essential folders (like "App_Data," "Logs," and "Images") are preserved.
- **Website Restart:** The IIS website is restarted post-update.
- **Logging:** All actions are logged for reference.
- **Undo Option:** Users can revert to the original server version if needed.

### 2. Create New Site and Server
To create a new server, the program performs the following:
- **ZIP File Selection:** Users choose the ZIP file with the server files.
- **Environment Selection:** Users specify which environments (Prod, Dev, Test) to set up.
- **Site Name and Directory Selection:** Users input a site name and choose a base directory for the new server.
- **Directory Structure Creation:** The program creates a new directory with subdirectories for each environment.
- **App Pool and Site Creation:** Application pools and a new website are created using `appcmd.exe`.
- **File Extraction:** Contents of the ZIP file are extracted to the appropriate directories.

### 3. Add Environments to Existing Site and Server
This process is similar to creating a new server but targets an existing setup:
- **Website Selection:** Users enter the name of the existing site.
- **ZIP File and Environment Selection:** Users provide the ZIP file and choose environments to add.
- **Subdirectory Creation:** New subdirectories are created for each selected environment.
- **File Extraction and App Pool Creation:** The program extracts files and creates application pools for the new environments.

### 4. Delete Server
The deletion process is straightforward and ensures proper cleanup of the existing setup:
- **User Types in Site Name**
- **Application Pool is Stopped**
- **Site is Deleted**
- **Application Pool is Deleted**
- **All Files Are Still Available; Just the Website is Gone.**

## Safety Measures
The Server Updater incorporates various safety measures to minimize risks:
- **Input Validation:** Ensures that paths are valid before proceeding.
- **Website Stoppage:** Stops the IIS site to prevent data inconsistencies.
- **Backup Creation:** Allows for easy restoration of the previous server version.
- **Selective Data Copying:** Only essential folders are copied to avoid overwriting important data.
- **Error Handling:** Informs users of issues during the process for easier troubleshooting.
- **Rollback Options:** Provides undo functionality to revert changes.
- **Logging:** All actions are logged and recorded.

## How to Use the Server Updater

### Updating an Existing Server
1. Launch the Server Updater program.
2. Select the "Update Existing Server Version" option.
3. Choose the ZIP file and the current server folder.
4. Click "Extract and Copy" to start the update.
5. Monitor progress in the log window.
6. Confirm completion or select the undo option if needed.

### Creating a New Site and Server
1. Launch the Server Updater.
2. Choose the "Create New Site and Server" option.
3. Select the ZIP file and specify environments.
4. Enter the desired site name.
5. Choose a base directory and click "Create."

### Adding Environments to an Existing Server
1. Launch the Server Updater.
2. Select the option to add environments.
3. Enter the existing site name and select the ZIP file.
4. Choose new environments and set the base directory.
5. Click "Create" to proceed.

### Deleting Site from IIS
1. Launch the Server Updater.
2. Select the Delete Option.
3. Enter the existing site name.
4. Click "Delete" to proceed.

## Conclusion
The Server Updater program provides a reliable and efficient method for managing IIS Servers. With robust safety measures and a user-friendly interface, it simplifies server updates, creation, and management, making it an essential tool for both technical and non-technical users.
