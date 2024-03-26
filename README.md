# PowerShell Script for System Configuration Analysis (CheckARM)

## Description
CheckARM is a multifunctional PowerShell script designed for analysis and reporting on various aspects of computer configurations and system status. The script is developed for system administrators and IT professionals who require detailed information about the working environment of individual computers or groups of computers in a network. Its modular design allows for flexible execution modes and reporting.

## Key Features
- **Multiple Modes of Operation**: Supports analysis of single computers, multiple computers, and Active Directory organizational units.
- **Extensive Reporting**: Generates detailed reports, including system information, network settings, security configurations, and much more.
- **Multithreading Capability**: Enables checking computers in a multithreaded mode for faster analysis and reporting.
- **Customizable Parameters**: Allows users to specify parameters for individual analysis, including system version, report version, and parallel execution settings.
- **Convenient Interface**: Offers an interactive console-based interface with clear instructions and options.
- **Advanced Error Handling**: Includes reliable error-checking mechanisms to ensure accuracy and reliability of execution.

## Analyzed OS Sections
CheckARM analyzes various OS sections, depending on the report type:
#### The Short Report
1. General Information
2. Operating System
3. Network Configuration
4. Network Time Protocol
5. Windows Updates
6. Local Accounts
7. Local Groups
8. Shared Folders
9. Antivirus Software
10. Third-Party Software
11. Password Policy
#### The Full Report (Includes the Short Report)
1. Autorun
2. Task Scheduler
3. System Software (embedded)
4. Services
5. System Event Log
6. Security Event Log

## Windows Event Logs Analysis

The CheckARM script includes a detailed analysis of Windows Event Logs, focusing on both Security Events and System Events to provide insights into security-related activities and the operational health of the system. This analysis covers events from the last 60 days, ensuring a comprehensive view of recent activities on the system.

### Security Events

The Security Events subsection aims to highlight critical security-related events that could indicate potential security breaches, policy changes, or other security risks. The script specifically analyzes the following Event IDs, each corresponding to a particular type of security event:

- **4609**: Windows is shutting down.
- **4625**: An account failed to log on.
- **4697**: A service was installed in the system.
- **4698**: A scheduled task was created.
- **4719**: System audit policy was changed.
- **4720**: A user account was created.
- **4722**: A user account was enabled.
- **4723**: An attempt was made to change an account's password.
- **4724**: An attempt was made to reset an account's password.
- **4725**: A user account was disabled.
- **4726**: A user account was deleted.
- **4728**: A member was added to a security-enabled global group.
- **4729**: A member was removed from a security-enabled global group.
- **4731**: A security-enabled local group was created.
- **4732**: A member was added to a security-enabled local group.
- **4733**: A member was removed from a security-enabled local group.
- **4734**: A security-enabled local group was deleted.
- **4735**: A security-enabled local group was changed.
- **4738**: A user account was changed.
- **4740**: A user account was locked out.
- **4776**: The domain controller attempted to validate the credentials for an account.

### System Events

This subsection focuses on the analysis of the System Log, which includes information about significant system operations, changes in service statuses, and other vital system-related activities. The script captures only unique events from the System Log without filtering by specific Event IDs.

## Project Structure

The CheckARM project is organized into several directories, each serving a specific purpose in the analysis and reporting process:

- **errors/**: Contains text files logging errors (one for each computer analyzed).
- **reports/**: Directory where generated reports are saved.
- **src/**: Contains additional files used by the script.
    - `config.json`: Basic operation parameters. The script reads parameters from this file.  
- **src/data/**: Contains data required for parameter comparison and the application icon.
	  - `Services.json`: Lists standard OS service parameters for comparison against the computer's services.
    - `Software.json`: Lists names and versions of standard software for comparison purposes.
    - `groups.json`: Contains names of standard local groups of the operating system.
	  - `checkarm_ico.ico`: Icon image used in the application.
- **src/functions/**: PowerShell scripts that are included and used by the main script:
    - `main.ps1`: Contains core functions and utilities used by `CheckARM.ps1`.
    - `multi.ps1`: Handles analysis of multiple computers or IP ranges.
    - `single.ps1`: Dedicated to analyzing a single computer.
    - `scriptblocks.ps1`: Contains script blocks for parallel execution and specific checks.
- **src/jsons/**: Directory for storing JSON files with computer paraeters for each block. This directory is cleared when the script starts.
- **src/reporting/**: Contains files necessary for report creation.
	  - `report_create.exe`: The compiled script for generating reports in DOCX format.
- **src/reporting/templates/**: Contains report templates in .docx format.
- **src/scripts/**: Contains PowerShell script blocks that extract computer parameters in JSON format to the `src/jsons` directory. Each script corresponds to a section of the report:
    - `section1.ps1`: General Information
    - `section2.ps1`: Operating System
    - `section3.ps1`: Network Configuration
    - `section4.ps1`: Network Time Protocol
    - `section5.ps1`: Windows Updates
    - `section6.ps1`: Local Accounts and Local Groups
    - `section8.ps1`: Shared Folders
    - `section9.ps1`: Antivirus Software
    - `section10.ps1`: Third-Party Software and System Software
    - `section11.ps1`: Autorun
    - `section12.ps1`: Task Scheduler
    - `section13.ps1`: System and Security Event Logs
    - `section14.ps1`: Services
    - `section16.ps1`: Password Policy

## Service Files
The script utilizes several JSON service files to tailor its operation and reporting. These files can be modified to fit specific analysis needs:

- **Config.json**: Contains basic parameters for script operation. Parameters from this file are read at script execution and can be altered during runtime. To describe parameters, run the script with the `--help` argument.
- **Services.json**: Lists standard operating system services parameters. The script compares the list of services extracted from the computer against those specified in this file. Services not found in the standard list but present on the computer are highlighted in the report. The file should contain a list of services with the following parameters for each service: "Name", "Startmode", "State", "DisplayName". Services are compared only when generating the full report.
- **Software.json**: Contains names and versions of software installed in the operating system. The script compares the list of software installed on the computer against the software listed in this file. Software not found in the standard list but installed on the computer is highlighted in the report. Version checks are also performed. The file should list software in the format: "App1", "Version app1", "App2", "Version app2".
- **Groups.json**: Contains names of standard local groups of the operating system. The script compares the list of local groups on the computer with the standard ones. If there are non-standard groups on the computer, the script highlights them in the report.

## Usage
CheckARM offers several modes of operation, catering to different scenarios of system analysis. Here are the primary features and how to use them:

### Single Mode
- **Description**: Analyze a single computer. By default, the script targets the computer it's executed on. Alternatively, you can specify a computer name or IP address.
- **Default Usage**: `CheckARM.exe -c PCNAME` or `CheckARM.exe -c 192.168.0.1`
- **With Custom Options**: `CheckARM.exe -c PCNAME -rv short -bn 19000 -st 4`
- **Measure Execution Time**: `CheckARM.exe -c 192.168.0.1 -mt`
- **Create JSON Files Without Report**: `CheckARM.exe -c PCNAME -jo`

### Multiple Mode
- **Description**: Analyze multiple computers. Specify a path to a .txt file (each computer on a new line) or a .csv file (comma-separated). Supports IP address ranges.
- **Default Usage**: `CheckARM.exe -f C:\\Users\\User\\Desktop\\arms.txt` or `CheckARM.exe -i 192.168.0.1-192.168.0.100`
- **With Custom Options**: `CheckARM.exe -f C:\\Users\\User\\Desktop\\arms.txt -rv short -bn 19000 -st 5 -at 10`
- **Measure Execution Time**: `CheckARM.exe -i 192.168.0.1-192.168.0.100 -st 3 -mt`
- **Create JSON Files Without Report**: `CheckARM.exe -f C:\\Users\\User\\Desktop\\arms.txt -jo`

### Active Directory Mode
- **Description**: Analyze multiple computers within a specified organizational unit (OU). Requires the `Get-ADComputer` cmdlet in PowerShell.
- **Default Usage**: `CheckARM.exe -o OU=testou,DC=testdc`
- **With Custom Options**: `CheckARM.exe -o OU=testou,DC=testdc -rv short -bn 19000 -at 2`
- **Measure Execution Time**: `CheckARM.exe -o OU=testou,DC=testdc -mt`
- **Create JSON Files Without Report**: `CheckARM.exe -o OU=testou,DC=testdc -jo`

### Create a Report
- **Description**: Generates reports based on previously exported JSON files. Place folders named after computers containing JSON files in `src/jsons/`. This mode does not connect to devices.

### Change Options
Allows changing basic parameters such as:
- `report_version`: "short" or "full"
- `build_number`: OS build number. The script compares this with the actual build number on the device.
- `script_threads`: Number of threads for running PowerShell script blocks. Affects how many PowerShell script blocks (from `src/scripts/`) run concurrently on the target device. Setting values above 14 is not advisable.
- `arms_threads`: Number of threads for analyzing devices. Influences how many computers are analyzed simultaneously.
- `measure_time`: Displays the total execution time of the script.
- `json_only`: Exports parameters of specified devices into JSON files without generating a report.
- `report_only`: Generates reports for previously exported device parameters. Requires folders in `src/jsons/` named after the computers, containing JSON files.

### Show Help
- **Description**: Displays help information.

The script can operate in both interactive and non-interactive modes. For non-interactive mode, pass the arguments accordingly. For more detailed information, use the `--help` flag.

### Examples
**Single Mode:**
- Use the default options:
  - `CheckARM.exe -c PCNAME`
  - `CheckARM.exe -c 192.168.0.1`

**Multiple Mode:**
- Use the default options:
  - `CheckARM.exe -f C:\\Users\\User\\Desktop\\arms.txt`
  - `CheckARM.exe -i 192.168.0.1-192.168.0.100`

**Active Directory Mode:**
- Use the default options:
  - `CheckARM.exe -o OU=testou,DC=testdc`

**Create a Report Only:**
- `CheckARM.exe -ro`

## Requirements
- Windows environment with PowerShell 5.0 or higher.
- Administrative rights on the system where the script is executed.

## Report Creation

The Python script analyzes JSON files containing exported parameters of the computer being checked and generates a report in the .docx format. To prepare the script for execution, compile the file `./report_create/report_create.py` into `report_create.exe` and place it in the `./src/reporting/report_create.exe` directory.
Use `requirements.txt` for installing dependencies necessary for the script to run. This step ensures that all required Python libraries are installed in your environment, facilitating the correct functioning of the report creation script.

### How It Works

The script takes the exported JSON files as input, processes the contained information, and generates a comprehensive report detailing the analyzed computer's configuration and system status. This process involves comparing the computer's data against predefined criteria and highlighting significant findings in the .docx report.

For detailed information on how the script operates, including its architecture, input parameters, and customization options, please refer to the documentation available in `/report_create/README.md`.

### Compiling to Executable

To compile `report_create.py` into an executable file (`report_create.exe`), you can use a tool like PyInstaller. This step transforms the script into a standalone executable that can be run without requiring a Python interpreter on the target system. Here’s a simple command to do this using PyInstaller:

```shell
pyinstaller --onefile ./report_create/report_create.py
```

## Installation
Clone the repository to your local computer:

```shell
git clone https://github.com/kronigor/CheckARM.git
```

## License
This project is licensed under the MIT License. See the LICENSE file in the repository for more details.
