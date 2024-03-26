# Python Script for Computer Parameters Analysis and DOCX Report Generation

This Python script is designed to analyze computer parameters from JSON files and generate a detailed report in DOCX format. It is suitable for IT professionals and system administrators who need to collect and review configurations of computers in their network.

## Features

- **Modular Design**: Each section of the report is generated by a dedicated function, allowing for easy customization and expansion.
- **Comprehensive Analysis**: The script covers various aspects of system configuration, including general information, operating system details, network configuration, time service, OS updates, user accounts, shared folders, antivirus protection, software inventory, and more.
- **Highlighting Changes**: It includes functionality to highlight outdated or non-standard configurations and software, making it easier to identify potential issues.

## Sections of the Report

The report is divided into the following sections:

1. **General Information**: Provides basic information about the computer, such as name, domain, and current user.
2. **Operating System**: Details about the operating system, including caption, architecture, and version.
3. **Network Configuration**: Information on network settings, including IP addresses, gateways, DNS servers, and MAC addresses.
4. **Time Service**: Status of the network time protocol service.
5. **OS Updates**: List of installed OS updates and hotfixes.
6. **User Accounts**: Overview of user accounts present on the system, highlighting any inactive or administrator accounts.
7. **Groups**: Lists user groups, indicating any non-standard groups.
8. **Shared Folders**: Information on shared folders.
9. **Antivirus Protection**: Details about the installed antivirus software, including name, version, and update status.
10. **Software**: Inventory of installed software, highlighting any software not included in a predefined list of standard applications.
11. **Password Policy**: Details of the system's password policy settings.
12. **Autostart Applications**: Lists applications that start automatically with the system.
13. **Task Scheduler**: Overview of scheduled tasks.
14. **System Software**: Details about the system and embedded software.
15. **Services**: Information on system services, highlighting any non-standard services.
16. **System Log**: Entries from the system log.
17. **Security Log**: Entries from the security log.

## Prerequisites

- Python 3.6 or higher.
- `docx` library for creating DOCX files.
- `fuzzywuzzy` library for string matching, useful in software inventory analysis.
- Access to the JSON files containing the computer parameters to be analyzed.

To install the required Python libraries, you can use the provided `requirements.txt` file. Open a terminal or command prompt and navigate to the directory containing `requirements.txt`. Then, run the following command:

```bash
pip install -r requirements.txt
```

## Usage

This section explains how to use the script with examples of command-line arguments.


1. Ensure all prerequisites are installed and the JSON files are prepared.
2. Run the script with the name of the computer as an argument and specify the folder path for output.
3. The script will generate a DOCX report in the specified folder.

### Arguments

- `-c`, `--computer`: Specifies the name of the computer for which the report is to be generated. This argument is required.
- `-p`, `--path`: Specifies the root directory where the CheckARM.exe executable file is located. This argument is required.
- `-rv`, `--report_version`: Determines the version of the report to generate. Can be either `short` or `full`. The script selects the appropriate report template and determines the number of blocks to generate based on this version. This argument is required.
- `-ov`, `--os_version`: Specifies the default operating system version. In the "Operating System" section of the report, the script compares this specified version against the version of the OS of the computer being checked. This argument is required.
- `-r`, `--ratio`: Ratio used for fuzzy string matching of software names by the FuzzyWuzzy library. Defaults to 90.

## Examples

To generate a full report for a computer named PC1 located in the root directory C:\CheckARM\CheckARM.exe, where the default OS version is 10.0.18363 and using a fuzzy matching ratio of 80 for software name comparisons, you can use the following command:

```shell
python report_create.py -c PC1 -p C:\CheckARM\CheckARM.exe -rv full -bv 10.0.18363 -r 80
```
For a short report version, you can simply change the -rv argument to short:

```shell
python report_create.py -c PC1 -p C:\CheckARM\CheckARM.exe -rv short -bv 10.0.18363 -r 80
```

These commands will analyze the computer's parameters from the JSON files located in the specified path and generate a report in DOCX format according to the provided arguments.
This block provides a concise explanation of how to use the script, including the purpose of each argument and examples of how to execute the script for different scenarios.

## Customization

The script can be customized by modifying the functions corresponding to each report section or by adding new sections as needed. The `sections_dict` dictionary maps section numbers to their respective functions and file paths, making it easy to add or modify sections. You can tailor the analysis and reporting to meet specific needs by adjusting the existing sections or incorporating new ones for additional insights.

## License

This project is licensed under the MIT License - see the LICENSE file for details. The MIT License is a permissive free software license originating at the Massachusetts Institute of Technology (MIT). It allows for reuse within proprietary software provided that all copies of the licensed software include a copy of the MIT License terms.