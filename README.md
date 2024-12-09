# XML2Excel: XML to Excel Conversion Utility

## Overview

**XML2Excel** is a robust C# utility that converts XML files into Excel spreadsheets. This application monitors a specified folder for incoming XML files, allows batch conversion of existing files, and performs continuous file system monitoring. The tool is highly flexible and can handle complex XML structures by flattening them into an Excel-compatible format.

---

## Features

- **Batch Conversion**: Converts all existing XML files in the monitored folder at startup (optional).
- **File Monitoring**: Continuously monitors a folder for new XML files and converts them automatically.
- **Customizable Output**: Saves converted Excel files in a specified destination folder.
- **XML Flexibility**: Handles complex XML structures, flattening nested data for structured Excel sheets.
- **Seamless Setup**: Automatically ensures required folders exist.

---

## Installation and Setup

### Release Installation
Simply download from the [Latest release](https://github.com/yassir56069/XML2Excel/releases/latest) from the Releases tab on the right! 

### Manual Installation

1. **Clone or Download the Repository**  
   Clone this repository or download the ZIP file and extract it.

2. **Install Required Dependencies**  
   This project uses [EPPlus](https://www.epplussoftware.com/) for Excel file creation. Install it via NuGet:  
   ```
   Install-Package EPPlus -Version 5.8.0
   ```

3. **Compile the Project**  
   Open the project in Visual Studio or your preferred IDE, restore dependencies, and build the solution.

4. **Run the Application**  
   - The application creates a default folder structure:
     - **Input Folder**: `Documents/XMLInput`
     - **Output Folder**: `Documents/ConvertedXML`
   - Place your XML files in the input folder.
   - Run the executable to start monitoring and converting files.

---

## How It Works

1. **Initial Conversion**  
   When launched, the application scans the input folder (`XMLInput`) for existing XML files. You can choose to convert these files immediately.

2. **Continuous Monitoring**  
   A file watcher monitors the `XMLInput` folder for new XML files. As soon as a file is added, the application converts it to Excel and saves the output in the `ConvertedXML` folder.

3. **Output Organization**  
   - Converted Excel files are saved in `ConvertedXML`.
   - Initial batch-converted files are saved in a subfolder: `ConvertedXML/InitialFiles`.

---

## Folder Structure

- **Input Folder**: 
  - Default: `Documents/XMLInput`
  - Monitored for new XML files.
  
- **Output Folder**: 
  - Default: `Documents/ConvertedXML`
  - Stores converted Excel files.
  - Contains a subfolder `InitialFiles` for batch-converted files.

---

## Example

### Input: XML File

```xml
<root>
  <person id="1">
    <name>John Doe</name>
    <email>john.doe@example.com</email>
  </person>
  <person id="2">
    <name>Jane Smith</name>
    <email>jane.smith@example.com</email>
  </person>
</root>
```

### Output: Excel File

| ID | Name       | Email                |
|----|------------|----------------------|
| 1  | John Doe   | john.doe@example.com |
| 2  | Jane Smith | jane.smith@example.com |

---

## Requirements

- **.NET Framework/Runtime**: .NET Core 3.1 or higher.
- **EPPlus**: Used for Excel file creation.
- **Operating System**: Windows (tested on Windows 10).

---

## Usage

1. Run the application from the command line or your IDE.
2. Monitor the console for status updates.
3. Place XML files in the `XMLInput` folder for conversion.
4. Press **`Q`** in the console to quit the application.

---
