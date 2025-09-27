## to do next time/what to test:
- **download new version
- **save it in the extension folder
- **reload pyrevit
- **on projekt1 press the new updated extension to export csv

- **it should by default suggest the c drive as the saving location instead of the location of the extension
- **when choosing a location, the csv export should work - last test threw an error during the csv export.

## Future ideas:

- able to upload a csv from an excel sheet with headers - the code should try and recognize the columns for what is what, and do a suggestion, you can manually overwrite if you dont agree, and from this, you can calculate the esg - should also in a user friendly way let the user know if a set of information is missing, blocking it from calculating the emissions



# Material Export Extension

A pyRevit extension that exports comprehensive material data from Revit models to CSV format for quantity calculations and analysis.

## 📋 Features

- **System Families Export**: Walls, Floors, Roofs, and Ceilings with layer information
- **Loadable Families Export**: All family instances with available data
- **Detailed Material Information**: Including material names, thicknesses, areas, and volumes
- **CSV Output**: Semicolon-delimited format compatible with Excel and other analysis tools
- **User-Friendly Interface**: Confirmation dialogs and progress feedback

## 🛠 Prerequisites

Before installing this extension, ensure you have:

- ✅ **Autodesk Revit** installed (2019 or later recommended)
- ✅ **pyRevit extension** installed ([Download here](https://github.com/eirannejad/pyRevit/releases))

## 📥 Installation Guide

### Step 1: Download the Extension

1. Navigate to this GitHub repository
2. Click the **"Code"** button and select **"Download ZIP"**
3. Extract the downloaded ZIP file to a temporary location
4. Locate the `pyrevit_export_materials.extension` folder

### Step 2: Install the Extension

#### Option A: For Jens (Specific User Path)
Move the `.extension` folder to:
C:\Users\JENDAM\AppData\Roaming\pyRevit-Master\extensions


### Step 3: Configure pyRevit

1. **Open Revit** and navigate to the **pyRevit** tab
2. Click **"Settings"** in the pyRevit ribbon
3. In the Settings dialog, go to **"Custom Extension Directories"**
4. Click **"Add Folder"** and add the following path:
C:\Users\JENDAM\AppData\Roaming\pyRevit-Master\extensions\pyrevit_export_materials.extension
*(Replace `JENDAM` with your username if different)*
5. Click **"Save Settings"**
6. Click **"Reload"** to refresh pyRevit

### Step 4: Verify Installation

After reloading, you should see a new tab called **"Jens D Data Export"** in the Revit ribbon.

## 🚀 Usage

### Basic Usage

1. **Open your Revit model** with the elements you want to export
2. Navigate to the **"Jens D Data Export"** tab
3. Click the **"Export Material Data"** button
4. Click **"Yes"** in the confirmation dialog
5. Wait for the export to complete
6. The CSV file will be saved to your **Documents** folder as `RevitFamilies.csv`

### Output Format

The exported CSV file contains the following columns:

| Column | Description |
|--------|-------------|
| `ElementId` | Unique Revit element identifier |
| `Category` | Revit category (Walls, Floors, etc.) |
| `Family` | Family name |
| `Type` | Type name |
| `Material` | Material name for each layer |
| `Thickness (mm)` | Layer thickness in millimeters |
| `Area (m2)` | Element area in square meters |
| `Volume (m3)` | Element volume in cubic meters |

### Example Output

```csv
ElementId;Category;Family;Type;Material;Thickness (mm);Area (m2);Volume (m3)
12345;Walls;Basic Wall;Generic - 200mm;Concrete;150.0;25.5;3.83
12345;Walls;Basic Wall;Generic - 200mm;Insulation;50.0;25.5;1.28
```


## 🔧 Troubleshooting

### Extension Not Appearing
- Verify the extension folder is in the correct location
- Check that the path in pyRevit settings matches your installation
- Try reloading pyRevit or restarting Revit

### Export Errors
- Ensure you have an active Revit document open
- Check that you have write permissions to your Documents folder
- Verify that elements exist in your model

### File Access Issues
- Close any Excel files that might have the output CSV open
- Check your antivirus software isn't blocking file creation

## 📁 File Structure
```
pyrevit_export_materials.extension/
├── Jens D Data Export.tab/
│   └── Export Tools.panel/
│       └── Export Material Data.pushbutton/
│           ├── script.py
│           └── icon.png
└── README.md
```
## 🤝 Support

For issues, questions, or feature requests:

1. Check the troubleshooting section above
2. Review the pyRevit documentation
3. Contact the development team

## 🔄 Version History

- **v1.0.0** - Initial release with basic material export functionality

---

**Author**: Jens Damm & Hans BS  
**Created**: 2025  
**pyRevit Version**: Compatible with pyRevit 5.2