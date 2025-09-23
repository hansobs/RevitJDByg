"""
Export Material Data Script
This script exports material data from Revit to CSV format.
"""

__title__ = "Export Data"
__author__ = "Jens Damm & Hans Bohn Svendsen"

# Import required modules
from pyrevit import forms
import clr
import csv
import os
from datetime import datetime

# Add references to .NET assemblies
clr.AddReference('System.Windows.Forms')
clr.AddReference('System')

# Import .NET classes
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult, SaveFileDialog

# Import Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Get current Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

def get_material_data():
    """Collect material data from the Revit model"""
    material_data = []
    
    try:
        # Get all materials in the document
        materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
        
        for material in materials:
            # Get material properties
            material_info = {
                'Name': material.Name if material.Name else "Unnamed Material",
                'ID': material.Id.IntegerValue,
                'Class': material.MaterialClass if hasattr(material, 'MaterialClass') else "Unknown",
                'Category': material.MaterialCategory if hasattr(material, 'MaterialCategory') else "Unknown",
                'Color': get_material_color(material),
                'Shininess': material.Shininess if hasattr(material, 'Shininess') else "N/A",
                'Smoothness': material.Smoothness if hasattr(material, 'Smoothness') else "N/A",
                'Transparency': material.Transparency if hasattr(material, 'Transparency') else "N/A"
            }
            
            material_data.append(material_info)
            
    except Exception as e:
        raise Exception("Error collecting material data: {}".format(str(e)))
    
    return material_data

def get_material_color(material):
    """Get material color as RGB string"""
    try:
        if hasattr(material, 'Color'):
            color = material.Color
            return "RGB({}, {}, {})".format(color.Red, color.Green, color.Blue)
        else:
            return "No Color"
    except:
        return "Unknown"

def save_to_csv(material_data):
    """Save material data to CSV file"""
    try:
        # Create save file dialog
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        
        # Set default directory to C: drive
        save_dialog.InitialDirectory = r"C:\"
        
        # Set default filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_dialog.FileName = "Material_Export_{}".format(timestamp)
        
        # Show dialog
        if save_dialog.ShowDialog() == DialogResult.OK:
            file_path = save_dialog.FileName
            
            # Write CSV file (IronPython compatible way)
            with open(file_path, 'w') as csvfile:
                if material_data:
                    # Get field names from first material
                    fieldnames = material_data[0].keys()
                    
                    # Write header manually
                    header = ','.join(fieldnames)
                    csvfile.write(header + '\n')
                    
                    # Write data rows manually
                    for material in material_data:
                        row_values = []
                        for field in fieldnames:
                            value = str(material[field])
                            # Escape commas and quotes in CSV values
                            if ',' in value or '"' in value or '\n' in value:
                                value = '"' + value.replace('"', '""') + '"'
                            row_values.append(value)
                        csvfile.write(','.join(row_values) + '\n')
                else:
                    # Write empty file with headers
                    header = 'Name,ID,Class,Category,Color,Shininess,Smoothness,Transparency'
                    csvfile.write(header + '\n')
            
            return file_path, len(material_data)
        else:
            return None, 0
            
    except Exception as e:
        raise Exception("Error saving CSV file: {}".format(str(e)))

def main():
    """Main function that runs when the button is clicked"""
    try:
        # Show confirmation dialog
        result = MessageBox.Show(
            "This will generate a material list export in CSV format for calculation purposes.\n\nDo you want to proceed?",
            "Export Material Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        # Check user's response
        if result == DialogResult.Yes:
            # User clicked Yes - proceed with export
            
            # Show progress message
            MessageBox.Show(
                "Collecting material data from the model...\n\nThis may take a moment.",
                "Processing",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            
            # Collect material data
            material_data = get_material_data()
            
            if not material_data:
                MessageBox.Show(
                    "No materials found in the current model.",
                    "No Data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return
            
            # Save to CSV
            file_path, count = save_to_csv(material_data)
            
            if file_path:
                # Show success message
                MessageBox.Show(
                    "Export completed successfully!\n\n" +
                    "Materials exported: {}\n".format(count) +
                    "File saved to:\n{}".format(file_path),
                    "Export Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
            else:
                # User cancelled save dialog
                MessageBox.Show(
                    "Export cancelled - no file was saved.",
                    "Export Cancelled",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
        else:
            # User clicked No
            MessageBox.Show(
                "Export cancelled by user.",
                "Export Cancelled",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            
    except Exception as e:
        # Error handling
        MessageBox.Show(
            "An error occurred during export:\n\n{}".format(str(e)),
            "Export Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )

# This is the entry point when the button is clicked
if __name__ == '__main__':
    main()