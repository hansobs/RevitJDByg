"""
Export Material Data Script for ESG/CO2 Analysis
This script exports comprehensive material data from Revit to CSV format.
"""
__title__ = "Export ESG Data"
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

def format_number(value, decimals=5):
    """Format a number to a specific number of decimal places, removing trailing zeros"""
    if value == "N/A" or value is None:
        return "N/A"
    try:
        # Round to specified decimals and format as string
        formatted = "{:.{}f}".format(float(value), decimals)
        # Remove trailing zeros and unnecessary decimal point
        formatted = formatted.rstrip('0').rstrip('.')
        if formatted == "":
            formatted = "0"
        # Replace dot with comma for European Excel compatibility
        formatted = formatted.replace('.', ',')
        return formatted
    except:
        return "N/A"


def get_element_type_name(element):
    """Get the element type name"""
    try:
        # Debug: Check if element has TypeId
        type_id = element.GetTypeId()
        if type_id == ElementId.InvalidElementId:
            return "No TypeId"
        
        # Method 1: Direct type lookup
        element_type = doc.GetElement(type_id)
        if element_type:
            # Try different ways to get the name
            if hasattr(element_type, 'Name'):
                name = element_type.Name
                if name and name.strip():  # Check if name exists and is not empty
                    return name
            
            # Try get_Name() method if Name property doesn't work
            if hasattr(element_type, 'get_Name'):
                try:
                    name = element_type.get_Name()
                    if name and name.strip():
                        return name
                except:
                    pass
            
            # Try LookupParameter for "Type Name"
            try:
                type_name_param = element_type.LookupParameter("Type Name")
                if type_name_param and type_name_param.HasValue:
                    return type_name_param.AsString()
            except:
                pass
            
            # Try to get the element type's category and ID as fallback
            try:
                if element_type.Category:
                    return "{} - ID {}".format(element_type.Category.Name, element_type.Id.IntegerValue)
                else:
                    return "Type ID {}".format(element_type.Id.IntegerValue)
            except:
                pass
        
        # Method 2: Try getting from Symbol (for family instances)
        if hasattr(element, 'Symbol') and element.Symbol:
            return element.Symbol.Name
        
        # Method 3: Try specific element type properties
        if hasattr(element, 'WallType') and element.WallType:
            return element.WallType.Name
        elif hasattr(element, 'FloorType') and element.FloorType:
            return element.FloorType.Name
        elif hasattr(element, 'RoofType') and element.RoofType:
            return element.RoofType.Name
        
        return "No type found"
        
    except Exception as e:
        return "Exception: {}".format(str(e))

def get_comprehensive_material_data():
    material_usage_data = []
    try:
        # Get all elements that have materials
        elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        for element in elements:
            try:
                # Skip elements without geometry or materials
                if not element.Category:
                    continue
                # Get family and type information
                family_name = get_family_name(element)
                family_type = get_family_type(element)
                element_type = get_element_type_name(element)
                # Get element geometry for area/volume calculations
                element_volume = get_element_volume(element)
                element_area = get_element_area(element)
                # Get NEW parameters
                door_width = get_door_width(element)
                door_height = get_door_height(element)
                export_guid = get_export_guid(element)
                # Get materials from the element
                material_ids = get_element_material_ids(element)
                if not material_ids:
                    # If no specific materials found, try to get from element type
                    element_type_obj = doc.GetElement(element.GetTypeId())
                    if element_type_obj:
                        type_material_id = element_type_obj.LookupParameter("Material")
                        if type_material_id and type_material_id.HasValue:
                            mat_id = type_material_id.AsElementId()
                            if mat_id != ElementId.InvalidElementId:
                                material_ids = [mat_id]
                for material_id in material_ids:
                    material = doc.GetElement(material_id)
                    if material:
                        # Get thickness (layer-specific for walls, floors, etc.)
                        thickness = get_material_thickness(element, material_id)
                        # Calculate material-specific volume and area
                        material_volume = calculate_material_volume(element, material_id, thickness)
                        material_area = calculate_material_area(element, material_id)
                        
                        # FORMAT ALL NUMERICAL VALUES PROPERLY
                        material_info = {
                            # Element identification
                            'ElementId': element.Id.IntegerValue,
                            'ElementCategory': element.Category.Name if element.Category else "Unknown",
                            'ExportGUID': export_guid,  # NEW
                            # Family hierarchy
                            'FamilyName': family_name,
                            'FamilyType': family_type,
                            'Type': element_type,
                            'TypeId': element.GetTypeId().IntegerValue,
                            # NEW door parameters
                            'DoorWidth_mm': door_width,  # NEW
                            'DoorHeight_mm': door_height,  # NEW
                            # Material information
                            'MaterialId': material_id.IntegerValue,
                            'MaterialName': material.Name if material.Name else "Unnamed Material",
                            'MaterialClass': material.MaterialClass if hasattr(material, 'MaterialClass') else "Unknown",
                            # Thickness and quantities - PROPERLY FORMATTED
                            'Thickness_mm': format_number(thickness, 5),
                            'MaterialVolume_m3': format_number(material_volume, 5),
                            'MaterialArea_m2': format_number(material_area, 5),
                            'ElementTotalVolume_m3': format_number(element_volume, 5),
                            'ElementTotalArea_m2': format_number(element_area, 5)
                        }
                        
                        material_usage_data.append(material_info)
            except Exception as e:
                # Skip problematic elements but continue processing
                continue
    except Exception as e:
        raise Exception("Error collecting comprehensive material data: {}".format(str(e)))
    return material_usage_data

def get_family_name(element):
    """Get family name from element"""
    try:
        if hasattr(element, 'Symbol') and element.Symbol:
            return element.Symbol.Family.Name
        elif hasattr(element, 'WallType'):
            return "Wall"
        elif hasattr(element, 'FloorType'):
            return "Floor"
        elif hasattr(element, 'RoofType'):
            return "Roof"
        else:
            return element.Category.Name if element.Category else "Unknown"
    except:
        return "Unknown"

def get_family_type(element):
    """Get family type name from element using 'Family and Type' parameter"""
    try:
        # First try to get "Family and Type" parameter
        family_type_param = element.LookupParameter("Family and Type")
        if family_type_param and family_type_param.HasValue:
            return family_type_param.AsValueString()
        # Fallback methods if "Family and Type" parameter doesn't exist
        if hasattr(element, 'Symbol') and element.Symbol:
            return element.Symbol.Name
        elif hasattr(element, 'WallType'):
            return element.WallType.Name
        elif hasattr(element, 'FloorType'):
            return element.FloorType.Name
        elif hasattr(element, 'RoofType'):
            return element.RoofType.Name
        else:
            element_type = doc.GetElement(element.GetTypeId())
            return element_type.Name if element_type else "Unknown"
    except:
        return "Unknown"

def get_element_material_ids(element):
    """Get all material IDs used in an element"""
    material_ids = []
    try:
        # For walls, floors, roofs with compound structures
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType') or hasattr(element, 'RoofType'):
            element_type = doc.GetElement(element.GetTypeId())
            if hasattr(element_type, 'GetCompoundStructure'):
                compound_structure = element_type.GetCompoundStructure()
                if compound_structure:
                    layers = compound_structure.GetLayers()
                    for layer in layers:
                        if layer.MaterialId != ElementId.InvalidElementId:
                            material_ids.append(layer.MaterialId)
        # For other elements, get material from geometry or parameters
        try:
            element_material_ids = element.GetMaterialIds(False)
            material_ids.extend(element_material_ids)
        except:
            pass
        # Remove duplicates and invalid IDs
        valid_ids = []
        for mat_id in material_ids:
            if mat_id != ElementId.InvalidElementId and mat_id not in valid_ids:
                valid_ids.append(mat_id)
        return valid_ids
    except:
        return []

def get_material_thickness(element, material_id):
    """Get thickness of specific material in element"""
    try:
        # For compound structures (walls, floors, roofs)
        element_type = doc.GetElement(element.GetTypeId())
        if hasattr(element_type, 'GetCompoundStructure'):
            compound_structure = element_type.GetCompoundStructure()
            if compound_structure:
                layers = compound_structure.GetLayers()
                for layer in layers:
                    if layer.MaterialId == material_id:
                        # Convert from internal units to mm
                        thickness_feet = layer.Width
                        thickness_mm = UnitUtils.ConvertFromInternalUnits(thickness_feet, UnitTypeId.Millimeters)
                        return round(thickness_mm, 5)
        # For other elements, try to get thickness parameter
        thickness_param = element.LookupParameter("Thickness")
        if thickness_param and thickness_param.HasValue:
            thickness_mm = UnitUtils.ConvertFromInternalUnits(thickness_param.AsDouble(), UnitTypeId.Millimeters)
            return round(thickness_mm, 5)
        return "N/A"
    except:
        return "N/A"

def calculate_material_volume(element, material_id, thickness):
    """Calculate volume of specific material in element"""
    try:
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType'):
            area_param = element.LookupParameter("Area")
            if area_param and area_param.HasValue and thickness != "N/A":
                area_sqft = area_param.AsDouble()
                thickness_ft = UnitUtils.ConvertToInternalUnits(float(thickness), UnitTypeId.Millimeters)
                volume_cuft = area_sqft * thickness_ft
                volume_cum = UnitUtils.ConvertFromInternalUnits(volume_cuft, UnitTypeId.CubicMeters)
                return round(volume_cum, 5)
        # For other elements, use total volume
        volume_param = element.LookupParameter("Volume")
        if volume_param and volume_param.HasValue:
            volume_cuft = volume_param.AsDouble()
            volume_cum = UnitUtils.ConvertFromInternalUnits(volume_cuft, UnitTypeId.CubicMeters)
            return round(volume_cum, 5)
        return "N/A"
    except:
        return "N/A"

def calculate_material_area(element, material_id):
    """Calculate area of specific material in element"""
    try:
        area_param = element.LookupParameter("Area")
        if area_param and area_param.HasValue:
            area_sqft = area_param.AsDouble()
            area_sqm = UnitUtils.ConvertFromInternalUnits(area_sqft, UnitTypeId.SquareMeters)
            return round(area_sqm, 5)
        return "N/A"
    except:
        return "N/A"

def get_element_volume(element):
    """Get total volume of element"""
    try:
        volume_param = element.LookupParameter("Volume")
        if volume_param and volume_param.HasValue:
            volume_cuft = volume_param.AsDouble()
            volume_cum = UnitUtils.ConvertFromInternalUnits(volume_cuft, UnitTypeId.CubicMeters)
            return round(volume_cum, 5)
        return "N/A"
    except:
        return "N/A"

def get_element_area(element):
    """Get total area of element"""
    try:
        area_param = element.LookupParameter("Area")
        if area_param and area_param.HasValue:
            area_sqft = area_param.AsDouble()
            area_sqm = UnitUtils.ConvertFromInternalUnits(area_sqft, UnitTypeId.SquareMeters)
            return round(area_sqm, 5)
        return "N/A"
    except:
        return "N/A"
    
def get_door_width(element):
    """Get door width parameter"""
    try:
        width_param = element.get_Parameter(BuiltInParameter.DOOR_WIDTH)
        if width_param and width_param.HasValue:
            width_feet = width_param.AsDouble()
            width_mm = UnitUtils.ConvertFromInternalUnits(width_feet, UnitTypeId.Millimeters)
            return format_number(width_mm, 5)
        return "N/A"
    except:
        return "N/A"

def get_door_height(element):
    """Get door height parameter"""
    try:
        height_param = element.get_Parameter(BuiltInParameter.DOOR_HEIGHT)
        if height_param and height_param.HasValue:
            height_feet = height_param.AsDouble()
            height_mm = UnitUtils.ConvertFromInternalUnits(height_feet, UnitTypeId.Millimeters)
            return format_number(height_mm, 5)
        return "N/A"
    except:
        return "N/A"

def get_export_guid(element):
    """Get export GUID for element"""
    try:
        guid_str = ExportUtils.GetExportId(doc, element.Id)
        return str(guid_str) if guid_str else "N/A"
    except:
        return "N/A"

def save_to_csv(material_data):
    """Save comprehensive material data to CSV file with semicolon delimiter"""
    try:
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        save_dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        # Updated timestamp format: yyyymmdd hhmm
        timestamp = datetime.now().strftime("%Y%m%d %H%M")
        save_dialog.FileName = "ESG_Material_Export_{}".format(timestamp)
        if save_dialog.ShowDialog() == DialogResult.OK:
            file_path = save_dialog.FileName
            with open(file_path, 'wb') as csvfile:
                if material_data:
                    fieldnames = material_data[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames,
                                          delimiter=';', lineterminator='\n')
                    writer.writeheader()
                    for material in material_data:
                        writer.writerow(material)
                else:
                    writer = csv.writer(csvfile, delimiter=';', lineterminator='\n')
                    # Updated headers to include 'Type'
                    headers = [
                        'ElementId', 'ElementCategory', 'ExportGUID', 'FamilyName', 'FamilyType', 'Type', 'TypeId',
                        'DoorWidth_mm', 'DoorHeight_mm',
                        'MaterialId', 'MaterialName', 'MaterialClass',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2',
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2'
                    ]
                    writer.writerow(headers)
            return file_path, len(material_data)
        else:
            return None, 0
    except Exception as e:
        raise Exception("Error saving CSV file: {}".format(str(e)))

def main():
    """Main function that runs when the button is clicked"""
    try:
        result = MessageBox.Show(
            "This will generate a comprehensive material list export for ESG/CO2 analysis.\n\nDo you want to proceed?",
            "Export ESG Material Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        if result == DialogResult.Yes:
            MessageBox.Show(
                "Collecting comprehensive material data from the model...\n\nThis may take a moment.",
                "Processing",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            # Use the new comprehensive function
            material_data = get_comprehensive_material_data()
            if not material_data:
                MessageBox.Show(
                    "No materials found in the current model.",
                    "No Data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return
            file_path, count = save_to_csv(material_data)
            if file_path:
                MessageBox.Show(
                    "Export completed successfully!\n\n" +
                    "Material records exported: {}\n".format(count) +
                    "File saved to:\n{}".format(file_path),
                    "Export Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
            else:
                MessageBox.Show(
                    "Export cancelled - no file was saved.",
                    "Export Cancelled",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
        else:
            MessageBox.Show(
                "Export cancelled by user.",
                "Export Cancelled",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
    except Exception as e:
        MessageBox.Show(
            "An error occurred during export:\n\n{}".format(str(e)),
            "Export Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )

if __name__ == '__main__':
    main()
