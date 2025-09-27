"""
Export Material Data Script for ESG/CO2 Analysis
This script exports comprehensive material data from Revit to CSV format.
"""
__title__ = "Export ESG Data"
__author__ = "Jens Damm & Hans Bohn Svendsen"

# Import required modules
from pyrevit import forms
from pyrevit.framework import List
from pyrevit.revit import HOST_APP
from pyrevit.output import get_output
import clr
import csv
import os
import time
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

# Get output window for progress updates
output = get_output()

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

def get_material_layers(element):
    """Returns all layers with material + thickness, including 'By Category'"""
    results = []
    try:
        element_type = doc.GetElement(element.GetTypeId())
        if hasattr(element_type, 'GetCompoundStructure'):
            cs = element_type.GetCompoundStructure()
            if cs:
                layers = cs.GetLayers()
                for i, layer in enumerate(layers):
                    mat_id = layer.MaterialId
                    if mat_id == ElementId.InvalidElementId:
                        # By Category create label with category
                        mat_name = "By Category ({})".format(element.Category.Name)
                        mat_id_for_export = "By_Category"
                    else:
                        mat = doc.GetElement(mat_id)
                        mat_name = mat.Name if mat else "Unknown"
                        mat_id_for_export = mat_id.IntegerValue
                    
                    thickness_mm = UnitUtils.ConvertFromInternalUnits(
                        layer.Width, UnitTypeId.Millimeters
                    )
                    
                    results.append({
                        "LayerIndex": i,
                        "MaterialId": mat_id_for_export,
                        "MaterialName": mat_name,
                        "Thickness_mm": round(thickness_mm, 2)
                    })
            else:
                # fallback: try to get thickness parameter directly
                t_param = element.LookupParameter("Thickness")
                if t_param and t_param.HasValue:
                    thickness_mm = UnitUtils.ConvertFromInternalUnits(
                        t_param.AsDouble(), UnitTypeId.Millimeters
                    )
                    results.append({
                        "LayerIndex": 0,
                        "MaterialId": "N/A",
                        "MaterialName": "N/A",
                        "Thickness_mm": round(thickness_mm, 2)
                    })
    except Exception as e:
        results.append({
            "LayerIndex": -1,
            "MaterialId": "Error",
            "MaterialName": "Error",
            "Thickness_mm": str(e)
        })
    
    return results

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
        total_elements = len(elements)
        
        output.print_md("## Starting Material Data Collection")
        output.print_md("**Total elements to process:** {}".format(total_elements))
        
        # Initialize progress tracking
        processed_elements = 0
        material_records = 0
        
        # Process elements with progress updates
        for i, element in enumerate(elements):
            try:
                # Update progress every 50 elements or at milestones
                if i % 50 == 0 or i == total_elements - 1:
                    progress_percent = int((i + 1) * 100 / total_elements)
                    output.update_progress(i + 1, total_elements)
                    print("Processing element {} of {} ({}%) - {} material records so far".format(
                        i + 1, total_elements, progress_percent, material_records))
                
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
                # Get parameters
                element_width = get_element_width(element)
                element_height = get_element_height(element)
                export_guid = get_export_guid(element)
                
                # NEW: Get material layers (handles By Category)
                material_layers = get_material_layers(element)
                
                # If no layers found, try the old method as fallback
                if not material_layers:
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
                    
                    # Process old way for non-compound elements
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
                                'ExportGUID': export_guid,
                                # Family hierarchy
                                'FamilyName': family_name,
                                'FamilyType': family_type,
                                'Type': element_type,
                                'TypeId': element.GetTypeId().IntegerValue,
                                # Dimension parameters
                                'Width_mm': element_width,
                                'Height_mm': element_height,
                                # Layer information
                                'LayerIndex': 0,
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
                            material_records += 1
                else:
                    # NEW: Process using material layers
                    for layer in material_layers:
                        # Calculate material-specific volume and area for this layer
                        if layer['MaterialId'] != "By_Category" and layer['MaterialId'] != "N/A" and layer['MaterialId'] != "Error":
                            try:
                                material_id = ElementId(int(layer['MaterialId']))
                                material = doc.GetElement(material_id)
                                material_class = material.MaterialClass if material and hasattr(material, 'MaterialClass') else "Unknown"
                            except:
                                material = None
                                material_class = "Unknown"
                        else:
                            material = None
                            material_class = "By Category" if layer['MaterialId'] == "By_Category" else "Unknown"
                        
                        # Calculate volume for this specific layer
                        thickness = layer['Thickness_mm']
                        material_volume = calculate_layer_volume(element, thickness)
                        material_area = calculate_material_area(element, None)  # Use element area
                        
                        # FORMAT ALL NUMERICAL VALUES PROPERLY
                        material_info = {
                            # Element identification
                            'ElementId': element.Id.IntegerValue,
                            'ElementCategory': element.Category.Name if element.Category else "Unknown",
                            'ExportGUID': export_guid,
                            # Family hierarchy
                            'FamilyName': family_name,
                            'FamilyType': family_type,
                            'Type': element_type,
                            'TypeId': element.GetTypeId().IntegerValue,
                            # Dimension parameters
                            'Width_mm': element_width,
                            'Height_mm': element_height,
                            # Layer information
                            'LayerIndex': layer['LayerIndex'],
                            # Material information
                            'MaterialId': layer['MaterialId'],
                            'MaterialName': layer['MaterialName'],
                            'MaterialClass': material_class,
                            # Thickness and quantities - PROPERLY FORMATTED
                            'Thickness_mm': format_number(thickness, 5),
                            'MaterialVolume_m3': format_number(material_volume, 5),
                            'MaterialArea_m2': format_number(material_area, 5),
                            'ElementTotalVolume_m3': format_number(element_volume, 5),
                            'ElementTotalArea_m2': format_number(element_area, 5)
                        }
                        
                        material_usage_data.append(material_info)
                        material_records += 1
                
                processed_elements += 1
                
            except Exception as e:
                # Skip problematic elements but continue processing
                continue
        
        # Final progress update
        output.print_md("## Collection Complete!")
        output.print_md("**Elements processed:** {}".format(processed_elements))
        output.print_md("**Material records created:** {}".format(material_records))
        
    except Exception as e:
        raise Exception("Error collecting comprehensive material data: {}".format(str(e)))
    
    return material_usage_data

def calculate_layer_volume(element, thickness_mm):
    """Calculate volume for a specific layer thickness"""
    try:
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType'):
            area_param = element.LookupParameter("Area")
            if area_param and area_param.HasValue and thickness_mm != "N/A":
                area_sqft = area_param.AsDouble()
                thickness_ft = UnitUtils.ConvertToInternalUnits(float(thickness_mm), UnitTypeId.Millimeters)
                volume_cuft = area_sqft * thickness_ft
                volume_cum = UnitUtils.ConvertFromInternalUnits(volume_cuft, UnitTypeId.CubicMeters)
                return round(volume_cum, 5)
        return "N/A"
    except:
        return "N/A"


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

def get_element_width(element):
    """Get Width parameter from element - clean output without debug tags"""
    try:
        # Method 1: Try "Width" parameter on INSTANCE first
        w_param = element.LookupParameter("Width")
        if w_param and w_param.HasValue:
            width_mm = UnitUtils.ConvertFromInternalUnits(w_param.AsDouble(), UnitTypeId.Millimeters)
            return format_number(width_mm, 5)
        
        # Method 2: Try "Width" parameter on TYPE
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            w_param = element_type.LookupParameter("Width")
            if w_param and w_param.HasValue:
                width_mm = UnitUtils.ConvertFromInternalUnits(w_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(width_mm, 5)
        
        # Method 3: Try built-in DOOR_WIDTH parameter (backup)
        try:
            width_param = element.get_Parameter(BuiltInParameter.DOOR_WIDTH)
            if width_param and width_param.HasValue:
                width_mm = UnitUtils.ConvertFromInternalUnits(width_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(width_mm, 5)
        except:
            pass
        
        # Method 4: Try built-in WINDOW_WIDTH parameter (backup)
        try:
            width_param = element.get_Parameter(BuiltInParameter.WINDOW_WIDTH)
            if width_param and width_param.HasValue:
                width_mm = UnitUtils.ConvertFromInternalUnits(width_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(width_mm, 5)
        except:
            pass
        
        return "N/A"
    except:
        return "N/A"

def get_element_height(element):
    """Get Height parameter from element - clean output without debug tags"""
    try:
        # Method 1: Try "Height" parameter on INSTANCE first
        h_param = element.LookupParameter("Height")
        if h_param and h_param.HasValue:
            height_mm = UnitUtils.ConvertFromInternalUnits(h_param.AsDouble(), UnitTypeId.Millimeters)
            return format_number(height_mm, 5)
        
        # Method 2: Try "Height" parameter on TYPE
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            h_param = element_type.LookupParameter("Height")
            if h_param and h_param.HasValue:
                height_mm = UnitUtils.ConvertFromInternalUnits(h_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(height_mm, 5)
        
        # Method 3: Try built-in DOOR_HEIGHT parameter (backup)
        try:
            height_param = element.get_Parameter(BuiltInParameter.DOOR_HEIGHT)
            if height_param and height_param.HasValue:
                height_mm = UnitUtils.ConvertFromInternalUnits(height_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(height_mm, 5)
        except:
            pass
        
        # Method 4: Try built-in WINDOW_HEIGHT parameter (backup)
        try:
            height_param = element.get_Parameter(BuiltInParameter.WINDOW_HEIGHT)
            if height_param and height_param.HasValue:
                height_mm = UnitUtils.ConvertFromInternalUnits(height_param.AsDouble(), UnitTypeId.Millimeters)
                return format_number(height_mm, 5)
        except:
            pass
        
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
            # Show progress for CSV writing
            output.print_md("## Writing CSV File...")
            print("Writing {} material records to CSV...".format(len(material_data)))
            with open(file_path, 'wb') as csvfile:
                if material_data:
                    fieldnames = material_data[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames,
                                          delimiter=';', lineterminator='\n')
                    writer.writeheader()
                    # Write data with progress updates
                    for i, material in enumerate(material_data):
                        if i % 500 == 0:  # Update every 500 records
                            progress_percent = int((i + 1) * 100 / len(material_data))
                            print("Writing record {} of {} ({}%)".format(i + 1, len(material_data), progress_percent))
                        writer.writerow(material)
                else:
                    writer = csv.writer(csvfile, delimiter=';', lineterminator='\n')
                    headers = [
                        'ElementId', 'ElementCategory', 'ExportGUID', 'FamilyName', 'FamilyType', 'Type', 'TypeId',
                        'Width_mm', 'Height_mm', 'LayerIndex',
                        'MaterialId', 'MaterialName', 'MaterialClass',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2',
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2'
                    ]  # FIXED: Added missing closing bracket
                    writer.writerow(headers)
            output.print_md("## CSV Export Complete!")
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
            # Start timer
            start_time = time.time()
            # Clear output window and show progress
            output.close_others()
            output.print_md("# ESG Material Data Export")
            output.print_md("**Started:** {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            # Use the new comprehensive function with progress
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
            # Calculate elapsed time
            end_time = time.time()
            elapsed_time = round(end_time - start_time, 2)
            if file_path:
                output.print_md("## Export Results")
                output.print_md("**File:** {}".format(file_path))
                output.print_md("**Records:** {}".format(count))
                output.print_md("**Time:** {} seconds".format(elapsed_time))
                MessageBox.Show(
                    "Export completed successfully!\n\n" +
                    "Material records exported: {}\n".format(count) +
                    "Processing time: {} seconds\n".format(elapsed_time) +
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
