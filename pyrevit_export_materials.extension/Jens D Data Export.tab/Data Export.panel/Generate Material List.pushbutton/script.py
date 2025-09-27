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

def get_comprehensive_material_data():
    """Collect comprehensive family, type, and material data for ESG/CO2 analysis"""
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
                
                # Get element geometry for area/volume calculations
                element_volume = get_element_volume(element)
                element_area = get_element_area(element)
                
                # Get materials from the element
                material_ids = get_element_material_ids(element)
                
                if not material_ids:
                    # If no specific materials found, try to get from element type
                    element_type = doc.GetElement(element.GetTypeId())
                    if element_type:
                        type_material_id = element_type.LookupParameter("Material")
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
                        
                        material_info = {
                            # Element identification
                            'ElementId': element.Id.IntegerValue,
                            'ElementCategory': element.Category.Name if element.Category else "Unknown",
                            
                            # Family hierarchy
                            'FamilyName': family_name,
                            'FamilyType': family_type,
                            'TypeId': element.GetTypeId().IntegerValue,
                            
                            # Material information
                            'MaterialId': material_id.IntegerValue,
                            'MaterialName': material.Name if material.Name else "Unnamed Material",
                            'MaterialClass': material.MaterialClass if hasattr(material, 'MaterialClass') else "Unknown",
                            
                            # Thickness and quantities - the core data for CO2 calculations
                            'Thickness_mm': thickness,
                            'MaterialVolume_m3': material_volume,
                            'MaterialArea_m2': material_area,
                            'ElementTotalVolume_m3': element_volume,
                            'ElementTotalArea_m2': element_area
                        }
                        
                        material_usage_data.append(material_info)
                        
            except Exception as e:
                # Skip problematic elements but continue processing
                continue
                
    except Exception as e:
        raise Exception("Error collecting comprehensive material data: {}".format(str(e)))
    
    return material_usage_data    """Collect comprehensive family, type, and material data for ESG/CO2 analysis"""
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
                
                # Get element geometry for area/volume calculations
                element_volume = get_element_volume(element)
                element_area = get_element_area(element)
                
                # Get materials from the element
                material_ids = get_element_material_ids(element)
                
                if not material_ids:
                    # If no specific materials found, try to get from element type
                    element_type = doc.GetElement(element.GetTypeId())
                    if element_type:
                        type_material_id = element_type.LookupParameter("Material")
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
                        
                        material_info = {
                            # Element identification
                            'ElementId': element.Id.IntegerValue,
                            'ElementCategory': element.Category.Name if element.Category else "Unknown",
                            
                            # Family hierarchy
                            'FamilyName': family_name,
                            'FamilyType': family_type,
                            'TypeId': element.GetTypeId().IntegerValue,
                            
                            # Material information
                            'MaterialId': material_id.IntegerValue,
                            'MaterialName': material.Name if material.Name else "Unnamed Material",
                            'MaterialClass': material.MaterialClass if hasattr(material, 'MaterialClass') else "Unknown",
                            'MaterialCategory': material.MaterialCategory if hasattr(material, 'MaterialCategory') else "Unknown",
                            
                            # Thickness and quantities
                            'Thickness_mm': thickness,
                            'MaterialVolume_m3': material_volume,
                            'MaterialArea_m2': material_area,
                            'ElementTotalVolume_m3': element_volume,
                            'ElementTotalArea_m2': element_area,
                            
                            # Physical properties for CO2 calculations
                            'Density': get_material_property(material, 'Density'),
                            'ThermalConductivity': get_material_property(material, 'ThermalConductivity'),
                            'EstimatedMass_kg': calculate_material_mass(material_volume, material),
                            
                            # Custom parameters for ESG
                            'EmbodiedCarbon': get_custom_parameter(material, 'EmbodiedCarbon'),
                            'RecycledContent': get_custom_parameter(material, 'RecycledContent'),
                            'Manufacturer': get_custom_parameter(material, 'Manufacturer'),
                            
                            # Location in model
                            'Level': get_element_level(element),
                            'Location': get_element_location(element)
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
                        return round(thickness_mm, 2)
        
        # For other elements, try to get thickness parameter
        thickness_param = element.LookupParameter("Thickness")
        if thickness_param and thickness_param.HasValue:
            thickness_mm = UnitUtils.ConvertFromInternalUnits(thickness_param.AsDouble(), UnitTypeId.Millimeters)
            return round(thickness_mm, 2)
        
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
                return round(volume_cum, 4)
        
        # For other elements, use total volume
        volume_param = element.LookupParameter("Volume")
        if volume_param and volume_param.HasValue:
            volume_cuft = volume_param.AsDouble()
            volume_cum = UnitUtils.ConvertFromInternalUnits(volume_cuft, UnitTypeId.CubicMeters)
            return round(volume_cum, 4)
        
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
            return round(area_sqm, 2)
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
            return round(volume_cum, 4)
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
            return round(area_sqm, 2)
        return "N/A"
    except:
        return "N/A"

def get_material_property(material, property_name):
    """Get material physical property"""
    try:
        if hasattr(material, property_name):
            prop = getattr(material, property_name)
            return str(prop) if prop is not None else "N/A"
        return "N/A"
    except:
        return "N/A"

def calculate_material_mass(volume, material):
    """Calculate mass from volume and material density"""
    try:
        if volume != "N/A" and isinstance(volume, (int, float)):
            density_prop = getattr(material, 'Density', None)
            if density_prop:
                # This needs proper unit conversion - placeholder for now
                return "Mass calc needed"
        return "N/A"
    except:
        return "N/A"

def get_custom_parameter(material, param_name):
    """Get custom parameter value if it exists"""
    try:
        param = material.LookupParameter(param_name)
        if param and param.HasValue:
            if param.StorageType == StorageType.String:
                return param.AsString()
            elif param.StorageType == StorageType.Double:
                return str(param.AsDouble())
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
        return "N/A"
    except:
        return "N/A"

def get_element_level(element):
    """Get level of element"""
    try:
        level_param = element.LookupParameter("Level")
        if level_param and level_param.HasValue:
            level_id = level_param.AsElementId()
            level = doc.GetElement(level_id)
            return level.Name if level else "N/A"
        return "N/A"
    except:
        return "N/A"

def get_element_location(element):
    """Get element location"""
    try:
        location = element.Location
        if location and hasattr(location, 'Point'):
            point = location.Point
            return "X:{:.2f}, Y:{:.2f}, Z:{:.2f}".format(point.X, point.Y, point.Z)
        return "N/A"
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
                    headers = [
                        'ElementId', 'ElementCategory', 'FamilyName', 'FamilyType', 'TypeId',
                        'MaterialId', 'MaterialName', 'MaterialClass',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2', 
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2'
                    ]
                    writer.writerow(headers)
            
            return file_path, len(material_data)
        else:
            return None, 0
            
    except Exception as e:
        raise Exception("Error saving CSV file: {}".format(str(e)))    """Save comprehensive material data to CSV file with semicolon delimiter"""
    try:
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        save_dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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
                    # Updated headers - removed unwanted variables
                    headers = [
                        'ElementId', 'ElementCategory', 'FamilyName', 'FamilyType', 'TypeId',
                        'MaterialId', 'MaterialName', 'MaterialClass',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2', 
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2'
                    ]
                    writer.writerow(headers)
            
            return file_path, len(material_data)
        else:
            return None, 0
            
    except Exception as e:
        raise Exception("Error saving CSV file: {}".format(str(e)))    """Save comprehensive material data to CSV file with semicolon delimiter"""
    try:
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        save_dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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
                    headers = [
                        'ElementId', 'ElementCategory', 'FamilyName', 'FamilyType', 'TypeId',
                        'MaterialId', 'MaterialName', 'MaterialClass', 'MaterialCategory',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2', 
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2',
                        'Density', 'ThermalConductivity', 'EstimatedMass_kg',
                        'EmbodiedCarbon', 'RecycledContent', 'Manufacturer',
                        'Level', 'Location'
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
            
            # Use the new comprehensive function instead of get_material_data()
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
