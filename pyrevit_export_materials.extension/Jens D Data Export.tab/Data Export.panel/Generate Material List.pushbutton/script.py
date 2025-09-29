"""
Export Material Data Script for ESG/CO2 Analysis
This script exports comprehensive material data from Revit to CSV format.
"""
__title__ = "Export ESG Data"
__author__ = "Jens Damm & Hans Bohn Svendsen"

# Remove the conflicting pyrevit.forms import
from pyrevit.output import get_output
import clr
import csv
import os
import time
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System')

# Import only what you need from Windows Forms
from System.Windows.Forms import (
    MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult, 
    SaveFileDialog, FolderBrowserDialog
)
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Global variables
doc = __revit__.ActiveUIDocument.Document
output = get_output()


# Configuration constants
class Config:
    DEFAULT_DECIMALS = 5
    CSV_DELIMITER = ';'
    PROGRESS_UPDATE_INTERVAL = 50
    CSV_WRITE_INTERVAL = 500
    
    @staticmethod
    def get_timestamp():
        return datetime.now().strftime("%Y%m%d %H%M")

def safe_execution(default_return="N/A"):
    """Decorator to handle exceptions and return default value"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except:
                return default_return
        return wrapper
    return decorator

def format_number(value, decimals=Config.DEFAULT_DECIMALS):
    """Format a number to a specific number of decimal places, removing trailing zeros"""
    if value == "N/A" or value is None:
        return "N/A"
    try:
        formatted = "{:.{}f}".format(float(value), decimals)
        formatted = formatted.rstrip('0').rstrip('.')
        if formatted == "":
            formatted = "0"
        formatted = formatted.replace('.', ',')
        return formatted
    except:
        return "N/A"

def get_element_parameter(element, param_names, unit_type=None, fallback_builtin_params=None):
    """Generic function to get parameter value from element or its type"""
    try:
        # Try instance parameters first
        for param_name in param_names:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                value = param.AsDouble() if unit_type else param.AsString()
                if unit_type:
                    value = UnitUtils.ConvertFromInternalUnits(value, unit_type)
                return format_number(value, Config.DEFAULT_DECIMALS) if unit_type else value
        
        # Try type parameters
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            for param_name in param_names:
                param = element_type.LookupParameter(param_name)
                if param and param.HasValue:
                    value = param.AsDouble() if unit_type else param.AsString()
                    if unit_type:
                        value = UnitUtils.ConvertFromInternalUnits(value, unit_type)
                    return format_number(value, Config.DEFAULT_DECIMALS) if unit_type else value
        
        # Try built-in parameters as fallback
        if fallback_builtin_params:
            for builtin_param in fallback_builtin_params:
                try:
                    param = element.get_Parameter(builtin_param)
                    if param and param.HasValue:
                        value = param.AsDouble() if unit_type else param.AsString()
                        if unit_type:
                            value = UnitUtils.ConvertFromInternalUnits(value, unit_type)
                        return format_number(value, Config.DEFAULT_DECIMALS) if unit_type else value
                except:
                    continue
        
        return "N/A"
    except:
        return "N/A"

def get_element_width(element):
    """Get Width parameter from element"""
    return get_element_parameter(
        element,
        ["Width"],
        UnitTypeId.Millimeters,
        [BuiltInParameter.DOOR_WIDTH, BuiltInParameter.WINDOW_WIDTH]
    )

def get_element_height(element):
    """Get Height parameter from element"""
    return get_element_parameter(
        element,
        ["Height"],
        UnitTypeId.Millimeters,
        [BuiltInParameter.DOOR_HEIGHT, BuiltInParameter.WINDOW_HEIGHT]
    )

@safe_execution()
def get_export_guid(element):
    """Get export GUID for element"""
    guid_str = ExportUtils.GetExportId(doc, element.Id)
    return str(guid_str) if guid_str else "N/A"

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
        type_id = element.GetTypeId()
        if type_id == ElementId.InvalidElementId:
            return "No TypeId"
        
        element_type = doc.GetElement(type_id)
        if element_type:
            if hasattr(element_type, 'Name'):
                name = element_type.Name
                if name and name.strip():
                    return name
            
            if hasattr(element_type, 'get_Name'):
                try:
                    name = element_type.get_Name()
                    if name and name.strip():
                        return name
                except:
                    pass
            
            try:
                type_name_param = element_type.LookupParameter("Type Name")
                if type_name_param and type_name_param.HasValue:
                    return type_name_param.AsString()
            except:
                pass
            
            try:
                if element_type.Category:
                    return "{} - ID {}".format(element_type.Category.Name, element_type.Id.IntegerValue)
                else:
                    return "Type ID {}".format(element_type.Id.IntegerValue)
            except:
                pass
        
        if hasattr(element, 'Symbol') and element.Symbol:
            return element.Symbol.Name
        elif hasattr(element, 'WallType') and element.WallType:
            return element.WallType.Name
        elif hasattr(element, 'FloorType') and element.FloorType:
            return element.FloorType.Name
        elif hasattr(element, 'RoofType') and element.RoofType:
            return element.RoofType.Name
        
        return "No type found"
    except Exception as e:
        return "Exception: {}".format(str(e))

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
        family_type_param = element.LookupParameter("Family and Type")
        if family_type_param and family_type_param.HasValue:
            return family_type_param.AsValueString()
        
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
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType') or hasattr(element, 'RoofType'):
            element_type = doc.GetElement(element.GetTypeId())
            if hasattr(element_type, 'GetCompoundStructure'):
                compound_structure = element_type.GetCompoundStructure()
                if compound_structure:
                    layers = compound_structure.GetLayers()
                    for layer in layers:
                        if layer.MaterialId != ElementId.InvalidElementId:
                            material_ids.append(layer.MaterialId)
        
        try:
            element_material_ids = element.GetMaterialIds(False)
            material_ids.extend(element_material_ids)
        except:
            pass
        
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
        element_type = doc.GetElement(element.GetTypeId())
        if hasattr(element_type, 'GetCompoundStructure'):
            compound_structure = element_type.GetCompoundStructure()
            if compound_structure:
                layers = compound_structure.GetLayers()
                for layer in layers:
                    if layer.MaterialId == material_id:
                        thickness_feet = layer.Width
                        thickness_mm = UnitUtils.ConvertFromInternalUnits(thickness_feet, UnitTypeId.Millimeters)
                        return round(thickness_mm, 5)
        
        thickness_param = element.LookupParameter("Thickness")
        if thickness_param and thickness_param.HasValue:
            thickness_mm = UnitUtils.ConvertFromInternalUnits(thickness_param.AsDouble(), UnitTypeId.Millimeters)
            return round(thickness_mm, 5)
        return "N/A"
    except:
        return "N/A"

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

class MaterialDataExtractor:
    """Class to handle material data extraction with progress tracking"""
    
    def __init__(self, document, output_window):
        self.doc = document
        self.output = output_window
        self.processed_elements = 0
        self.material_records = 0
    
    def extract_all_materials(self):
        """Extract comprehensive material data from all elements"""
        material_usage_data = []
        
        try:
            elements = FilteredElementCollector(self.doc).WhereElementIsNotElementType().ToElements()
            total_elements = len(elements)
            
            self.output.print_md("## Starting Material Data Collection")
            self.output.print_md("*Total elements to process:* {}".format(total_elements))
            
            for i, element in enumerate(elements):
                if i % Config.PROGRESS_UPDATE_INTERVAL == 0 or i == total_elements - 1:
                    self._update_progress(i + 1, total_elements)
                
                element_data = self._process_element(element)
                if element_data:
                    material_usage_data.extend(element_data)
                    self.material_records += len(element_data)
                
                self.processed_elements += 1
            
            self._print_completion_stats()
            return material_usage_data
            
        except Exception as e:
            raise Exception("Error collecting comprehensive material data: {}".format(str(e)))
    
    def _process_element(self, element):
        """Process a single element and return its material data"""
        try:
            if not element.Category:
                return []
            
            element_info = self._get_element_info(element)
            material_layers = get_material_layers(element)
            
            if not material_layers:
                return self._process_element_fallback(element, element_info)
            else:
                return self._process_element_layers(element, element_info, material_layers)
                
        except Exception:
            return []
    
    def _get_element_info(self, element):
        """Get common element information"""
        return {
            'element_id': element.Id.IntegerValue,
            'category': element.Category.Name if element.Category else "Unknown",
            'export_guid': get_export_guid(element),
            'family_name': get_family_name(element),
            'family_type': get_family_type(element),
            'element_type': get_element_type_name(element),
            'type_id': element.GetTypeId().IntegerValue,
            'width': get_element_width(element),
            'height': get_element_height(element),
            'volume': get_element_volume(element),
            'area': get_element_area(element)
        }
    
    def _process_element_layers(self, element, element_info, material_layers):
        """Process element using material layers"""
        material_records = []
        
        for layer in material_layers:
            material_record = self._create_material_record(element, element_info, layer)
            material_records.append(material_record)
        
        return material_records
    
    def _process_element_fallback(self, element, element_info):
        """Fallback processing for elements without compound structures"""
        material_records = []
        material_ids = get_element_material_ids(element)
        
        for material_id in material_ids:
            material = self.doc.GetElement(material_id)
            if material:
                layer_info = {
                    'LayerIndex': 0,
                    'MaterialId': material_id.IntegerValue,
                    'MaterialName': material.Name if material.Name else "Unnamed Material",
                    'Thickness_mm': get_material_thickness(element, material_id)
                }
                material_record = self._create_material_record(element, element_info, layer_info)
                material_records.append(material_record)
        
        return material_records
    
    def _create_material_record(self, element, element_info, layer_info):
        """Create a standardized material record"""
        thickness = layer_info.get('Thickness_mm', "N/A")
        material_volume = calculate_layer_volume(element, thickness)
        material_area = calculate_material_area(element, None)
        
        return {
            'ElementId': element_info['element_id'],
            'ElementCategory': element_info['category'],
            'ExportGUID': element_info['export_guid'],
            'FamilyName': element_info['family_name'],
            'FamilyType': element_info['family_type'],
            'Type': element_info['element_type'],
            'TypeId': element_info['type_id'],
            'Width_mm': element_info['width'],
            'Height_mm': element_info['height'],
            'LayerIndex': layer_info.get('LayerIndex', 0),
            'MaterialId': layer_info.get('MaterialId', "N/A"),
            'MaterialName': layer_info.get('MaterialName', "N/A"),
            'MaterialClass': self._get_material_class(layer_info),
            'Thickness_mm': format_number(thickness),
            'MaterialVolume_m3': format_number(material_volume),
            'MaterialArea_m2': format_number(material_area),
            'ElementTotalVolume_m3': format_number(element_info['volume']),
            'ElementTotalArea_m2': format_number(element_info['area'])
        }
    
    def _get_material_class(self, layer_info):
        """Get material class from layer info"""
        material_id = layer_info.get('MaterialId')
        if material_id == "By_Category":
            return "By Category"
        elif material_id == "N/A" or material_id == "Error":
            return "Unknown"
        else:
            try:
                material = self.doc.GetElement(ElementId(int(material_id)))
                return material.MaterialClass if material and hasattr(material, 'MaterialClass') else "Unknown"
            except:
                return "Unknown"
    
    def _update_progress(self, current, total):
        """Update progress display"""
        progress_percent = int(current * 100 / total)
        self.output.update_progress(current, total)
        print("Processing element {} of {} ({}%) - {} material records so far".format(
            current, total, progress_percent, self.material_records))
    
    def _print_completion_stats(self):
        """Print completion statistics"""
        self.output.print_md("## Collection Complete!")
        self.output.print_md("*Elements processed:* {}".format(self.processed_elements))
        self.output.print_md("*Material records created:* {}".format(self.material_records))

def save_to_csv(material_data):
    """Save comprehensive material data to CSV file with semicolon delimiter"""
    try:
        # Create and configure the SaveFileDialog properly
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        
        # Use a more reliable path method
        try:
            initial_dir = os.path.join(os.path.expanduser("~"), "Documents")
            if os.path.exists(initial_dir):
                save_dialog.InitialDirectory = initial_dir
        except:
            pass  # Fall back to default directory
            
        save_dialog.FileName = "ESG_Material_Export_{}".format(Config.get_timestamp())
        
        # Show dialog and handle result
        dialog_result = save_dialog.ShowDialog()
        
        if dialog_result == DialogResult.OK:
            file_path = save_dialog.FileName
            output.print_md("## Writing CSV File...")
            print("Writing {} material records to CSV...".format(len(material_data)))
            
            # Use binary mode for Python 2.7 compatibility
            with open(file_path, 'wb') as csvfile:
                if material_data:
                    fieldnames = material_data[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames,
                                          delimiter=Config.CSV_DELIMITER, 
                                          lineterminator='\n')
                    writer.writeheader()
                    
                    for i, material in enumerate(material_data):
                        if i % Config.CSV_WRITE_INTERVAL == 0:
                            progress_percent = int((i + 1) * 100 / len(material_data))
                            print("Writing record {} of {} ({}%)".format(i + 1, len(material_data), progress_percent))
                        writer.writerow(material)
                else:
                    writer = csv.writer(csvfile, delimiter=Config.CSV_DELIMITER, 
                                      lineterminator='\n')
                    headers = [
                        'ElementId', 'ElementCategory', 'ExportGUID', 'FamilyName', 'FamilyType', 'Type', 'TypeId',
                        'Width_mm', 'Height_mm', 'LayerIndex',
                        'MaterialId', 'MaterialName', 'MaterialClass',
                        'Thickness_mm', 'MaterialVolume_m3', 'MaterialArea_m2',
                        'ElementTotalVolume_m3', 'ElementTotalArea_m2'
                    ]
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
        # Show confirmation dialog
        result = MessageBox.Show(
            "This will generate a comprehensive material list export for ESG/CO2 analysis.\n\nDo you want to proceed?",
            "Export ESG Material Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        if result == DialogResult.Yes:
            start_time = time.time()
            
            # Clear output window
            output.close_others()
            output.print_md("# ESG Material Data Export")
            output.print_md("*Started:* {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            
            # Extract material data
            extractor = MaterialDataExtractor(doc, output)
            material_data = extractor.extract_all_materials()
            
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
            
            end_time = time.time()
            elapsed_time = round(end_time - start_time, 2)
            
            if file_path:
                output.print_md("## Export Results")
                output.print_md("*File:* {}".format(file_path))
                output.print_md("*Records:* {}".format(count))
                output.print_md("*Time:* {} seconds".format(elapsed_time))
                
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
        error_msg = "An error occurred during export:\n\n{}".format(str(e))
        print("ERROR: {}".format(error_msg))  # Also print to console for debugging
        MessageBox.Show(
            error_msg,
            "Export Error", 
            MessageBoxButtons.OK, 
            MessageBoxIcon.Error
        )

if __name__ == '__main__':
    main()