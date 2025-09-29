"""
Export Material Data Script for ESG/CO2 Analysis
This script exports comprehensive material data from Revit to CSV format.
"""
__title__ = "Export ESG Data"
__author__ = "Jens Damm & Hans Bohn Svendsen"

from pyrevit.output import get_output
import clr
import csv
import os
import time
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System')

from System.Windows.Forms import (
    MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult, 
    SaveFileDialog
)
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Global variables
doc = __revit__.ActiveUIDocument.Document
output = get_output()

# Unit type compatibility handling
def get_unit_type_millimeters():
    """Get the correct unit type for millimeters based on Revit version"""
    try:
        return UnitTypeId.Millimeters
    except:
        try:
            return DisplayUnitType.DUT_MILLIMETERS
        except:
            return None

def get_unit_type_square_meters():
    """Get the correct unit type for square meters based on Revit version"""
    try:
        return UnitTypeId.SquareMeters
    except:
        try:
            return DisplayUnitType.DUT_SQUARE_METERS
        except:
            return None

def get_unit_type_cubic_meters():
    """Get the correct unit type for cubic meters based on Revit version"""
    try:
        return UnitTypeId.CubicMeters
    except:
        try:
            return DisplayUnitType.DUT_CUBIC_METERS
        except:
            return None

def convert_from_internal_units(value, unit_type):
    """Convert from internal units with version compatibility"""
    if unit_type is None:
        return value
    try:
        return UnitUtils.ConvertFromInternalUnits(value, unit_type)
    except:
        return value

# Safe BuiltInParameter access
def get_safe_builtin_params(param_names):
    """Safely get BuiltInParameter values that exist in the current Revit version"""
    safe_params = []
    for param_name in param_names:
        try:
            param_value = getattr(BuiltInParameter, param_name)
            safe_params.append(param_value)
        except AttributeError:
            continue
    return safe_params

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

def get_category_material_info(element):
    """Get material information from element's category"""
    try:
        if element.Category:
            category = element.Category
            material = category.Material
            if material:
                return {
                    'material_id': material.Id.IntegerValue,
                    'material_name': material.Name,
                    'material_class': getattr(material, 'MaterialClass', 'Unknown')
                }
        return None
    except:
        return None

def get_parameter_value_comprehensive(element, builtin_param_names, unit_type=None, fallback_param_names=None):
    """Comprehensive parameter value retrieval with multiple fallback methods"""
    try:
        # Method 1: Try safe BuiltInParameters
        safe_builtin_params = get_safe_builtin_params(builtin_param_names)
        for builtin_param in safe_builtin_params:
            try:
                param = element.get_Parameter(builtin_param)
                if param and param.HasValue:
                    if unit_type:
                        value = param.AsDouble()
                        value = convert_from_internal_units(value, unit_type)
                        return format_number(value, Config.DEFAULT_DECIMALS)
                    else:
                        return param.AsString() or param.AsValueString()
            except:
                continue
        
        # Method 2: Try type parameters with BuiltInParameters
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            for builtin_param in safe_builtin_params:
                try:
                    param = element_type.get_Parameter(builtin_param)
                    if param and param.HasValue:
                        if unit_type:
                            value = param.AsDouble()
                            value = convert_from_internal_units(value, unit_type)
                            return format_number(value, Config.DEFAULT_DECIMALS)
                        else:
                            return param.AsString() or param.AsValueString()
                except:
                    continue
        
        # Method 3: Try string parameter lookup (instance)
        if fallback_param_names:
            for param_name in fallback_param_names:
                try:
                    param = element.LookupParameter(param_name)
                    if param and param.HasValue:
                        if unit_type:
                            value = param.AsDouble()
                            value = convert_from_internal_units(value, unit_type)
                            return format_number(value, Config.DEFAULT_DECIMALS)
                        else:
                            return param.AsString() or param.AsValueString()
                except:
                    continue
            
            # Method 4: Try string parameter lookup (type)
            if element_type:
                for param_name in fallback_param_names:
                    try:
                        param = element_type.LookupParameter(param_name)
                        if param and param.HasValue:
                            if unit_type:
                                value = param.AsDouble()
                                value = convert_from_internal_units(value, unit_type)
                                return format_number(value, Config.DEFAULT_DECIMALS)
                            else:
                                return param.AsString() or param.AsValueString()
                    except:
                        continue
        
        return "N/A"
    except:
        return "N/A"

def get_element_area_robust(element):
    """Get area with multiple fallback methods"""
    try:
        # Method 1: Try comprehensive parameter search
        area_result = get_parameter_value_comprehensive(
            element,
            ["HOST_AREA_COMPUTED", "ROOM_AREA"],
            get_unit_type_square_meters(),
            ["Area", "Gross Surface Area", "Net Surface Area"]
        )
        
        if area_result != "N/A":
            return area_result
        
        # Method 2: Direct parameter access for common elements
        try:
            if hasattr(element, 'WallType') or hasattr(element, 'FloorType') or hasattr(element, 'RoofType'):
                # Try to get area parameter directly
                area_param = element.LookupParameter("Area")
                if area_param and area_param.HasValue:
                    area_sqft = area_param.AsDouble()
                    area_sqm = convert_from_internal_units(area_sqft, get_unit_type_square_meters())
                    return format_number(area_sqm)
        except:
            pass
        
        # Method 3: Try all parameters to find area-related ones
        try:
            for param in element.Parameters:
                if param.Definition.Name.lower() in ['area', 'surface area', 'gross area', 'net area']:
                    if param.HasValue and param.StorageType == StorageType.Double:
                        value = param.AsDouble()
                        if value > 0:  # Only positive areas make sense
                            area_sqm = convert_from_internal_units(value, get_unit_type_square_meters())
                            return format_number(area_sqm)
        except:
            pass
        
        return "N/A"
    except:
        return "N/A"

def get_element_volume_robust(element):
    """Get volume with multiple fallback methods"""
    try:
        # Method 1: Try comprehensive parameter search
        volume_result = get_parameter_value_comprehensive(
            element,
            ["HOST_VOLUME_COMPUTED", "ROOM_VOLUME"],
            get_unit_type_cubic_meters(),
            ["Volume", "Gross Volume", "Net Volume"]
        )
        
        if volume_result != "N/A":
            return volume_result
        
        # Method 2: Direct parameter access
        try:
            volume_param = element.LookupParameter("Volume")
            if volume_param and volume_param.HasValue:
                volume_cuft = volume_param.AsDouble()
                volume_cum = convert_from_internal_units(volume_cuft, get_unit_type_cubic_meters())
                return format_number(volume_cum)
        except:
            pass
        
        # Method 3: Try all parameters to find volume-related ones
        try:
            for param in element.Parameters:
                if param.Definition.Name.lower() in ['volume', 'gross volume', 'net volume']:
                    if param.HasValue and param.StorageType == StorageType.Double:
                        value = param.AsDouble()
                        if value > 0:  # Only positive volumes make sense
                            volume_cum = convert_from_internal_units(value, get_unit_type_cubic_meters())
                            return format_number(volume_cum)
        except:
            pass
        
        return "N/A"
    except:
        return "N/A"

def get_element_width(element):
    """Get Width parameter from element using comprehensive search"""
    return get_parameter_value_comprehensive(
        element,
        ["DOOR_WIDTH", "WINDOW_WIDTH", "GENERIC_WIDTH", "FAMILY_WIDTH_PARAM"],
        get_unit_type_millimeters(),
        ["Width", "Rough Width", "Opening Width"]
    )

def get_element_height(element):
    """Get Height parameter from element using comprehensive search"""
    return get_parameter_value_comprehensive(
        element,
        ["DOOR_HEIGHT", "WINDOW_HEIGHT", "GENERIC_HEIGHT", "FAMILY_HEIGHT_PARAM", "WALL_USER_HEIGHT_PARAM"],
        get_unit_type_millimeters(),
        ["Height", "Rough Height", "Opening Height", "Unconnected Height"]
    )

def get_element_thickness(element):
    """Get Thickness parameter from element using comprehensive search"""
    return get_parameter_value_comprehensive(
        element,
        ["WALL_ATTR_WIDTH_PARAM", "GENERIC_THICKNESS", "FAMILY_THICKNESS_PARAM"],
        get_unit_type_millimeters(),
        ["Thickness", "Width"]
    )

@safe_execution()
def get_export_guid(element):
    """Get export GUID for element"""
    try:
        guid_str = ExportUtils.GetExportId(doc, element.Id)
        return str(guid_str) if guid_str else "N/A"
    except:
        return "N/A"

def get_material_layers(element):
    """Returns all layers with material + thickness, including enhanced 'By Category' handling"""
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
                        # Enhanced "By Category" handling
                        category_material = get_category_material_info(element)
                        if category_material:
                            mat_name = "By Category: {}".format(category_material['material_name'])
                            mat_id_for_export = "ByCategory_{}".format(category_material['material_id'])
                        else:
                            mat_name = "By Category: {}".format(element.Category.Name)
                            mat_id_for_export = "ByCategory_Unknown"
                    else:
                        mat = doc.GetElement(mat_id)
                        mat_name = mat.Name if mat else "Unknown Material"
                        mat_id_for_export = mat_id.IntegerValue
                    
                    thickness_mm = convert_from_internal_units(layer.Width, get_unit_type_millimeters())
                    results.append({
                        "LayerIndex": i,
                        "MaterialId": mat_id_for_export,
                        "MaterialName": mat_name,
                        "Thickness_mm": round(thickness_mm, 2)
                    })
            else:
                # Try to get thickness using comprehensive search
                thickness = get_element_thickness(element)
                if thickness != "N/A":
                    results.append({
                        "LayerIndex": 0,
                        "MaterialId": "N/A",
                        "MaterialName": "N/A",
                        "Thickness_mm": thickness
                    })
    except Exception as e:
        results.append({
            "LayerIndex": -1,
            "MaterialId": "Error",
            "MaterialName": "Error: {}".format(str(e)),
            "Thickness_mm": "N/A"
        })
    return results

def get_element_type_name(element):
    """Get the element type name using comprehensive search"""
    try:
        # Try comprehensive parameter search first
        type_name = get_parameter_value_comprehensive(
            element,
            ["ELEM_TYPE_PARAM", "SYMBOL_NAME_PARAM"],
            None,
            ["Type Name", "Family and Type", "Type"]
        )
        
        if type_name != "N/A":
            return type_name
        
        # Fallback to element type access
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
                if element_type.Category:
                    return "{} - ID {}".format(element_type.Category.Name, element_type.Id.IntegerValue)
                else:
                    return "Type ID {}".format(element_type.Id.IntegerValue)
            except:
                pass
        
        # Additional fallbacks
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
    """Get family name from element using comprehensive search"""
    try:
        # Try comprehensive parameter search first
        family_name = get_parameter_value_comprehensive(
            element,
            ["ELEM_FAMILY_PARAM", "SYMBOL_FAMILY_NAME_PARAM"],
            None,
            ["Family", "Family Name"]
        )
        
        if family_name != "N/A":
            return family_name
        
        # Fallback methods
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
    """Get family type name from element using comprehensive search"""
    try:
        # Try comprehensive parameter search first
        family_type = get_parameter_value_comprehensive(
            element,
            ["ELEM_FAMILY_AND_TYPE_PARAM", "SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM"],
            None,
            ["Family and Type", "Type"]
        )
        
        if family_type != "N/A":
            return family_type
        
        # Fallback methods
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
                        thickness_mm = convert_from_internal_units(thickness_feet, get_unit_type_millimeters())
                        return round(thickness_mm, 5)
        
        # Use comprehensive thickness search
        thickness = get_element_thickness(element)
        return thickness if thickness != "N/A" else "N/A"
    except:
        return "N/A"

def calculate_layer_volume(element, thickness_mm):
    """Calculate volume for a specific layer thickness"""
    try:
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType'):
            # Use robust area calculation
            area_str = get_element_area_robust(element)
            if area_str != "N/A" and thickness_mm != "N/A":
                try:
                    area_sqm = float(str(area_str).replace(',', '.'))
                    thickness_m = float(thickness_mm) / 1000.0  # Convert mm to m
                    volume_cum = area_sqm * thickness_m
                    return round(volume_cum, 5)
                except:
                    pass
        return "N/A"
    except:
        return "N/A"

def calculate_material_volume(element, material_id, thickness):
    """Calculate volume of specific material in element"""
    try:
        if hasattr(element, 'WallType') or hasattr(element, 'FloorType'):
            area_str = get_element_area_robust(element)
            if area_str != "N/A" and thickness != "N/A":
                try:
                    area_sqm = float(str(area_str).replace(',', '.'))
                    thickness_m = float(thickness) / 1000.0  # Convert mm to m
                    volume_cum = area_sqm * thickness_m
                    return round(volume_cum, 5)
                except:
                    pass
        
        # Fallback to element volume
        volume_str = get_element_volume_robust(element)
        if volume_str != "N/A":
            return volume_str
        
        return "N/A"
    except:
        return "N/A"

def calculate_material_area(element, material_id):
    """Calculate area of specific material in element"""
    try:
        area_str = get_element_area_robust(element)
        return area_str if area_str != "N/A" else "N/A"
    except:
        return "N/A"

class MaterialDataExtractor:
    """Class to handle material data extraction with progress tracking"""
    def __init__(self, document, output_window):
        self.doc = document
        self.output = output_window
        self.processed_elements = 0
        self.material_records = 0
        self.debug_info = {
            'total_elements': 0,
            'elements_with_category': 0,
            'elements_with_materials': 0,
            'elements_with_layers': 0,
            'elements_processed': 0,
            'errors': 0,
            'elements_with_area': 0,
            'elements_with_volume': 0,
            'by_category_materials': 0
        }

    def extract_all_materials(self):
        """Extract comprehensive material data from all elements"""
        material_usage_data = []
        try:
            elements = FilteredElementCollector(self.doc).WhereElementIsNotElementType().ToElements()
            self.debug_info['total_elements'] = len(elements)
            
            self.output.print_md("## Starting Material Data Collection")
            self.output.print_md("*Total elements to process:* {}".format(len(elements)))
            
            for i, element in enumerate(elements):
                if i % Config.PROGRESS_UPDATE_INTERVAL == 0 or i == len(elements) - 1:
                    self._update_progress(i + 1, len(elements))
                
                element_data = self._process_element(element)
                if element_data:
                    material_usage_data.extend(element_data)
                    self.material_records += len(element_data)
                    self.debug_info['elements_with_materials'] += 1
                
                self.processed_elements += 1
            
            self._print_completion_stats()
            self._print_debug_info()
            return material_usage_data
        except Exception as e:
            raise Exception("Error collecting comprehensive material data: {}".format(str(e)))

    def _process_element(self, element):
        """Process a single element and return its material data"""
        try:
            if not element.Category:
                return []
            
            self.debug_info['elements_with_category'] += 1
            
            element_info = self._get_element_info(element)
            
            # Track elements with area/volume for debugging
            if element_info['area'] != "N/A":
                self.debug_info['elements_with_area'] += 1
            if element_info['volume'] != "N/A":
                self.debug_info['elements_with_volume'] += 1
            
            material_layers = get_material_layers(element)
            
            if material_layers:
                self.debug_info['elements_with_layers'] += 1
                return self._process_element_layers(element, element_info, material_layers)
            else:
                # Try fallback method
                fallback_data = self._process_element_fallback(element, element_info)
                if fallback_data:
                    return fallback_data
                else:
                    # Create a basic record even if no materials found
                    return self._create_basic_element_record(element, element_info)
        except Exception as e:
            self.debug_info['errors'] += 1
            print("Error processing element {}: {}".format(element.Id.IntegerValue, str(e)))
            return []

    def _create_basic_element_record(self, element, element_info):
        """Create a basic record for elements without specific materials"""
        try:
            # Only create records for certain categories that should have materials
            relevant_categories = [
                "Walls", "Floors", "Roofs", "Ceilings", "Structural Framing", 
                "Structural Columns", "Doors", "Windows", "Furniture", "Casework"
            ]
            
            if element.Category and element.Category.Name in relevant_categories:
                basic_record = {
                    'ElementId': element_info['element_id'],
                    'ElementCategory': element_info['category'],
                    'ExportGUID': element_info['export_guid'],
                    'FamilyName': element_info['family_name'],
                    'FamilyType': element_info['family_type'],
                    'Type': element_info['element_type'],
                    'TypeId': element_info['type_id'],
                    'Width_mm': element_info['width'],
                    'Height_mm': element_info['height'],
                    'LayerIndex': 0,
                    'MaterialId': "No_Material",
                    'MaterialName': "No Material Assigned",
                    'MaterialClass': "Unknown",
                    'Thickness_mm': get_element_thickness(element),
                    'MaterialVolume_m3': "N/A",
                    'MaterialArea_m2': element_info['area'],
                    'ElementTotalVolume_m3': element_info['volume'],
                    'ElementTotalArea_m2': element_info['area']
                }
                return [basic_record]
            return []
        except:
            return []

    def _get_element_info(self, element):
        """Get common element information using robust parameter access"""
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
            'volume': get_element_volume_robust(element),
            'area': get_element_area_robust(element)
        }

    def _process_element_layers(self, element, element_info, material_layers):
        """Process element using material layers"""
        material_records = []
        for layer in material_layers:
            material_record = self._create_material_record(element, element_info, layer)
            material_records.append(material_record)
            
            # Track by category materials
            if str(layer.get('MaterialId', '')).startswith('ByCategory'):
                self.debug_info['by_category_materials'] += 1
                
        return material_records

    def _process_element_fallback(self, element, element_info):
        """Fallback processing for elements without compound structures"""
        material_records = []
        material_ids = get_element_material_ids(element)
        
        if material_ids:
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
        """Get material class from layer info with enhanced By Category handling"""
        material_id = layer_info.get('MaterialId')
        material_name = layer_info.get('MaterialName', '')
        
        if str(material_id).startswith('ByCategory'):
            if 'By Category:' in material_name:
                return "By Category (Resolved)"
            else:
                return "By Category (Unresolved)"
        elif material_id == "N/A" or material_id == "Error":
            return "Unknown"
        elif material_id == "No_Material":
            return "No Material"
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

    def _print_debug_info(self):
        """Print debug information"""
        self.output.print_md("## Debug Information")
        self.output.print_md("*Total elements:* {}".format(self.debug_info['total_elements']))
        self.output.print_md("*Elements with category:* {}".format(self.debug_info['elements_with_category']))
        self.output.print_md("*Elements with materials:* {}".format(self.debug_info['elements_with_materials']))
        self.output.print_md("*Elements with layers:* {}".format(self.debug_info['elements_with_layers']))
        self.output.print_md("*Elements with area:* {}".format(self.debug_info['elements_with_area']))
        self.output.print_md("*Elements with volume:* {}".format(self.debug_info['elements_with_volume']))
        self.output.print_md("*By Category materials:* {}".format(self.debug_info['by_category_materials']))
        self.output.print_md("*Processing errors:* {}".format(self.debug_info['errors']))

def save_to_csv(material_data):
    """Save comprehensive material data to CSV file with semicolon delimiter"""
    try:
        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        save_dialog.FilterIndex = 1
        save_dialog.RestoreDirectory = True
        
        try:
            initial_dir = os.path.join(os.path.expanduser("~"), "Documents")
            if os.path.exists(initial_dir):
                save_dialog.InitialDirectory = initial_dir
        except:
            pass
            
        save_dialog.FileName = "ESG_Material_Export_{}".format(Config.get_timestamp())
        
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
        result = MessageBox.Show(
            "This will generate a comprehensive material list export for ESG/CO2 analysis.\n\nDo you want to proceed?",
            "Export ESG Material Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        if result == DialogResult.Yes:
            start_time = time.time()
            
            output.close_others()
            output.print_md("# ESG Material Data Export (Enhanced By Category Handling)")
            output.print_md("*Started:* {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            
            extractor = MaterialDataExtractor(doc, output)
            material_data = extractor.extract_all_materials()
            
            if not material_data:
                MessageBox.Show(
                    "No materials found in the current model.\n\nPlease check the debug information in the output window.", 
                    "No Data",
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning
                )
                return
            
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
        print("ERROR: {}".format(error_msg))
        MessageBox.Show(
            error_msg,
            "Export Error", 
            MessageBoxButtons.OK, 
            MessageBoxIcon.Error
        )

if __name__ == '__main__':
    main()
