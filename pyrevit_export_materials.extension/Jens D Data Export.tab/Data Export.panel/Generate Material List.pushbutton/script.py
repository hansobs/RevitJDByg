"""
Export Material Data Script
This script exports all families (system + loadable) with layers to CSV for material calculations.
"""

__title__ = "Export Material Data"
__author__ = "Jens Damm"

# Import required modules
from pyrevit import forms
from Autodesk.Revit.DB import *
import csv
import os
import clr

# Add reference to Windows Forms
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

def export_material_list():
    """Function to export material list to CSV"""
    try:
        doc = __revit__.ActiveUIDocument.Document
        
        # Output path
        out_path = os.path.expanduser("~/Documents/RevitFamilies.csv")
        
        rows = []
        rows.append(["ElementId", "Category", "Family", "Type", "Material", "Thickness (mm)", "Area (m2)", "Volume (m3)"])
        
        # ---- System families (Walls, Floors, Roofs, Ceilings) ----
        for cls in [Wall, Floor, RoofBase, Ceiling]:
            elements = FilteredElementCollector(doc).OfClass(cls).ToElements()
            for el in elements:
                etype = doc.GetElement(el.GetTypeId())
                famname = etype.FamilyName if hasattr(etype, "FamilyName") else "System Family"
                typename = etype.Name if etype else "No Type"
                
                # Area and volume
                area = el.LookupParameter("Area")
                vol = el.LookupParameter("Volume")
                area_m2 = round(UnitUtils.ConvertFromInternalUnits(area.AsDouble(), UnitTypeId.SquareMeters), 2) if area and area.HasValue else ""
                vol_m3 = round(UnitUtils.ConvertFromInternalUnits(vol.AsDouble(), UnitTypeId.CubicMeters), 2) if vol and vol.HasValue else ""
                
                # Layers (CompoundStructure)
                cs = etype.GetCompoundStructure() if hasattr(etype, "GetCompoundStructure") else None
                if cs:
                    for layer in cs.GetLayers():
                        thick = UnitUtils.ConvertFromInternalUnits(layer.Width, UnitTypeId.Millimeters)
                        mat = doc.GetElement(layer.MaterialId) if layer.MaterialId != ElementId.InvalidElementId else None
                        matname = mat.Name if mat else "No Material"
                        rows.append([el.Id.IntegerValue, el.Category.Name, famname, typename, matname, round(thick, 1), area_m2, vol_m3])
                else:
                    rows.append([el.Id.IntegerValue, el.Category.Name, famname, typename, "", "", area_m2, vol_m3])
        
        # ---- Loadable families (FamilyInstances) ----
        fam_instances = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
        for inst in fam_instances:
            etype = doc.GetElement(inst.GetTypeId())
            famname = etype.FamilyName if hasattr(etype, "FamilyName") else "FamilyInstance"
            typename = etype.Name if etype else "No Type"
            
            vol = inst.LookupParameter("Volume")
            area = inst.LookupParameter("Area")
            area_m2 = round(UnitUtils.ConvertFromInternalUnits(area.AsDouble(), UnitTypeId.SquareMeters), 2) if area and area.HasValue else ""
            vol_m3 = round(UnitUtils.ConvertFromInternalUnits(vol.AsDouble(), UnitTypeId.CubicMeters), 2) if vol and vol.HasValue else ""
            
            rows.append([inst.Id.IntegerValue, inst.Category.Name, famname, typename, "", "", area_m2, vol_m3])
        
        # ---- Save as CSV ----
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerows(rows)
        
        return out_path, len(rows) - 1  # Return path and number of exported items (excluding header)
        
    except Exception as e:
        raise Exception("Error during material export: {}".format(str(e)))

def main():
    """Main function that runs when the button is clicked"""
    try:
        # Show confirmation dialog before export
        result = MessageBox.Show(
            "This will generate a material list export in CSV format for calculation purposes.\n\nDo you want to proceed?",
            "Export Material Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        if result == System.Windows.Forms.DialogResult.Yes:
            # Perform the export
            file_path, item_count = export_material_list()
            
            # Show success message
            MessageBox.Show(
                "Export completed successfully!\n\nFile saved to:\n{}\n\nTotal items exported: {}".format(file_path, item_count),
                "Export Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            
            # Alternative using pyRevit's TaskDialog (more Revit-native)
            # TaskDialog.Show("Export Complete", "Export completed!\nFile saved: {}\nItems exported: {}".format(file_path, item_count))
        
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
