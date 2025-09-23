"""
Export Material Data Script
This script exports material data from Revit to CSV format.
"""

__title__ = "Export Data"
__author__ = "Your Name"

# Import required modules
from pyrevit import forms
import clr

# Add references to .NET assemblies
clr.AddReference('System.Windows.Forms')
clr.AddReference('System')

# Import .NET classes
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

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
            # User clicked Yes - show success message for now
            MessageBox.Show(
                "Export would happen here!\n\n(This is just a demo - actual export functionality would be implemented here)",
                "Export Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
        else:
            # User clicked No - show cancellation message
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
