"""
Hello World Script
This is my first pyRevit script that displays a simple dialog.
"""

__title__ = "Hello World"
__author__ = "Your Name"

# Import required modules
from pyrevit import forms
import clr

# Add reference to Windows Forms
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

def main():
    """Main function that runs when the button is clicked"""
    try:
        # Show a simple message box
        MessageBox.Show(
            "Hello World! This is my first pyRevit script!",
            "Hello World Dialog",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        
        # Alternative using pyRevit's forms (more advanced)
        # forms.alert("Hello World! This is my first pyRevit script!", title="Hello World")
        
    except Exception as e:
        # Error handling
        MessageBox.Show(
            "An error occurred: {}".format(str(e)),
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )

# This is the entry point when the button is clicked
if __name__ == '__main__':
    main()