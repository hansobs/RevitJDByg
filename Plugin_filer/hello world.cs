using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;


namespace FirstPlugin
{
    public class Class1 : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Myribbonpanel");
            
            // Lav en knap som trigger en eller anden command og tilføj den til ribbon panel
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("cmdMyTest", "My Test", thisAssemblyPath, "FirstPlugin.MyTest");
            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
            
            // ellers kan du lave andre egenskaber til kanppen som fx tooltip
            pushButton.ToolTip = "Hej hans se hvad jeg har lavet";

            // bitmap icon
            Uri urlImage = new Uri(@"C:\Users\jensd\Documents\JD-Tegnogbyg\Plugin\Test\Icons\Hans.png");
            BitmapImage bitmapImage = new BitmapImage(urlImage);
            pushButton.LargeImage = bitmapImage;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class MyTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var app = uiapp.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            TaskDialog.Show("Revit", "Hej med dig");

            return Result.Succeeded;

        }
    }
}
