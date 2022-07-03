#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command01 : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            int range = 100;

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            TaskDialog.Show("Hello", "This is Command 01");

            int number = 1;
            double number2 = 10.5;
            string text = "Reit Add-in Academy";
            XYZ point = new XYZ(0, 0, 0);
            XYZ point2 = new XYZ(0, 0, 0);


            double math = number * number2 + 100;
            double math2 = math % number2;

            List<string> strings = new List<string>();
            strings.Add("item 1");
            strings.Add("item 2");

            List<XYZ> points = new List<XYZ>();
            points.Add(point);
            points.Add(point2);



            for (int i = 1; i<range; i++)
            {
                number = number + 1;
            }

            string newString = "";
            foreach(string s in strings)
            {
                if(s == "item 1")
                {
                    newString = "got to 1";
                }
                else if(s == "item 2")
                {
                    newString = "got to 2";
                }
                else 
                {
                    newString = "Got somewhere else";
                }
                newString = newString + s;
            }

            double newNumber = Method01(100, 100);

            return Result.Succeeded;
        }

        internal double Method01(double a, double b)
        {
            double c = a + b;
            Debug.Print("Got here: " + c.ToString());
            return c;
        }
    }
}
