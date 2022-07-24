#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;


#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand //rename for each new cmd module
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;



            using (Transaction t = new Transaction(doc)) 
            {
                t.Start("Create Revit stuff");

                t.Commit();
            }
            TaskDialog.Show("Complete", "We got here");
                        
            return Result.Succeeded;
        }

    }
}

