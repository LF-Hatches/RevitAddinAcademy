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
    public class Command01Challenge : IExternalCommand
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

            string textFizz = "FIZZ";
            string textBuzz = "BUZZ";
            string textBoth = "FIZZBUZZ";
            string curString = "";


            string fileName = doc.PathName;

            double offset = 0.05; //base unit of current model  so that is feet
            double offsetCalc = offset * doc.ActiveView.Scale;


            XYZ curPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note"); //what appears in the undo
            t.Start();

            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                if( ((i%3)==0) && ((i%5)==0))
                {
                    curString = textBoth;
                }
                else if ((i % 3) == 0)
                {
                    curString = textFizz;
                }
                else if ((i % 5) == 0)
                {
                    curString = textBuzz;
                }
                else
                {
                    curString = i.ToString();
                }
                //TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "            This is Line " + i.ToString(), collector.FirstElementId());
                TextNote curNote2 = TextNote.Create(doc, doc.ActiveView.Id, curPoint, curString, collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);
            }

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

    }
}
