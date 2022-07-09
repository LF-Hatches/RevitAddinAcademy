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

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            string excelFile = @"C:\Users\LFarrell\Desktop\Revit Add-in Academy\Class Week 2\Session02_CombinationSheetList-220706-171323.xlsx";

            Excel.Application excelApp = new Excel.Application();        //Open Application Excel
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); //Workbook File
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[1];        //Worksheet Sheet - First sheet is 1 not 0

            Excel.Range excelRng = excelWs.UsedRange;
            int rowCount = excelRng.Rows.Count;
            // Read Excel - Transforming Rows and Columns to List of Arrays
            List<string[]> dataList = new List<string[]>(); //Collection of string arrays
            for (int i=1; i<=rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1]; //First cell of first column
                Excel.Range cell2 = excelWs.Cells[i, 2]; //First cell of second column

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];      //two elements in array
                dataArray[0] = data1;
                dataArray[1] = data2;
                dataList.Add(dataArray);
            }
            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create some Revit stuff"); //Start transaction
                Level curLevel = Level.Create(doc, 100);     //imperial feet
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks); //get Titleblock Type Category
                collector.WhereElementIsElementType();  //Types of titleblock types
                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId()); //uses first type of titleblock kind
 /*  Build in Error Checking Procedures for sheets, for levels
                //check for if placeholder - overwrite or leave
                //check for if sheet number exists already - get list of current sheets in doc
                //loop through current sheet numbers to check
                //if placeholder, delete placeholder and make new sheet
                //if not placeholder - skip entry?
 */
 
                curSheet.SheetNumber = "A101010"; //Directly exposed elements. Checker checkedby etc not avail.
                curSheet.Name = "New Sheet";

                t.Commit(); //Commit Transaction
            }


            excelWb.Close();
            excelApp.Quit();

            TaskDialog.Show("Hello", "This is my first command add-in.");
            TaskDialog.Show("HEllo again", "This is another window");
            
            return Result.Succeeded;
        }
    }
}
