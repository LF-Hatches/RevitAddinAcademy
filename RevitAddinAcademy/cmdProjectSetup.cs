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
    public class cmdProjectSetup : IExternalCommand  //rename for each new cmd file
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

            string excelFile = @"Session02_Challenge-220706-113155.xlsx"; //pointed to new file
            int NumWBSheets = 0;

            Excel.Application excelApp = new Excel.Application();        //Open Application Excel
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); //Workbook File
            NumWBSheets = excelWb.Sheets.Count;             //Number of sheets to loop through
            List<string[]> dataList = new List<string[]>(); //Collection of string arrays outside loop

            //Get all data first. Then do revit actions.
            //Levels live on sheet 1 - index 1; Sheets live in 2.
            for (int i=0; i<NumWBSheets; i++) //Loop through all WB sheets
            {
                //Workbook Sheet - First sheet is 1 not 0
                Excel.Worksheet excelWs = excelWb.Worksheets.Item[i];   //linked to loop     
                Excel.Range excelRng = excelWs.UsedRange;
                int rowCount = excelRng.Rows.Count;

                //Make method of ReadExcel later
                //Read Excel - Transforming Rows and Columns to List of Arrays
                for (int j = 1; j <= rowCount; j++)
                {
                    Excel.Range cell1 = excelWs.Cells[j, 1]; //First cell of first column
                    Excel.Range cell2 = excelWs.Cells[j, 2]; //First cell of second column

                    string data1 = cell1.Value.ToString();
                    string data2 = cell2.Value.ToString();

                    string[] dataArray = new string[2];      //two elements in array
                    dataArray[0] = data1;
                    dataArray[1] = data2;
                    dataList.Add(dataArray);
                    Debug.Print("Data 1: " + data1.ToString());  //Check-in
                    Debug.Print("Data 2: " + data2.ToString());  //Check-in
                }

            }


            //Do Revit Actions
            //Check for which data has what type
            //If sheet, add sheet
            //If level, add level


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create some Revit stuff"); //Start transaction
                Level curLevel = Level.Create(doc, 100);     //create level - default imperial feet
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
