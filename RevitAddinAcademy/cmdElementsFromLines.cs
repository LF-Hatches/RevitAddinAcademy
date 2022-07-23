#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections;              //ArrayList, toChar and toInt
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;    //Added for Get File Name

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

            //Get Filepath from Dialog Box
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog(); //dialog box initialized
            dialog.InitialDirectory = @"C:\"; //initial directory
            dialog.Multiselect = false;       //single file
            dialog.Filter = "Excel Files | *.xlsx; *.xls"; //excel files
            //dialog.ShowHelp = true;
            //dialog.HelpRequest = "";

            string filePath = "";             //initialize filepath
            string[] filePaths;               //multiple filepaths
            if (dialog.Multiselect == true)    //switch between multiple and single files
            {
                if (dialog.ShowDialog() == Forms.DialogResult.OK)
                {
                    filePaths = dialog.FileNames;
                }
            }
            else
            {
                if (dialog.ShowDialog() == Forms.DialogResult.OK)
                {
                    filePath = dialog.FileName;
                }
            }

            //single filepath
            string excelFile = filePath; //pointed to new file; 
            //multiple filepaths - loop through with list of strings - not used here.

 


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Generate Lines"); //Start transaction



                t.Commit(); //Commit Transaction
            }



            TaskDialog.Show("Hello", "This is my first command add-in.");
            TaskDialog.Show("HEllo again", "This is another window");

            return Result.Succeeded;
        }

        //Class
        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach (View curView in collector)
            {
                if (curView.Name == viewName)
                {
                    return curView;
                }
            }
            return null;
        }

        //READ EXCEL CLASS
        internal List<SheetStruct> ReadExcelSheets(Excel.Workbook excelWb, int i)    //dataSheetList = ReadExcelSheets(excelWb, i);    //code this method
        {
            //initialize struct //set to new things in loop //leave these as blanks
            //SHEETS
            SheetStruct sheetData;           //5 in current data file, Sheets is WB name
            sheetData.SheetNumber = "A-101";     //Sheet Number
            sheetData.SheetName = "Sheet Name";  //Sheet Name
            sheetData.ViewLevel = "Level Name";  //View Level Name
            sheetData.DrawnBy = "AA";            //DrawnBy
            sheetData.CheckedBy = "BB";          //CheckedBy
            sheetData.SheetDisc = "Architectural"; //Discipline Header
            sheetData.SheetSort = 5.0;           //Sort Order
            sheetData.LevelElemID = 0;           //LevelElemID

            //initialize list
            List<SheetStruct> structList = new List<SheetStruct>();

            //loop through WB range i j x
            //Workbook Sheet - First sheet is 1 not 0

            Excel.Worksheet excelWs = excelWb.Worksheets.Item[i];
            Excel.Range excelRng = excelWs.UsedRange;
            int rowCount = excelRng.Rows.Count;
            int colCount = excelRng.Columns.Count;
            //List<int[]> indexList = new List<int[]>(); //Collection of header nums, use if out of order

            for (int j = 1; j <= rowCount; j++) //loop through rows
            {
                for (int k = 1; k <= colCount; k++) //loop through cols. last one is double, so change
                {
                    Excel.Range cell1 = excelWs.Cells[j, k]; //Cell at first row and 1st --- 5th cell of first column, 
                    if (k == 7)
                    {
                        double data1 = cell1.Value.ToDouble();
                        sheetData.SetDoubleAtIndex(data1, k); //Add double to struct at k
                    }
                    else
                    {
                        string data1 = cell1.Value.ToString();
                        sheetData.SetStringAtIndex(data1, k); //Add string to struct at k
                    }
                }
                sheetData.SetIntAtIndex(0, 8);           //LevelElemID initialized to zero
                //Add row (struct data set) to struct list
                structList.Add(sheetData);
            }
            Debug.Print("Got here sheets");  //Check-in

            return structList;
        }
        internal List<LevelStruct> ReadExcelLevels(Excel.Workbook excelWb, int i)    //dataLevelList = ReadExcelLevels(excelWb, i);    //code this method
        {
            //initialize struct //set to new things in loop //leave these as blanks
            //LEVELS
            LevelStruct levelData;           //3 in current data file, Levels is WB name
            levelData.Name = "Level Name";  //Level Names
            levelData.Elevation = 100.00;   //Level Elevations
            levelData.ElevationM = 10.5;    //Metric Elevations
            levelData.ElemID = 0;           //Element ID

            //initialize list
            List<LevelStruct> structList = new List<LevelStruct>();

            //loop through WB range i j x
            //Workbook Sheet - First sheet is 1 not 0
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[i];
            Excel.Range excelRng = excelWs.UsedRange;
            int rowCount = excelRng.Rows.Count;
            int colCount = excelRng.Columns.Count;

            for (int j = 1; j <= rowCount; j++) //loop through rows
            {
                for (int k = 1; k <= colCount; k++) //loop through cols. First row only is string.
                {
                    if (k == 1)
                    {
                        Excel.Range cell1 = excelWs.Cells[j, k]; //Cell at first row and 1st --- 2nd cell of first column 
                        string data1 = cell1.Value.ToString();
                        levelData.SetStringAtIndex(data1, k);    //Add string to struct at k     
                    }
                    else
                    {
                        Excel.Range cell1 = excelWs.Cells[j, k]; //Cell at first row and 1st --- 2nd cell of first column 
                        double data1 = cell1.Value.ToDouble();
                        levelData.SetDoubleAtIndex(data1, k);    //Add string to struct at k   
                    }
                }
                levelData.SetIntAtIndex(0, 4);           //Element ID initialized
                //Add row (struct data set) to struct list
                structList.Add(levelData);
            }
            Debug.Print("Got here Levels");  //Check-in
            return structList;
        }
        
        //DATA STRUCTURES
        internal struct LevelStruct  //3 in current data file, Levels is WB name
        {
            //Define variables accessed from outside
            public string Name;
            public double Elevation;
            public double ElevationM;
            public int    ElemID;           //Element ID

            //constructor method
            //method inside structure that is named the same; specify arguments inside it.
            public LevelStruct(string name, double elevation, double elevationM, int elemid)
            {
                Name = name;
                Elevation = elevation;
                ElevationM = elevationM;
                ElemID = elemid;

            }
            public void SetStringAtIndex(string passedstring, int index)    //struct method
            {
                if (index == 1) { Name = passedstring; }
                return;
            }
            public void SetDoubleAtIndex(double passedvalue, int index)    //struct method
            {
                if (index == 2) { Elevation = passedvalue; }
                else if (index == 3) { ElevationM = passedvalue; }
                return;
            }
            public void SetIntAtIndex(int passedval, int index)    //struct method
            {
                if (index == 4) { ElemID = passedval; }
                return;
            }

        }

        internal struct SheetStruct  //5 in current data file, Sheets is WB name
        {
            //Define variables accessed from outside
            public string SheetNumber;      //Sheet Number
            public string SheetName;        //Sheet Name
            public string ViewLevel;        //View Level Name
            public string DrawnBy;          //DrawnBy
            public string CheckedBy;        //CheckedBy
            public string SheetDisc;        //Discipline Header
            public double SheetSort;        //Sort Order
            public int LevelElemID;         //LevelElemID

            public SheetStruct(            //constructor
                string sheetnumber,
                string sheetname,
                string viewlevel,
                string drawnby,
                string checkedby,
                string sheetdisc,
                double sheetsort,
                int levelelemid)
            {
                SheetNumber = sheetnumber;    //Sheet Number
                SheetName = sheetname;        //Sheet Name
                ViewLevel = viewlevel;        //View Level Name
                DrawnBy = drawnby;            //DrawnBy
                CheckedBy = checkedby;        //CheckedBy
                SheetDisc = sheetdisc;        //Discipline Header
                SheetSort = sheetsort;        //Sort Order
                LevelElemID = levelelemid;    //LevelElemID
            }
            public void SetStringAtIndex(string passedstring, int index)    //struct method
            {
                if (index == 1) { SheetNumber = passedstring; }
                else if (index == 2) { SheetName = passedstring; }
                else if (index == 3) { ViewLevel = passedstring; }
                else if (index == 4) { DrawnBy = passedstring; }
                else if (index == 5) { CheckedBy = passedstring; }
                else if (index == 6) { SheetDisc = passedstring; }
                return;
            }
            public void SetDoubleAtIndex(double passedvalue, int index)    //struct method
            {
                if (index == 7) { SheetSort = passedvalue; }
                return;
            }
            public void SetIntAtIndex(int passedvalue, int index)    //struct method
            {
                if (index == 8) { LevelElemID = passedvalue; }
                return;
            }
        }
    }
}
