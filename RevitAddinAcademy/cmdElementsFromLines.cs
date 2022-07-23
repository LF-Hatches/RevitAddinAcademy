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

            Excel.Application excelApp = new Excel.Application();        //Open Application Excel
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); //Workbook File
            int NumWBSheets = excelWb.Sheets.Count;                      //Number of sheets to loop through

            //THIS IS WHERE WE SWITCH TO LIST OF STRUCTS INSTEAD OF STRINGS
            List<SheetStruct> dataSheetList = new List<SheetStruct>();
            List<LevelStruct> dataLevelList = new List<LevelStruct>();

            //Get all data first. Then do revit actions.
            //Read Excel - Transforming Rows and Columns to List of Structs
            for (int i = 1; i <= NumWBSheets; i++) //Loop through all WB sheets
            {
                string excelWSName = excelWb.Sheets[i].Name;
                //if named Sheets - run sheet method
                if (excelWSName == "Sheets")
                {
                    dataSheetList = ReadExcelSheets(excelWb, i);
                }
                //if named Levels - run levels method
                else if (excelWSName == "Levels")
                {
                    dataLevelList = ReadExcelLevels(excelWb, i);
                }
            }
            //Information has been stored in dataSheetList and dataLevelList;
            //Lists can't contain Structs of different types...

            //Do Revit Actions

            /*
             
            //View Creation

          
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));
            ViewFamilyType curVFT = null;
            ViewFamilyType curRCPVFT = null;
            //base type is Element - most generic use
            foreach(ViewFamilyType curElem in collector)
            {
                //if don't know name
                if(curElem.ViewFamily == ViewFamily.FloorPlan)
                {
                    curVFT = curElem;
                }
                //if do know name
                if (curElem.Name == "Floor Plan")
                {
                    curVFT = curElem;
                }
                else if(curElem.ViewFamily == ViewFamily.CeilingPlan)
                {
                    curRCPVFT = curElem;
                }
            }


            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector2.WhereElementIsElementType();

             
             */


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Project Setup"); //Start transaction

                //Setup Viewport Type (session 03)
                FilteredElementCollector collector3 = new FilteredElementCollector(doc);
                collector3.OfClass(typeof(ViewFamilyType));
                ViewFamilyType curVFT = null;
                ViewFamilyType curRCPVFT = null;
                //base type is Element - most generic use
                foreach (ViewFamilyType curElem in collector3)
                {
                    //if don't know name
                    if (curElem.ViewFamily == ViewFamily.FloorPlan)
                    {
                        curVFT = curElem;
                    }
                    //if do know name
                    if (curElem.Name == "Floor Plan")
                    {
                        curVFT = curElem;
                    }
                    else if (curElem.ViewFamily == ViewFamily.CeilingPlan)
                    {
                        curRCPVFT = curElem;
                    }
                }

                //Setup Titleblock type (session 02) first/any titleblock (Session 03 repeat)
                FilteredElementCollector collector2 = new FilteredElementCollector(doc);
                collector2.OfCategory(BuiltInCategory.OST_TitleBlocks); //get Titleblock Type Category
                collector2.WhereElementIsElementType();  //Types of titleblock types

                //FOR LEVELS
                //processing: dataLevelList
                int count1 = dataLevelList.Count;
                for (int i = 1; i < count1; ++i) //i=1 skips header row
                {
                    Level curLevel = Level.Create(doc, dataLevelList[i].Elevation); //create level - default imperial feet
                    curLevel.Name = dataLevelList[i].Name;
                    int idInteger = curLevel.Id.IntegerValue;
                    dataLevelList[i].SetIntAtIndex(idInteger, 4); //4th item in struct
                }

                //FOR SHEETS
                //processing: dataSheetList
                /*
                //Got level ID in ElemID - item 4 in struct
                //Put ElemID in Sheet List if SHEET:View Level matches LEVELS:Level Name
                */
                int count2 = dataSheetList.Count;
                for (int i = 1; i < count2; ++i) //process ELNUMS
                {
                    for (int j = 1; j < count1; ++j) //loops through levels
                    {
                        if (dataSheetList[i].ViewLevel == dataLevelList[j].Name)
                        {
                            int ElementId = dataLevelList[j].ElemID;
                            dataSheetList[i].SetIntAtIndex(ElementId, 8); //8th item in struct
                        }
                    }
                }
                for (int i = 1; i < count2; ++i) //i=1 skips header row
                {
                    //SHEETS
                    ViewSheet curSheet = ViewSheet.Create(doc, collector2.FirstElementId()); //uses first type of titleblock kind
                    curSheet.SheetNumber = dataSheetList[i].SheetNumber;
                    curSheet.Name = dataSheetList[i].SheetName;


                    //FOR VIEWS
                    //Loop through sheet view list

                    ElementId id = new ElementId(dataSheetList[i].LevelElemID);
                    if (dataSheetList[i].SheetName.Contains("RCP"))
                    {
                        ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, id);
                        curRCP.Name += "RCP"; //delete RCPs out of view level column in excel
                        //Class call
                        View existingView = GetViewByName(doc, dataSheetList[i].ViewLevel); //enter "level 3" abstract
                        if (existingView != null)
                        {
                            Viewport newVP = Viewport.Create(doc, existingView.Id, curRCP.Id, new XYZ(0, 0, 0));
                        }
                        else
                        {
                            TaskDialog.Show("got here", "got here11"); //add more from 
                        }
                    }
                    else
                    {
                        ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, id);
                        //Class call

                        View existingView2 = GetViewByName(doc, dataSheetList[i].ViewLevel); //enter "level 3" abstract
                        if (existingView2 != null)
                        {
                            Viewport newVP = Viewport.Create(doc, existingView2.Id, curPlan.Id, new XYZ(0, 0, 0));
                        }
                        else
                        {
                            TaskDialog.Show("got here", "got here22"); //add more from 
                        }
                    }

                    //Set protected variables
                    /*
                    DrawnBy = drawnby;            //DrawnBy
                    CheckedBy = checkedby;        //CheckedBy
                    SheetDisc = sheetdisc;        //Discipline Header
                    SheetSort = sheetsort;        //Sort Order
                    */
                    //string paramValue = "";
                    foreach (Parameter curParam in curSheet.Parameters)
                    {
                        if (curParam.Definition.Name == "Drawn By")
                        {
                            curParam.Set(dataSheetList[i].DrawnBy);   //DrawnBy "MK"
                            //paramValue = curParam.AsString; //returning 
                        }
                        if (curParam.Definition.Name == "Checked By")
                        {
                            curParam.Set(dataSheetList[i].CheckedBy); //Chekced by "CC"
                            //paramValue = curParam.AsString; //returning 
                        }
                        if (curParam.Definition.Name == "Sheet Discipline")
                        {
                            curParam.Set(dataSheetList[i].SheetDisc); //
                            //paramValue = curParam.AsString; //returning 
                        }
                        if (curParam.Definition.Name == "Sort Order")
                        {
                            curParam.Set(dataSheetList[i].SheetSort); //
                            //paramValue = curParam.AsString; //returning 
                        }

                    }

                }

                t.Commit(); //Commit Transaction
            }


            excelWb.Close();
            excelApp.Quit();

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
