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
            if(dialog.Multiselect == true)    //switch between multiple and single files
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

            //DATA STRUCTURES
            //LEVELS
            LevelStruct levelData;           //3 in current data file, Levels is WB name
            levelData.Name = "Level Name";  //Level Names
            levelData.Elevation = 100.00;   //Level Elevations
            levelData.ElevationM = 10.5;    //Metric Elevations

            //SHEETS
            SheetStruct sheetData;           //5 in current data file, Sheets is WB name
            sheetData.SheetNumber = "A-101";     //Sheet Number
            sheetData.SheetName = "Sheet Name";  //Sheet Name
            sheetData.ViewLevel = "Level Name";  //View Level Name
            sheetData.DrawnBy   = "AA";          //DrawnBy
            sheetData.CheckedBy = "BB";          //CheckedBy
            sheetData.SheetDisc = "Architectural"; //Discipline Header
            sheetData.SheetSort = 5.0;           //Sort Order

            //How to call/initialize a struct; use in read excel.
            //List < SheetStruct >
            //TestStruct struct2 = new TestStruct("Structure 1", 10, 1004.4);
            //List<TestStruct> structList = new List<TestStruct>();
            //structList.Add(struct1);

            //single filepath
            string excelFile = filePath; //pointed to new file; 
            //multiple filepaths - loop through with list of strings - not used here.

            Excel.Application excelApp = new Excel.Application();        //Open Application Excel
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); //Workbook File
            int NumWBSheets = excelWb.Sheets.Count;                      //Number of sheets to loop through
            //List<List<string[]>> dataMultiList = new List<List<string[]>>(); //Collection of Lists outside loop
 
            //THIS IS WHERE WE SWITCH TO LIST OF STRUCTS INSTEAD OF STRINGS
            List<SheetStruct> dataSheetList = new List<SheetStruct>();
            List<LevelStruct> dataLevelList = new List<LevelStruct>();

            //Get all data first. Then do revit actions.
            for (int i=1; i<=NumWBSheets; i++) //Loop through all WB sheets
            {
                //Read Excel - Transforming Rows and Columns to List of Arrays
                //Make ReadExcel Method later  //returns Struct type, takes Excel name, Range 

                string excelWSName = excelWb.Sheets[i].Name;
                //if named Sheet - run sheet method
                if (excelWSName == "Sheets")
                {
                    dataSheetList = ReadExcelSheets(excelWb, i);    //code this method
                }
                //else if named Level - run levels method
                else if (excelWSName == "Levels")
                {
                    dataLevelList = ReadExcelLevels(excelWb, i);    //code this method
                }
                //Workbook Sheet - First sheet is 1 not 0
                Excel.Worksheet excelWs = excelWb.Worksheets.Item[i];    
                Excel.Range excelRng = excelWs.UsedRange;
                int rowCount = excelRng.Rows.Count;
                List<string[]> dataList = new List<string[]>(); //Collection of string arrays inside loop

                for (int j = 1; j <= rowCount; j++)
                {
                    Excel.Range cell1 = excelWs.Cells[j, 1]; //First cell of first column
                    Excel.Range cell2 = excelWs.Cells[j, 2]; //First cell of second column

                    string data1 = cell1.Value.ToString();
                    string data2 = cell2.Value.ToString();   //This makes ALL THE DATA Strings ****** conversion on the back end to doubles

                    string[] dataArray = new string[2];      //two elements in array
                    dataArray[0] = data1;
                    dataArray[1] = data2;
                    dataList.Add(dataArray);                 //Debug.Print("Data 1: " + data1.ToString());  //Check-in
                }
                dataMultiList.Add(dataList);                 //Add WBList to MultiList
            }

            //Do Revit Actions

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Sheet and View Setup from Excel"); //Start transaction
                bool doLevels = true;  //Start with Levels
                bool firstLine = true; //For skipping header rows

                //Setup Titleblock type
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks); //get Titleblock Type Category
                collector.WhereElementIsElementType();  //Types of titleblock types

                /*
                //FOR LEVELS
                List<string[]> subListLevels = dataMultiList[0]; //List copy
                int count1 = subListLevels.Count;
                for (int i=1; i<count1; ++i)
                {
                    //LEVELS
                    string strData1 = Convert.ToString(subListLevels[i][0]); //Level Name
                    double numData1 = Double.Parse(subListLevels[i][1]);
                    Level curLevel = Level.Create(doc, numData1); //create level - default imperial feet
                    curLevel.Name = strData1;
                }
                
                //FOR SHEETS
                List<string[]> subListSheets = dataMultiList[1]; //List copy
                int count2 = subListSheets.Count;
                for (int i= 1; i < count2; ++i)
                {
                    //SHEETS
                    string strData1 = subListSheets[i][0].ToString();  //Sheet Number
                    string strData2 = subListSheets[i][1].ToString();  //Sheet Name
                    ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId()); //uses first type of titleblock kind
                    curSheet.SheetNumber = strData1;
                    curSheet.Name = strData2;
                }
                */
                /* so many errors... =( */
                //for (int i = 1; i<=2; i++) //alt hardcoding it
                for (int i=1; i<=dataMultiList.Count; i++)
                {
                    List<string[]> subList = dataMultiList[i-1]; //List copy
                    for (int j=1; j<subList.Count; j++)
                    {
                        string[] value = subList[j-1]; 
                        if (firstLine)
                        {
                            //Skip header and set firstLine to false
                            if (value[0] == "Level Name")
                            {
                                doLevels = true; //set for levels going forward
                            }
                            else
                            {
                                doLevels = false; //set to sheets
                            }
                            firstLine = false;
                        }
                        else
                        {
                            if (doLevels) //make Levels
                            {
                                string strData1 = Convert.ToString(value[0]); //Level Name                                
                                double numData1 = Double.Parse(value[1]);     //Level Height
                               
                                Level curLevel = Level.Create(doc, numData1);     //create level - default imperial feet
                                if (curLevel == null)
                                {
                                    throw new Exception("Create new level failed. ");
                                }
                                curLevel.Name = strData1;
                            }
                            else  //make Sheets
                            {
                                  //element collector check against all existing names
                                  string strData1 = value[0].ToString();  //Sheet Number
                                  string strData2 = value[1].ToString();  //Sheet Name

                                  ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId()); //uses first type of titleblock kind
                                  if (null == curSheet)
                                  {
                                      throw new Exception("Create new sheet failed. ");
                                  }
                                  curSheet.SheetNumber = strData1;       
                                  curSheet.Name = strData2;       
                            }
                        }
                    }
                    doLevels = false; //set default back to false
                    firstLine = true; //reset for header row
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
        internal List<SheetStruct> ReadExcelSheets(Excel.Workbook excelWb, int i)                    //dataLevelList = ReadExcelLevels(excelWb, i);    //code this method
        {

            //initialize struct //set to new things in loop //leave these as blanks
            //SHEETS
            SheetStruct sheetData;           //5 in current data file, Sheets is WB name
            sheetData.SheetNumber = "A-101";     //Sheet Number
            sheetData.SheetName = "Sheet Name";  //Sheet Name
            sheetData.ViewLevel = "Level Name";  //View Level Name
            sheetData.DrawnBy = "AA";          //DrawnBy
            sheetData.CheckedBy = "BB";          //CheckedBy
            sheetData.SheetDisc = "Architectural"; //Discipline Header
            sheetData.SheetSort = 5.0;           //Sort Order

            //initialize list
            List<SheetStruct> structList = new List<SheetStruct>();

            //loop through WB range i j x
            //Workbook Sheet - First sheet is 1 not 0

            Excel.Worksheet excelWs = excelWb.Worksheets.Item[i];
            Excel.Range excelRng = excelWs.UsedRange;
            int rowCount = excelRng.Rows.Count;
            int colCount = excelRng.Columns.Count;
            //List<string[]> dataList = new List<string[]>(); //Collection of string arrays inside loop

            for (int j = 1; j <= rowCount; j++) //loop through cols
            {
                for (int k = 1; k <= colCount; k++) //loop through rows last one is double, so change
                {
                    Excel.Range cell1 = excelWs.Cells[j, k]; //Cell at first row and 1st --- 5th cell of first column, 
                    if (k == 7)
                    {
                        double data1 = cell1.Value.ToDouble();
                        sheetData.setDoubleAtIndex(data1, k); //Add double to struct at k
                    }
                    else
                    {
                        string data1 = cell1.Value.ToString();
                        sheetData.setStringAtIndex(data1, k); //Add string to struct at k
                    }

                }
                //Add row (struct data set) to struct list
                structList.Add(sheetData);
            }
            Debug.Print("Got here");  //Check-in

            return structList;
        }

        internal struct LevelStruct  //3 in current data file, Levels is WB name
        {
            //Define variables accessed from outside
            public string Name;
            public double Elevation;
            public double ElevationM;

            //constructor method
            //method inside structure that is named the same; specify arguments inside it.
            public LevelStruct(string name, double elevation, double elevationM)
            {
                Name = name;
                Elevation = elevation;
                ElevationM = elevationM;
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

            public SheetStruct (            //constructor
                string sheetnumber, 
                string sheetname, 
                string viewlevel, 
                string drawnby, 
                string checkedby, 
                string sheetdisc, 
                double sheetsort)
            {
                SheetNumber = sheetnumber;    //Sheet Number
                SheetName = sheetname;        //Sheet Name
                ViewLevel = viewlevel;        //View Level Name
                DrawnBy = drawnby;            //DrawnBy
                CheckedBy = checkedby;        //CheckedBy
                SheetDisc = sheetdisc;        //Discipline Header
                SheetSort = sheetsort;        //Sort Order
            }
            public void setStringAtIndex(string passedstring, int index)    //struct method
            {
                if      (index == 1) { SheetNumber = passedstring; }
                else if (index == 2) { SheetName   = passedstring; }
                else if (index == 3) { ViewLevel   = passedstring; }
                else if (index == 4) { DrawnBy     = passedstring; }
                else if (index == 5) { CheckedBy   = passedstring; }
                else if (index == 6) { SheetDisc   = passedstring; }
                return;
            }
            public void setDoubleAtIndex(double passedvalue, int index)    //struct method
            {
                if (index == 7) { SheetSort = passedvalue; }
                return;
            }
            public string addSuffix(string passedstring, string suffx)    //struct method
            {
                return (passedstring + suffx);
            }
            public string addPrefix(string passedstring, string prefx)    //struct method
            {
                return (prefx + passedstring);
            }
        }
    }
