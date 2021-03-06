#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections;           //ArrayList, toChar and toInt
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

            string excelFile = @"C:\Users\LFarrell\Desktop\Revit Add-in Academy\Class Week 2\Session02_Challenge-220706-NoRoof.xlsx"; //pointed to new file

            Excel.Application excelApp = new Excel.Application();        //Open Application Excel
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); //Workbook File
            int NumWBSheets = excelWb.Sheets.Count;                      //Number of sheets to loop through
            List<List<string[]>> dataMultiList = new List<List<string[]>>(); //Collection of Lists outside loop

            //Get all data first. Then do revit actions.
            for (int i=1; i<=NumWBSheets; i++) //Loop through all WB sheets
            {
                //Read Excel - Transforming Rows and Columns to List of Arrays
                //Make ReadExcel Method later  //returns List type, takes Excel name, Range 

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
    }
}
