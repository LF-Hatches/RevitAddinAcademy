#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;  //new add

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

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog(); //dialog box
            dialog.InitialDirectory = @"C:\";   //initial directory
            dialog.Multiselect = false; //Single File vs multiple Files true
            dialog.Filter = "Revit Files | *.rvt; *.rfa";     //revit   if not specified, check file
            dialog.Filter = "Excel Files | *xlsx; *.xls *.xlxm"; //excel
            
            string filePath = "";
            string[] filePaths; //if multiple files
            //FolderBrowserDialog vs FileBrowserDialog
            //Forms.DialogResult = dialog.ShowDialog();
            if (dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                //filePath = dialog.FileName;
                filePaths = dialog.FileNames;
            }

            Forms.FolderBrowserDialog folderDialog = new Forms.FolderBrowserDialog();

            string folderPath = "";
            if(folderDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                folderPath = folderDialog.SelectedPath;
            }

            //Lists and Arrays (previously reviewed)
            //Tuple (multiple variables of different types)

            Tuple<string, int> t1 = new Tuple<string, int>("string 1", 55);
            Tuple<string, int> t2 = new Tuple<string, int>("string 2", 155);
            //t1.Item1;

            //Structure-Struct
            TestStruct struct1;
            struct1.Name = "Structure 1";
            struct1.Value = 100; 
            struct1.Value2 = 10.5;

            //double Num = struct1.Value + struct1.Value2;
            //calling a constructor
            TestStruct struct2 = new TestStruct("Structure 1", 10, 1004.4);
            double var = struct2.addNumber(); //calling constructor function

            List<TestStruct> structList = new List<TestStruct>();
            structList.Add(struct1);

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


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Stuff");

                Level newLevel = Level.Create(doc, 100);
                ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, newLevel.Id);
                curRCP.Name = curRCP.Name + "RCP";
                //Class clal
                View existingView = GetViewByName(doc, "Level 3");
                if(existingView 1= null)
                {
                    Viewport newVP = Viewport.Create(doc, newSheet.Id, curPlan.Id, new XYZ(0, 0, 0));
                }
                else
                {
                    TaskDialog.Show(); //add more from 
                }

                ViewSheet newSheet = ViewSheet.Create(doc, collector2.FirstElementId());
                Viewport newVP = Viewport.Create(doc, newSheet.Id, curPlan.Id, new XYZ(0, 0, 0));

                newSheet.Name = "TEST SHEET";
                newSheet.SheetNumber = "A10111";
                //drawn by or checked by not visible
                //collector parameters in API, 
                //loop through element/project parameters
                string paramValue = "";
                foreach(Parameter curParam in newSheet.Parameters)
                {
                    if(curParam.Definition.Name == "Drawn By")
                    {
                        curParam.Set("MK");
                        //paramValue = curParam.AsString; //returning 
                    }
                }



                t.Commit();
            }



            //Class

            return Result.Succeeded;
        }

        //Class
        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach(View curView in collector)
            {
                if(curView.Name == viewName)
                {
                    return curView;
                }
            }
            return null;
        }

        internal struct TestStruct 
        {
            //Define variables accessed from outside
            public string Name;
            public int Value;
            public double Value2;

            //constructor method
            //method inside structure that is named the same; specify arguments inside it.
            public TestStruct(string name, int value, double value2)
            {
                Name = name;
                Value = value;
                Value2 = value2;
            }
            public double addNumber()
            {
                return Value + Value2;
            }
        }
    }
}

