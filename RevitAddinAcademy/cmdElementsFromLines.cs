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
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;  //rooms

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdElementsFromLines : IExternalCommand  //rename for each new cmd file
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements");
            List<CurveElement> curveList = new List<CurveElement>();

            Level curLevel = GetLevelByName(doc, "Level 1");

            //A - GLAZ - Storefront wall
            //A - WALL - Generic 8" wall
            //M - DUCT - Default duct
            //P - PIPE - Default pipe

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8"""); //Generic 8" wall
            WallType curWallType2 = GetWallTypeByName(doc, @"Storefront");   //Storefront wall
            //CurtainSystemType curCurtainWall = GetCWTypeByName(doc, "Storefront"); //Storefront wall

            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water"); //domestic cold water, sanitary
            PipeType curPipeType = GetPipeTypeByName(doc, "Default"); //default pipe

            MEPSystemType curHVACType = GetSystemTypeByName(doc, "Supply Air"); //mechanical duct
            DuctType curDuctType = GetDuctTypeByName(doc, "Default"); //default duct

            int linecount = 0;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Generate Lines"); //Start transaction

                foreach (Element element in pickList)
                {
                    //Filter selection Curve, Line, point, etc.
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        //CurveElement curve = element as CurveElement;
                        //curveList.Add(curve);

                        Curve curCurve = curve.GeometryCurve;

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;


                        //SWITCH STATEMENT
                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Debug.Print("Found a storefront line");
                                Wall newSFWall = Wall.Create(
                                    doc,
                                    curCurve,
                                    curWallType2.Id,
                                    curLevel.Id,
                                    15,
                                    0,
                                    false,
                                    false);
                                ++linecount;
                                break;

                            case "A-WALL":
                                Debug.Print("Found a wall line");
                                Wall newWall = Wall.Create(
                                    doc, 
                                    curCurve, 
                                    curWallType.Id, 
                                    curLevel.Id, 
                                    15, 
                                    0, 
                                    false, 
                                    false);
                                ++linecount;
                                break;

                            case "M-DUCT":
                                Debug.Print("Found a duct line");
                                XYZ startpoint = curCurve.GetEndPoint(0); //argument zero is endpoint 1 of 2
                                XYZ endpoint = curCurve.GetEndPoint(1); //argument one is endpoint 2 of 2
                                Duct newDuct = Duct.Create(
                                    doc,
                                    curHVACType.Id,
                                    curDuctType.Id,
                                    curLevel.Id,
                                    startpoint,
                                    endpoint);
                                ++linecount;
                                break;
                            case "P-PIPE":
                                Debug.Print("Found a pipe line");
                                XYZ startpoint2 = curCurve.GetEndPoint(0); //argument zero is endpoint 1 of 2
                                XYZ endpoint2 = curCurve.GetEndPoint(1);   //argument one is endpoint 2 of 2
                                ++linecount;

                                Pipe newPipe = Pipe.Create(
                                    doc,
                                    curSystemType.Id,
                                    curPipeType.Id,
                                    curLevel.Id,
                                    startpoint2,
                                    endpoint2);

                                break;
                            default:
                                Debug.Print("Found something else");
                                break;
                        }
                        Debug.Print(curGS.Name);
                        Debug.Print("Linecount: " + linecount.ToString());
                    }
                }


                t.Commit(); //Commit Transaction
            }

            TaskDialog.Show("Completed", linecount.ToString());

            return Result.Succeeded;
        }



        //Class
        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));
            foreach (Element curElem in collector)
            {
                WallType wallType = curElem as WallType;
                if (wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }
        private CurtainSystemType GetCWTypeByName(Document doc, string cwallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(CurtainSystemType));
            foreach (Element curElem in collector)
            {
                CurtainSystemType cwallType = curElem as CurtainSystemType;
                if (cwallType.Name == cwallTypeName)
                    return cwallType;
            }
            return null;
        }
        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));
            foreach (Element curElem in collector)
            {
                Level level = curElem as Level;
                if (level.Name == levelName)
                    return level;
            }
            return null;
        }
        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));
            foreach (Element curElem in collector)
            {
                MEPSystemType curType = curElem as MEPSystemType;
                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));
            foreach (Element curElem in collector)
            {
                PipeType curType = curElem as PipeType;
                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
        private DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));
            foreach (Element curElem in collector)
            {
                DuctType curType = curElem as DuctType;
                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
    }
}