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
    public class cmdElementsFromLines : IExternalCommand
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some element");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            Level curLevel = GetLevelByName(doc, "Level 1");
            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");
            using (Transaction t = new Transaction(doc)) 
            {
                t.Start("Create Revit stuff");
                foreach (Element element in pickList)
                {
                    //Filter selection Curve, Line, point, etc.
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        CurveElement curve2 = element as CurveElement;
                        curveList.Add(curve);

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                        Curve curCurve = curve.GeometryCurve;
                        XYZ startpoint = curCurve.GetEndPoint(0); //argument zero is endpoint 1 of 2
                        XYZ endpoint = curCurve.GetEndPoint(1); //argument one is endpoint 2 of 2

                        /*
                        //SWITCH STATEMENT
                        switch (curGS.Name)
                        {
                            case "<Medium>":
                                Debug.Print("Found a medium line");
                                break;

                            case "<Thin lines>":
                                Debug.Print("Found a thine line");
                                break;

                            case "<Wide lines>":
                                Debug.Print("Found a wide line");
                                Pipe newPipe = Pipe.Create(
                                    doc,
                                    curSystemType.Id,
                                    curPipeType.Id,
                                    curLevel.Id,
                                    startpoint,
                                    endpoint);
                                break;

                            default:
                                Debug.Print("Found something else");
                                break;
                        }

                        */

                        //Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);
                        Pipe newPipe = Pipe.Create(
                            doc,
                            curSystemType.Id,
                            curPipeType.Id,
                            curLevel.Id,
                            startpoint,
                            endpoint);
                        
                        Debug.Print(curGS.Name);

                    } 
                }
                t.Commit();
            }
            TaskDialog.Show("Complete", curveList.Count.ToString());
                        
            return Result.Succeeded;
        }
        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));
            foreach(Element curElem in collector)
            {
                WallType wallType = curElem as WallType;
                if (wallType.Name == wallTypeName)
                    return wallType;
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
    }
}

