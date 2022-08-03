#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            // step 1: create ribbon tab
            try
            {
                a.CreateRibbonTab("Test Tab");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            // step 2: create ribbon panel        
            //RibbonPanel curPanel = a.CreateRibbonPanel("Test Tab", "Test Panel");
            RibbonPanel curPanel = CreateRibbonPanel(a, "Test Tab", "Test Panel");

            // step 3: create button data instances
            //use a method to get a
            PushButtonData pData1 = new PushButtonData("button1", "This is \n Button 1", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command"); //Get assembly name vs "RevitAddinAcademy.dll"
            PushButtonData pData2 = new PushButtonData("button2", "This is Button 2", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
            PushButtonData pData3 = new PushButtonData("button3", "This is Button 3", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
            PushButtonData pData4 = new PushButtonData("button4", "This is Button 4", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
            PushButtonData pData5 = new PushButtonData("button5", "This is Button 5", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
            PushButtonData pData6 = new PushButtonData("button6", "This is Button 6", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
            PushButtonData pData7 = new PushButtonData("button7", "This is Button 7", "RevitAddinAcademy.dll", "RevitAddinAcademy.Command");
 
            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "Pulldown \n Button 1");

            // step 4: add images
            
            pData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            pData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);

            pData2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_16);
            pData2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_32);

            pData3.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_16);
            pData3.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_32);

            pData4.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_16);
            pData4.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_32);

            pData5.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            pData5.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);

            pData6.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            pData6.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);

            pData7.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            pData7.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);

            pbData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            pbData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            
            // step 5: add tooltip info
            pData1.ToolTip = "Button 1 tool tip";
            pData2.ToolTip = "Button 1 tool tip";
            pData3.ToolTip = "Button 1 tool tip";
            pData4.ToolTip = "Button 1 tool tip";
            pData5.ToolTip = "Button 1 tool tip";
            pData6.ToolTip = "Button 1 tool tip";
            pData7.ToolTip = "Button 1 tool tip";
            pbData1.ToolTip = "Group of tools";



            // step 6: create buttons
            PushButton b1 = curPanel.AddItem(pData1) as PushButton;

            curPanel.AddStackedItems(pData2, pData3);
            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData4);
            splitButton1.AddPushButton(pData5);

            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData6);
            pulldownButton1.AddPushButton(pData7);

            return Result.Succeeded;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                {
                    return tmpPanel;
                }
            }
            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);

            return returnPanel;             
        }
  /*      
        private string GetAssemblyName()
        {
            //return Assembly.GetExecutingAssembly().Location;
        }
    */    

        private BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using(MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
