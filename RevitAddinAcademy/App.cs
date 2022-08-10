#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection; //Mirror for Assembly Name
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
                a.CreateRibbonTab("Revit Add-in Academy");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            // step 2: create ribbon panel        
            //RibbonPanel curPanel = a.CreateRibbonPanel("Test Tab", "Test Panel");
            RibbonPanel curPanel = CreateRibbonPanel(a, "Revit Add-in Academy", "Campus Data");
            RibbonPanel curPanel2 = CreateRibbonPanel(a, "Revit Add-in Academy", "Send Elements to Campus Files");
            RibbonPanel curPanel3 = CreateRibbonPanel(a, "Revit Add-in Academy", "Move Elements in Campus Files");

            // step 3: create button data instances
            //use a method to get a
            PushButtonData pData1 = new PushButtonData("Tool1", "Load \n Campus", GetAssemblyName(), "RevitAddinAcademy.Command"); //Get assembly name vs "RevitAddinAcademy.dll"
            PushButtonData pData2 = new PushButtonData("Tool2", "Detail\n Out", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData3 = new PushButtonData("Tool3", "Parameter\n Out", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData4 = new PushButtonData("Tool4", "Template\n Out", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData5 = new PushButtonData("Tool5", "Drawing list PHS", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData6 = new PushButtonData("Tool6", "Align to Current Sheet", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData7 = new PushButtonData("Tool7", "Align Sim Sheets in Set", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData8 = new PushButtonData("Tool8", "Align Element on Sim Sheets", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData9 = new PushButtonData("Tool9", "Send Family", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData10 = new PushButtonData("ToolTen", "Send Material", GetAssemblyName(), "RevitAddinAcademy.Command");

            //Campus Sets
            PushButtonData cSet1 = new PushButtonData("CSet1", "Set A - 3, 11, 15, 21", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet2 = new PushButtonData("CSet2", "Set B - 5, 8, 17, 19", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet3 = new PushButtonData("CSet3", "Set C - 4, 9, 16, 20", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet4 = new PushButtonData("CSet4", "Set D - 12, 14, 22, 23", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet5 = new PushButtonData("CSet5", "Set E - 6, 18", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet6 = new PushButtonData("CSet6", "Set F - 2, 29", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet7 = new PushButtonData("CSet7", "Type 1", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet8 = new PushButtonData("CSet8", "Type 2", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet9 = new PushButtonData("CSet9", "Type 7", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet10 = new PushButtonData("CSet10", "Type 13", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet11 = new PushButtonData("CSet11", "Type 22A", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet12 = new PushButtonData("CSet12", "Type 25", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet13 = new PushButtonData("CSet13", "Type 26", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet14 = new PushButtonData("CSet14", "Type 27", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet15 = new PushButtonData("CSet15", "Type 28", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData cSet16 = new PushButtonData("CSet16", "Test File", GetAssemblyName(), "RevitAddinAcademy.cmdProjectSetup");

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Tools 6 and 7");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "Select Set");
            PulldownButtonData pbData2 = new PulldownButtonData("pulldownButton2", "More Tools");

            // step 4: add images

            pData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.C_IN_BUTTON_16);
            pData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.CAMPUS_IN_BUTTON_32);

            pData2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.D_OUT_BUTTON_16);
            pData2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.D_OUT_BUTTON_32);

            pData3.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.P_OUT_BUTTON_16);
            pData3.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.P_OUT_BUTTON_32);

            pData4.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.C_OUT_BUTTON_16);
            pData4.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.CAMPUS_OUT_BUTTON_32);

            pData5.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB4_16);
            pData5.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB4_32);

            pData6.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB5_16);
            pData6.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB5_32);

            pData7.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB6_16);
            pData7.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB6_32);

            pData8.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB3_16);
            pData8.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB3_32);

            pData9.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB2_16);
            pData9.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB2_32);

            pData10.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB1_16);
            pData10.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.DB1_32);

            pbData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PINATA_C_BUTTON_16);
            pbData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PINATA_C_BUTTON_32);

            pbData2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BOOKS_B_BUTTON_16);
            pbData2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BOOKS_B_BUTTON_32);

            //ADD CAMPUS IMAGES
            cSet1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);
            cSet1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);

            cSet2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);
            cSet2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);

            cSet3.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);
            cSet3.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);

            cSet4.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB4_16);
            cSet4.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB4_16);

            cSet5.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB5_16);
            cSet5.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB5_16);

            cSet6.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB6_16);
            cSet6.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB6_16);

            cSet7.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);
            cSet7.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);

            cSet8.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);
            cSet8.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);

            cSet9.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);
            cSet9.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);

            cSet10.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB4_16);
            cSet10.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB4_16);

            cSet11.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB5_16);
            cSet11.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB5_16);

            cSet12.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB6_16);
            cSet12.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB6_16);

            cSet13.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);
            cSet13.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB1_16);

            cSet14.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);
            cSet14.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB2_16);

            cSet15.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);
            cSet15.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.PB3_16);

            cSet16.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BC5_16);
            cSet16.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BC5_16);



            // step 5: add tooltip info
            pData1.ToolTip = "Select Campus excel file with Building File Paths";
            //"Select All Buildings in Campus";
            //"Select Type Set A, B, C...";
            pData2.ToolTip = "Push a Detail to Campus"; //D OUT
            pData3.ToolTip = "Push a Parameter to Campus"; //P OUT
            pData4.ToolTip = "Push a View Template to Campus"; // C OUT
            pData5.ToolTip = "Button 5 tool tip"; //Other Campus tools
            pData6.ToolTip = "Button 6 tool tip";
            pData7.ToolTip = "Button 7 tool tip";
            pData8.ToolTip = "Button 8 tool tip";
            pData9.ToolTip = "Button 9 tool tip"; 
            pData10.ToolTip = "Button 10 tool tip";
            pbData1.ToolTip = "Please Select Campus Set";
            pbData2.ToolTip = "Group of more tools";
            cSet1.ToolTip = "Set A";
            cSet2.ToolTip = "Set B";
            cSet3.ToolTip = "Set C";
            cSet4.ToolTip = "Set D";
            cSet5.ToolTip = "Set E";
            cSet6.ToolTip = "Set F";
            cSet7.ToolTip = "Type 1";
            cSet8.ToolTip = "Type 2";
            cSet9.ToolTip = "Type 7";
            cSet10.ToolTip = "Type 13";
            cSet11.ToolTip = "Type 22A";
            cSet3.ToolTip = "Type 25";
            cSet3.ToolTip = "Type 26";
            cSet3.ToolTip = "Type 27";
            cSet3.ToolTip = "Type 28";
            cSet3.ToolTip = "Test this on a selected file";

            // step 6: create buttons

            PushButton b1 = curPanel.AddItem(pData1) as PushButton;  // Tool 1 Get Excel Filenames
            AddRadioGroup(curPanel); //How to access current property RadioButtonGroup.Current

            //Select Campus Set
            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(cSet1);
            pulldownButton1.AddPushButton(cSet2);
            pulldownButton1.AddPushButton(cSet3);
            pulldownButton1.AddPushButton(cSet4);
            pulldownButton1.AddPushButton(cSet5);
            pulldownButton1.AddPushButton(cSet6);
            pulldownButton1.AddPushButton(cSet7);
            pulldownButton1.AddPushButton(cSet8);
            pulldownButton1.AddPushButton(cSet9);
            pulldownButton1.AddPushButton(cSet10);
            pulldownButton1.AddPushButton(cSet11);
            pulldownButton1.AddPushButton(cSet12);
            pulldownButton1.AddPushButton(cSet13);
            pulldownButton1.AddPushButton(cSet14);
            pulldownButton1.AddPushButton(cSet15);
            pulldownButton1.AddPushButton(cSet16);


            PushButton b2 = curPanel2.AddItem(pData2) as PushButton;  // Send D Out Detail
            PushButton b3 = curPanel2.AddItem(pData3) as PushButton;  // Send P Out Parameter
            PushButton b4 = curPanel2.AddItem(pData4) as PushButton;  // Send C Out Template

            //Stacks
            curPanel3.AddStackedItems(pData5, pData6, pData7); //Tools 5,6,7
            curPanel3.AddSeparator();
            curPanel3.AddStackedItems(pData8, pData9, pData10); //Tools 8,9,10

            SplitButton splitButton1 = curPanel3.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData6);
            splitButton1.AddPushButton(pData7);

            PulldownButton pulldownButton2 = curPanel3.AddItem(pbData2) as PulldownButton;
            pulldownButton2.AddPushButton(pData8);
            pulldownButton2.AddPushButton(pData9);
            pulldownButton2.AddPushButton(pData10);


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
 
        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }
    

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

        private void AddRadioGroup(RibbonPanel panel)
        {
            // add radio button group
            RadioButtonGroupData radioData = new RadioButtonGroupData("radioGroup");
            RadioButtonGroup radioButtonGroup = panel.AddItem(radioData) as RadioButtonGroup;

            // create toggle buttons and add to radio button group
            ToggleButtonData tb1 = new ToggleButtonData("toggleButton1", "All Bldgs");
            tb1.ToolTip = "All the buildings in the set, all types";            
            tb1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BC2_32);
            ToggleButtonData tb2 = new ToggleButtonData("toggleButton2", "Type Set");
            tb2.ToolTip = "All the buildings of the same Type";
            tb2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.BC5_32);

            radioButtonGroup.AddItem(tb1);
            radioButtonGroup.AddItem(tb2);

        }
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
