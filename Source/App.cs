/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018,2019,2020,2021 by Attila Kalina <attilakalina.arch@gmail.com>
 *
 * This file is part of Experimental Apps.
 * Exp Apps has been developed from June 2017 until end of March 2018 under the endorsement and for the use of hungarian BackOffice of Trimble VDC Services.
 * 
 * Experimental Apps is free software: you can redistribute 
 * it and/or modify it under the terms of the GNU General Public 
 * License as published by the Free Software Foundation, either 
 * version 3 of the License, or (at your option) any later version.
 * 
 * Some open source application is distributed in the hope that it will 
 * be useful, but WITHOUT ANY WARRANTY; without even the implied warranty 
 * of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with Foobar.  If not, see <http://www.gnu.org/licenses/>.
 *
 * @license GPL-3.0+ <https://www.gnu.org/licenses/gpl.html>
 */
using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace _ExpApps
{
    public class StoreExp
    {
        public static List<View> quickViews = new List<View> { null, null, null, null, null, null };
        public static string level = "Active PlanView"; public static string ThreeDview = "Same Name";
        public static Double vrOpt1; public static Double vrOpt2; public static Double vrOpt3;
        public static Double vrOpt4; public static Double vrOpt5; public static Double vrOpt6;
        public static ICollection<ElementId> SelectionMemory = new List<ElementId> { };
        public static Level GetLevel(FilteredElementCollector allLevels, string levelName)
        {
            Level selectedLevel = null;
            foreach (Level level in allLevels)
            {
                if (level.Name == levelName)
                    selectedLevel = level;
            }
            return selectedLevel;
        }
        public static ComboBox GetComboBox(List<RibbonPanel> ribbonPanels, string panelName, string ribbonItemName )
        {
            RibbonPanel inputpanel = null;
            ComboBox inputbox = null;
            foreach (RibbonPanel panel in ribbonPanels)
            {
                if (panel.Name == panelName)
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == ribbonItemName)
                { inputbox = (ComboBox)item; }
            }
            return inputbox;
        }
    }
    class App : IExternalApplication
    {
        public static UIControlledApplication UiCtrApp;
        public static Document doc;
        public static RibbonPanel panel_ViewSetup;
        public static string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
        public static string root = thisAssemblyPath.Remove(thisAssemblyPath.Length - 11);
        public enum IconImageType
        { None = 0, Largeimage, Noimage }
        public static string IconsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Icons\\");
        
        public PushButtonData CreateButton(string name, string assembly, string classname, IconImageType type = IconImageType.None, bool off = false)
        {
            PushButtonData pushButtonData = new PushButtonData(name, name, root + assembly, classname);
            if (type != IconImageType.Noimage)
            {
                string path = IconsPath + name + ".png";
                Uri uriImage = new Uri(path);
                BitmapImage image = new BitmapImage(uriImage);
                if (type == IconImageType.Largeimage) {
                    string pathlarge = IconsPath + name + "_L.png";
                    Uri uriImagelarge = new Uri(pathlarge);
                    BitmapImage imagelarge = new BitmapImage(uriImagelarge);
                    pushButtonData.Image = image;
                    pushButtonData.LargeImage = imagelarge;
                }
                else if (type == IconImageType.None) { pushButtonData.Image = image; }
            }
            if (off)
            {
                pushButtonData.Text = "OFF";
            }
            return pushButtonData;
        }
        public void SetTextBox(RibbonItem item,string name, string prompt,string tooltip,double width)
        {
            if (item.Name == name)
            {
                TextBox textBox = (TextBox)item;
                textBox.Width = width;
                textBox.PromptText = prompt; textBox.ToolTip = tooltip;
            }
        }
        public Result OnStartup(UIControlledApplication a)
        {
            UiCtrApp = a;
            try
            {
                UiCtrApp.ViewActivated += new EventHandler<Autodesk.Revit.UI.Events.ViewActivatedEventArgs>(getdoc);
            }
            catch (Exception)
            {
                return Result.Failed;
            }
            a.CreateRibbonTab("Exp. Add-Ins");
            RibbonPanel panel_Export = a.CreateRibbonPanel("Exp. Add-Ins", "Export");
            panel_ViewSetup = a.CreateRibbonPanel("Exp. Add-Ins", "View Tools");
            RibbonPanel panel_Reelevate = a.CreateRibbonPanel("Exp. Add-Ins", "Re-Elevate");
            RibbonPanel panel_Annot = a.CreateRibbonPanel("Exp. Add-Ins", "Annotation");
            RibbonPanel panel_Modifiers = a.CreateRibbonPanel("Exp. Add-Ins", "Universal Modifiers");
            RibbonPanel panel_Managers = a.CreateRibbonPanel("Exp. Add-Ins", "Managers");
            RibbonPanel panel_Selections = a.CreateRibbonPanel("Exp. Add-Ins", "Selections");
            RibbonPanel panel_Qt = a.CreateRibbonPanel("Exp. Add-Ins", "Quick Tools");
            ComboBoxData CB_ShiftRange = new ComboBoxData("ShiftRange");
            ComboBoxData CBD_ExpLevel = new ComboBoxData("ExpLevel");

            PushButtonData PBD_printrevision = CreateButton("Print Revision", "DWGExport.dll", "PrintRevision.PrintRevision",
                IconImageType.Largeimage);
            PBD_printrevision.ToolTip = "Select and Print a certain Revision using the current print settings.";

            PushButtonData PBD_shiftbu = CreateButton("B+", "SetViewRange.dll", "SetViewRange.Shift_BU");
            PBD_shiftbu.ToolTip = "Shifts Bottom of View Range by set value -Up.";

            PushButtonData PBD_shiftbd = CreateButton("B-",  "SetViewRange.dll", "SetViewRange.Shift_BD");
            PBD_shiftbd.ToolTip = "Shifts Bottom of View Range by set value - Down.";

            PushButtonData PBD_shifttu = CreateButton("T+", "SetViewRange.dll", "SetViewRange.Shift_TU");
            PBD_shiftbu.ToolTip = "Shifts Top of View Range by set value - Up.";

            PushButtonData PBD_shifttd = CreateButton("T-", "SetViewRange.dll", "SetViewRange.Shift_TD");
            PBD_shiftbd.ToolTip = "Shifts Top of View Range by set value - Down.";

            PushButtonData PBD_tl = CreateButton("Toggle Links", "MultiDWG.dll", "MultiDWG.ToggleLink",
                IconImageType.Noimage);
            PBD_tl.ToolTip = "Toggles visibility of all Links in the active view";

            PushButtonData PBD_tpc = CreateButton("Toggle PCs", "MultiDWG.dll", "MultiDWG.TogglePC",
                IconImageType.Noimage);
            PBD_tpc.ToolTip = "Toggles visibility of all Point Clouds in the active view";

            PushButtonData PBD_setviewrange = CreateButton("Set View Range per 3D", "SetViewRange.dll", "SetViewRange.SetPer3D", 
                IconImageType.Noimage);
            PBD_tpc.ToolTip = "Sets View Range of active Plan View to Top and Bottom planes of the Section Box used on identically named 3d View.";

            string versioned = "RehostElements";
            if (a.ControlledApplication.VersionName.Contains("201"))
            { versioned = "Exp_apps_R17"; }
            PushButtonData PBD_rehostelements = CreateButton("RehostElements", versioned + ".dll", versioned + ".RehostElements", 
                IconImageType.Largeimage);
            PBD_tpc.ToolTip = "Sets the Reference Level of selected elements to the active Plan View's Associated Level";

            PushButtonData PBD_dim2grid = CreateButton("Dim2Grid", "AnnoTools.dll", "AnnoTools.RackDim");
            PBD_dim2grid.ToolTip = "Create Dimension referring the selected element's centerlines and Grids.";
           
            PushButtonData PBD_rack = CreateButton("Rack", "AnnoTools.dll", "AnnoTools.Rack");
            PBD_rack.ToolTip = "Create tag for Conduit Rack, listing conduits: left to right \\ top to bottom";

            PushButtonData PBD_lin = CreateButton("Linear", "AnnoTools.dll", "AnnoTools.LinearAnnotation");
            PBD_lin.ToolTip = "Create Dimension for objects with more distance in-between";

            PushButtonData PBD_setupqv = CreateButton("Options", "SetViewRange.dll", "QuickViews.QuickViews",
                IconImageType.Noimage);
            PBD_setupqv.ToolTip = "Set quick access to views and more.";

            versioned = "AnnoTools";
            if (a.ControlledApplication.VersionName.Contains("2017"))
            { versioned = "Exp_apps_R17"; }
            PushButtonData PBD_mtag = CreateButton("MultiTag", versioned + ".dll", versioned + ".MultiTag", 
                IconImageType.Largeimage);
            PBD_mtag.ToolTip = "Create Tags by Category for multiple selected elements at once";
            
            PushButtonData PBD_managerefs = CreateButton("Mng. Ref.Planes", "MultiDWG.dll","MultiDWG.ManageRefPlanes", 
                IconImageType.Largeimage);
            PBD_managerefs.ToolTip = "Create Reference Planes from at the origins of 3 selected items, or Delete Ref.Planes";

            PushButtonData PBD_managerevs = CreateButton("Mng. Revisions", "Revision_Editor.dll","Revision_Editor.Revision_Editor", 
                IconImageType.Largeimage);
            PBD_managerevs.ToolTip = "Manage Revisions";

            PushButtonData PBD_memadd = CreateButton("Mem Add", "MultiDWG.dll", "MultiDWG.MemoryAdd");
            PBD_managerevs.ToolTip = "Add selected to stored selection";

            PushButtonData PBD_memdel = CreateButton("Mem Del", "MultiDWG.dll", "MultiDWG.MemoryDel");
            PBD_managerevs.ToolTip = "Clear stored selection";

            PushButtonData PBD_memsel = CreateButton("Mem Sel", "MultiDWG.dll", "MultiDWG.MemorySel");
            PBD_managerevs.ToolTip = "Select stored";

            PushButtonData PBD_qv1 = CreateButton("1", "SetViewRange.dll", "QuickViews.QuickView1",
                IconImageType.Noimage);
            PushButtonData PBD_qv2 = CreateButton("2", "SetViewRange.dll", "QuickViews.QuickView2",
                IconImageType.Noimage);
            PushButtonData PBD_qv3 = CreateButton("3", "SetViewRange.dll", "QuickViews.QuickView3",
                IconImageType.Noimage);
            PushButtonData PBD_qv4 = CreateButton("4", "SetViewRange.dll", "QuickViews.QuickView4",
                IconImageType.Noimage);
            PushButtonData PBD_qv5 = CreateButton("5", "SetViewRange.dll", "QuickViews.QuickView5",
                IconImageType.Noimage);
            PushButtonData PBD_qv6 = CreateButton("6", "SetViewRange.dll", "QuickViews.QuickView6",
                IconImageType.Noimage);

            PBD_qv1.ToolTip = "Switch View to 'QuickView 1'"; PBD_qv2.ToolTip = "Switch View to 'QuickView 2'";
            PBD_qv3.ToolTip = "Switch View to 'QuickView 3'"; PBD_qv4.ToolTip = "Switch View to 'QuickView 4'";
            PBD_qv5.ToolTip = "Switch View to 'QuickView 5'"; PBD_qv6.ToolTip = "Switch View to 'QuickView 6'";

            PushButtonData qt1 = CreateButton("Filter Vert", "MultiDWG.dll", "MultiDWG.FindVert",
                IconImageType.Largeimage);
            PushButtonData qt2 = CreateButton("Filter Round Hosted", "AnnoTools.dll", "AnnoTools.CheckTag",
                IconImageType.Largeimage);
            PushButtonData qt3 = CreateButton("Align Identicals", "AnnoTools.dll", "AnnoTools.Cleansheet",
                IconImageType.Largeimage);
            PushButtonData qt4 = CreateButton("Replace in Parameter", "MultiDWG.dll", "MultiDWG.ReplaceInParam",
                IconImageType.Largeimage);
            PushButtonData qt5 = CreateButton("Duplicate Sheets", "MultiDWG.dll", "MultiDWG.DuplicateSheets",
                IconImageType.Largeimage);

            qt1.ToolTip = "Filters Vertical elements from selection" + Environment.NewLine + ":1: controls vertical sensitivity";
            qt2.ToolTip = "Filter the selected tags that are hosted to Round duct" + Environment.NewLine + ":3: Type for 'Rectangular' filter";
            qt3.ToolTip = "Merges selected tags with same content." + Environment.NewLine + ":A: and :B: controls sensitivity";
            qt4.ToolTip = "Replaces text in parameter of selection." + Environment.NewLine + ":A: - Parameter name" + Environment.NewLine
                        + ":B: - Original" + Environment.NewLine + ":C: - Replace";
            qt5.ToolTip = "Duplicates the selected sheets." + Environment.NewLine + ":A: - Suffix - Sheet Number" + Environment.NewLine
                        + ":B: - Suffix - Sheet Name" + Environment.NewLine + ":C: - Type for Dependent view duplicates";
            panel_Export.AddItem(PBD_printrevision);
            panel_ViewSetup.AddStackedItems(PBD_shiftbu, PBD_shiftbd);
            panel_ViewSetup.AddStackedItems(PBD_shifttu, PBD_shifttd);
            panel_ViewSetup.AddStackedItems(PBD_tl, PBD_tpc);
            IList<RibbonItem> ribbonItemsStacked = panel_ViewSetup.AddStackedItems(PBD_setviewrange, CB_ShiftRange, CBD_ExpLevel);
            ComboBox CB_ExpLevel = (Autodesk.Revit.UI.ComboBox)(ribbonItemsStacked[2]);
            IList<ComboBoxMember> members = CB_ExpLevel.GetItems();
            for (int i = 0; i < 50; i++)
            {
                CB_ExpLevel.AddItem(new ComboBoxMemberData("Level_"+i, "x"));
            }
            foreach (ComboBoxMember member in members)
            { member.Visible = false; }
            CB_ExpLevel.DropDownOpened += cb_Opened;
            panel_ViewSetup.AddStackedItems(PBD_qv1, PBD_qv2, PBD_qv3);
            panel_ViewSetup.AddStackedItems(PBD_qv4, PBD_qv5, PBD_qv6);
            panel_ViewSetup.AddItem(PBD_setupqv);
            panel_Reelevate.AddItem(PBD_rehostelements);
            panel_Annot.AddStackedItems(PBD_dim2grid, PBD_lin, PBD_rack);
            panel_Managers.AddItem(PBD_managerefs);
            panel_Managers.AddItem(PBD_managerevs);
            panel_Selections.AddStackedItems(PBD_memadd, PBD_memdel, PBD_memsel);
            panel_Annot.AddItem(PBD_mtag);

            PulldownButtonData QtData = new PulldownButtonData("Quicktools", "QuickTools");
            PulldownButton QtButtonGroup = panel_Qt.AddItem(QtData) as PulldownButton;

            QtButtonGroup.AddPushButton(qt1);
            QtButtonGroup.AddPushButton(qt2);
            QtButtonGroup.AddPushButton(qt3);
            QtButtonGroup.AddPushButton(qt4);
            QtButtonGroup.AddPushButton(qt5);

            PushButtonData PBD_unitogglered = CreateButton("Universal Toggle Red OFF", "MultiDWG.dll",
              "MultiDWG.ToggleRed",off : true);
            PBD_unitogglered.ToolTip = "Universal Toggle Red OFF";
            PushButtonData PBD_unitogglegreen = CreateButton("Universal Toggle Green OFF", "MultiDWG.dll",
             "MultiDWG.ToggleGreen", off: true);
            PBD_unitogglegreen.ToolTip = "Universal Toggle Green OFF";
            PushButtonData PBD_unitoggleblue = CreateButton("Universal Toggle Blue OFF", "MultiDWG.dll",
             "MultiDWG.ToggleBlue", off: true);
            PBD_unitoggleblue.ToolTip = "Universal Toggle Blue OFF";

            TextBoxData Uni_A = new TextBoxData("A");
            TextBoxData Uni_B = new TextBoxData("B");
            TextBoxData Uni_C = new TextBoxData("C");
            TextBoxData Uni_1 = new TextBoxData("1");
            TextBoxData Uni_2 = new TextBoxData("2");
            TextBoxData Uni_3 = new TextBoxData("3");

            panel_Modifiers.AddStackedItems(Uni_A, Uni_B, Uni_C);
            panel_Modifiers.AddStackedItems(Uni_1, Uni_2, Uni_3);
            panel_Modifiers.AddStackedItems(PBD_unitogglered, PBD_unitogglegreen, PBD_unitoggleblue);
            foreach (RibbonItem item in panel_Modifiers.GetItems())
            {
                SetTextBox(item, "A", ":A:", "Universal Modifier - A", 60);
                SetTextBox(item, "B", ":B:", "Universal Modifier - B", 60);
                SetTextBox(item, "C", ":C:", "Universal Modifier - C", 60);
                SetTextBox(item, "1", ":1:", "Universal Modifier - 1", 60);
                SetTextBox(item, "2", ":2:", "Universal Modifier - 2", 60);
                SetTextBox(item, "3", ":3:", "Universal Modifier - 3", 60);
            }

            a.ApplicationClosing += a_ApplicationClosing;

            //Set Application to Idling
            a.Idling += a_Idling;

            return Result.Succeeded;
        }
        void cb_Opened(object sender, Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs e)
        {
            int c = 0;
            ComboBox cb = sender as ComboBox;
            IList<ComboBoxMember> members = cb.GetItems();
            foreach(ComboBoxMember member in members)
            { 
                member.Visible = false;
            }
            //{ comboBoxMember.Visible = false; }
            FilteredElementCollector allLevels = new FilteredElementCollector(e.Application.ActiveUIDocument.Document).OfClass(typeof(Level));
            //TaskDialog.Show("count_levels", allLevels.GetElementCount().ToString());
            foreach (Level level in allLevels)
            {
                members[c].Visible = true;
                members[c].ItemText = level.Name;
                c += 1;
            }
        }
        //*****************************a_Idling()*****************************
        void a_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
        }
        //*****************************a_ApplicationClosing()*****************************
        void a_ApplicationClosing(object sender, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs e)
        {
            throw new NotImplementedException();
        }
        public Result OnShutdown(UIControlledApplication a)
        {
            // a.ControlledApplication.DocumentSynchronizingWithCentral -= application_Sync;
            return Result.Succeeded;
        }

        public void getdoc(object sender, ViewActivatedEventArgs args)
        {
            doc = args.Document;
            ComboBoxMemberData Range1 = new ComboBoxMemberData("0", "1/2\"");
            ComboBoxMemberData Range2 = new ComboBoxMemberData("1", "1\"");
            ComboBoxMemberData Range3 = new ComboBoxMemberData("2", "3\"");
            ComboBoxMemberData Range4 = new ComboBoxMemberData("3", "1' 0\"");
            ComboBoxMemberData Range5 = new ComboBoxMemberData("4", "3' 0\"");
            ComboBoxMemberData Range6 = new ComboBoxMemberData("5", "10' 0\"");
            if (doc.DisplayUnitSystem == DisplayUnit.IMPERIAL)
            {  
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "1/2\"", out StoreExp.vrOpt1);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "1\"", out StoreExp.vrOpt2);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "3\"", out StoreExp.vrOpt3);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "1'0\"", out StoreExp.vrOpt4);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "3'0\"", out StoreExp.vrOpt5);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "10'0\"", out StoreExp.vrOpt6);
            }
            else
            {
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "1 cm", out StoreExp.vrOpt1);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "2 cm", out StoreExp.vrOpt2);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "10 cm", out StoreExp.vrOpt3);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "30 cm", out StoreExp.vrOpt4);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "90 cm", out StoreExp.vrOpt5);
                UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "300 cm", out StoreExp.vrOpt6);
                Range1.Text = "1 cm"; Range2.Text = "2 cm"; Range3.Text = "10 cm";
                Range4.Text = "30 cm"; Range5.Text = "90 cm"; Range6.Text = "300 cm";
            }
            UiCtrApp.ViewActivated -= getdoc;
            foreach (RibbonItem item in panel_ViewSetup.GetItems())
            {
                if (item.Name == "ShiftRange")
                {
                    ComboBox ShiftRange_CB = (ComboBox)item;
                    ShiftRange_CB.AddItem(Range1); ShiftRange_CB.AddItem(Range2);
                    ShiftRange_CB.AddItem(Range3); ShiftRange_CB.AddItem(Range4);
                    ShiftRange_CB.AddItem(Range5); ShiftRange_CB.AddItem(Range6);
                }
            }
            doc = null;
        }
    }
}