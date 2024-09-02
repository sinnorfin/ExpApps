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
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace _ExpApps
{
    class App : IExternalApplication
    {
        public enum IconImageType
        { None = 0, Largeimage, Noimage }
        public PushButtonData CreateButton(string name, string assembly, string classname, IconImageType type = IconImageType.None, bool off = false)
        {
            string IconsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Icons\\");
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            string root = thisAssemblyPath.Remove(thisAssemblyPath.Length - 11);
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
        public void SetTextBox(RibbonItem item, string name, string prompt, string tooltip, double width)
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
            UIControlledApplication UiCtrApp = a;
            a.CreateRibbonTab("Exp. Add-Ins");
            RibbonPanel panel_Export = a.CreateRibbonPanel("Exp. Add-Ins", "Export");
            RibbonPanel panel_ViewSetup = a.CreateRibbonPanel("Exp. Add-Ins", "View Tools");
            RibbonPanel panel_Reelevate = a.CreateRibbonPanel("Exp. Add-Ins", "Re-Elevate");
            RibbonPanel panel_Annot = a.CreateRibbonPanel("Exp. Add-Ins", "Annotation");
            RibbonPanel panel_Modifiers = a.CreateRibbonPanel("Exp. Add-Ins", "Universal Modifiers");
            RibbonPanel panel_Managers = a.CreateRibbonPanel("Exp. Add-Ins", "Managers");
            RibbonPanel panel_Selections = a.CreateRibbonPanel("Exp. Add-Ins", "Selections");
            RibbonPanel panel_Qt = a.CreateRibbonPanel("Exp. Add-Ins", "Quick Tools");
            ComboBoxData CBD_ShiftRange = new ComboBoxData("ShiftRange");
            ComboBoxData CBD_ExpLevel = new ComboBoxData("ExpLevel");

            PushButtonData PBD_printrevision = CreateButton("Print Revision", "DWGExport.dll", "PrintRevision.PrintRevision",
                IconImageType.Largeimage);
            PBD_printrevision.ToolTip = "Select and Print a certain Revision using the current print settings.";

            //*** SAVE A SET INSTEAD OF PRINT

            PushButtonData PBD_shiftbu = CreateButton("B+", "SetViewRange.dll", "SetViewRange.Shift_BU");
            PBD_shiftbu.ToolTip = "Shifts Bottom of View Range by set value -Up.";

            PushButtonData PBD_shiftbd = CreateButton("B-", "SetViewRange.dll", "SetViewRange.Shift_BD");
            PBD_shiftbd.ToolTip = "Shifts Bottom of View Range by set value - Down.";

            PushButtonData PBD_shifttu = CreateButton("T+", "SetViewRange.dll", "SetViewRange.Shift_TU");
            PBD_shiftbu.ToolTip = "Shifts Top of View Range by set value - Up.";

            PushButtonData PBD_shifttd = CreateButton("T-", "SetViewRange.dll", "SetViewRange.Shift_TD");
            PBD_shiftbd.ToolTip = "Shifts Top of View Range by set value - Down.";

            //*** NEEDS FIX, R20+

            PushButtonData PBD_tl = CreateButton("Toggle Links", "MultiDWG.dll", "MultiDWG.ToggleLink",
                IconImageType.Noimage);
            PBD_tl.ToolTip = "Toggles visibility of all Links in the active view";

            PushButtonData PBD_tpc = CreateButton("Toggle PCs", "MultiDWG.dll", "MultiDWG.TogglePC",
                IconImageType.Noimage);
            PBD_tpc.ToolTip = "Toggles visibility of all Point Clouds in the active view";

            //*** NEEDS FIX, R20+

            PushButtonData PBD_setviewrange = CreateButton("Set View Range per 3D", "SetViewRange.dll", "SetViewRange.SetPer3D",
                IconImageType.Noimage);
            PBD_tpc.ToolTip = "Sets View Range of active Plan View to Top and Bottom planes of the Section Box used on identically named 3d View.";

            string versioned = "RehostElements";
            if (a.ControlledApplication.VersionName.Contains("201"))
            { versioned = "Exp_apps_R17"; }
            PushButtonData PBD_rehostelements = CreateButton("RehostElements", versioned + ".dll", versioned + ".RehostElements",
                IconImageType.Largeimage);
            PBD_tpc.ToolTip = "Sets the Reference Level of selected elements to the active Plan View's Associated Level";

            PushButtonData PBD_setupqv = CreateButton("Options", "SetViewRange.dll", "QuickViews.QuickViews",
                IconImageType.Noimage);
            PBD_setupqv.ToolTip = "Set quick access to views and more.";

            versioned = "AnnoTools";
            if (a.ControlledApplication.VersionName.Contains("2017"))
            { versioned = "Exp_apps_R17"; }
            PushButtonData PBD_mtag = CreateButton("MultiTag", versioned + ".dll", versioned + ".MultiTag",
                IconImageType.Largeimage);
            PBD_mtag.ToolTip = "Create Tags by Category for multiple selected elements at once";

            PushButtonData PBD_managerefs = CreateButton("Mng. Ref.Planes", "MultiDWG.dll", "MultiDWG.ManageRefPlanes",
                IconImageType.Largeimage);
            PBD_managerefs.ToolTip = "Create Reference Planes from at the origins of 3 selected items, or Delete Ref.Planes";
            //*** MOVE TO SECONDARY TOOLS
            PushButtonData PBD_managerevs = CreateButton("Mng. Revisions", "Revision_Editor.dll", "Revision_Editor.Revision_Editor",
                IconImageType.Largeimage);
            PBD_managerevs.ToolTip = "Manage Revisions";
            //*** ADD ABILITY TO LOCATE / REMOVE CLOUDS

            PushButtonData PBD_memadd = CreateButton("Mem Add", "MultiDWG.dll", "MultiDWG.MemoryAdd");
            PBD_memadd.ToolTip = "Add selected to stored selection";

            PushButtonData PBD_memdel = CreateButton("Mem Del", "MultiDWG.dll", "MultiDWG.MemoryDel");
            PBD_memdel.ToolTip = "Clear stored selection";

            PushButtonData PBD_memsel = CreateButton("Mem Sel", "MultiDWG.dll", "MultiDWG.MemorySel");
            PBD_memsel.ToolTip = "Select stored";

            PushButtonData PBD_linkedId = CreateButton("Link-ID", "MultiDWG.dll", "MultiDWG.IdofLinkedElement", IconImageType.Noimage);
            PBD_linkedId.ToolTip = "Get ID of element in linked model";

            PushButtonData PBD_selanno = CreateButton("Anno", "MultiDWG.dll", "MultiDWG.OnlyAnnotation", IconImageType.Noimage);
            PBD_selanno.ToolTip = "Return all annotation elements - 'Red' - Return model elements instead";

            PushButtonData PBD_allonlevel = CreateButton("Level", "MultiDWG.dll", "MultiDWG.SelectAllOnLevel", IconImageType.Noimage);
            PBD_allonlevel.ToolTip = "Return elements on level specified in dropdown list - 'Red'+'A': Override level by name - 'Green': Invert";

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
            //*** Assign current view by shift + Number?

            PBD_qv1.ToolTip = "Switch View to 'QuickView 1'"; PBD_qv2.ToolTip = "Switch View to 'QuickView 2'";
            PBD_qv3.ToolTip = "Switch View to 'QuickView 3'"; PBD_qv4.ToolTip = "Switch View to 'QuickView 4'";
            PBD_qv5.ToolTip = "Switch View to 'QuickView 5'"; PBD_qv6.ToolTip = "Switch View to 'QuickView 6'";

            PushButtonData qt1 = CreateButton("Filter Vert", "MultiDWG.dll", "MultiDWG.FindVert",
                IconImageType.Noimage);
            PushButtonData qt2 = CreateButton("Filter Round Hosted", "AnnoTools.dll", "AnnoTools.CheckTag",
                IconImageType.Noimage);
            //*** Add option to invert
            PushButtonData qt3 = CreateButton("Apply Insulations", "MultiDWG.dll", "MultiDWG.ApplyInsulations",
                IconImageType.Noimage);
            PushButtonData qt4 = CreateButton("Replace in Parameter", "MultiDWG.dll", "MultiDWG.ReplaceInParam",
                IconImageType.Noimage);
            //*** Options
            PushButtonData qt5 = CreateButton("Duplicate Sheets", "MultiDWG.dll", "MultiDWG.DuplicateSheets",
                IconImageType.Noimage);
            PushButtonData qt6 = CreateButton("Toggle Tap/Tee", "MultiDWG.dll", "MultiDWG.SwitchTeeTap",
                IconImageType.Noimage);
            PushButtonData qt7 = CreateButton("Disconnect MEP", "MultiDWG.dll", "MultiDWG.ExplodeMEP",
                IconImageType.Noimage);
            PushButtonData qt8 = CreateButton("Duplicate Sheet for different level", "MultiDWG.dll", "MultiDWG.SheetForLevel",
                IconImageType.Noimage);
            PushButtonData qt9 = CreateButton("Place Selected views on Selected Sheet", "MultiDWG.dll", "MultiDWG.PlaceViewsonSheets",
           IconImageType.Noimage);
            PushButtonData qt10 = CreateButton("Shrink view-extents of Grids/Levels", "MultiDWG.dll", "MultiDWG.ShiftExtents",
           IconImageType.Noimage);
            PushButtonData qt11 = CreateButton("Clone Title properties", "MultiDWG.dll", "MultiDWG.ViewPortTitle",
           IconImageType.Noimage);
            PushButtonData qt12 = CreateButton("Check Flow in Selected", "MultiDWG.dll", "MultiDWG.BkFlowCheck",
   IconImageType.Noimage);
            PushButtonData qt13 = CreateButton("Weld MEP", "MultiDWG.dll", "MultiDWG.WeldMEP",
   IconImageType.Noimage);
            //*** No need for R23+
            qt1.ToolTip = "Filters Vertical elements from selection" + Environment.NewLine + ":1: controls vertical sensitivity";
            qt2.ToolTip = "Filter the selected tags that are hosted by Round duct" + Environment.NewLine + "'Red' - hosted by Rectangular";
            // ALIGN ALL ON VIEW** qt3.ToolTip = "Merges selected tags with same content." + Environment.NewLine + ":A: and :B: controls sensitivity";
            qt3.ToolTip = "Applies insulation to the systems of selected elements according to rules set." + Environment.NewLine +
                            ":RED: - Toggle to apply insulation where :1: parameter matches value of :2:" + Environment.NewLine +
                            ":RED + BLUE: - Toggle to apply insulation where :1: parameter DOES NOT match value of :2:" + Environment.NewLine +
                            ":GREEN: - Toggle to load new Rule-file" + Environment.NewLine;
            qt4.ToolTip = "Replaces text in parameter of selection." + Environment.NewLine + ":A: - Parameter name" + Environment.NewLine
                        + ":B: - Original" + Environment.NewLine + ":C: - Replace";
            qt5.ToolTip = "Duplicates the selected sheets." + Environment.NewLine + ":A: - Suffix - Sheet Number" + Environment.NewLine
                        + ":B: - Suffix - Sheet Name" + Environment.NewLine + ":C: - Type for Dependent view duplicates";
            qt6.ToolTip = "Adjusts the routing prefence to use Tee/Taps.";
            qt7.ToolTip = "Disconnects the selected MEP element from its neighbours, retains tags.";
            qt8.ToolTip = "Duplicates the selected sheets but for a different level, input negatives (-) for levels below" + Environment.NewLine + ":1: Difference in mm, between the target and source level" + Environment.NewLine
                       + ":2: Number of levels to step the source and target level" + Environment.NewLine + ":A: - Suffix - Sheet Number" + Environment.NewLine
                        + ":B: - Suffix - Sheet Name" + Environment.NewLine + ":C: - Type for Dependent view duplicates";
            qt9.ToolTip = "Select Views and Sheet in Project Browser." + Environment.NewLine + 
                        ":RED: - ON: Horizontal placement - OFF: Vertical placement";
            qt10.ToolTip = "Select Grids/Levels to adjust" + Environment.NewLine + "If Views are selected instead, adjustments will be applied on all grids/levels in selected views"
                + Environment.NewLine + ":RED: - Hide/Show bubbles" + Environment.NewLine + ":RED: + :GREEN: - Flip Bubbles" + Environment.NewLine +
                ":BLUE: - Reset to default" + Environment.NewLine + ":1: - Controls shrink distance";
             qt11.ToolTip = "Pick a Viewport to take Title position/ line length from" + Environment.NewLine + "note: Reference is bottom left corner of Bounding box of Viewport"
                + Environment.NewLine + "Bounds are affected by extensions of grids/levels, placed sections etc." + Environment.NewLine + " Works best if used on similar boundaries";
            qt12.ToolTip = "Check if Flow typed in parameter :A: equals the real flow from model." + Environment.NewLine + ":RED: - Copy 'Real' flow to 'Typed' Flow"
               + Environment.NewLine + "Gives Report of differences, copies involved ID-s to clipboard"
            + Environment.NewLine + ":GREEN: - No Report";
            qt13.ToolTip = "Connects selected MEP element to its properly aligned neighbours.";
            panel_Export.AddItem(PBD_printrevision);
            panel_ViewSetup.AddStackedItems(PBD_shiftbu, PBD_shiftbd);
            panel_ViewSetup.AddStackedItems(PBD_shifttu, PBD_shifttd);
            panel_ViewSetup.AddStackedItems(PBD_tl, PBD_tpc);
            IList<RibbonItem> ribbonItemsStacked = panel_ViewSetup.AddStackedItems(PBD_setviewrange, CBD_ShiftRange, CBD_ExpLevel);
            ComboBox CB_Distances = (Autodesk.Revit.UI.ComboBox)(ribbonItemsStacked[1]);
            IList<ComboBoxMember> member_distances = CB_Distances.GetItems();
            ComboBox CB_ExpLevel = (Autodesk.Revit.UI.ComboBox)(ribbonItemsStacked[2]);
            IList<ComboBoxMember> member_levels = CB_ExpLevel.GetItems();
            for (int i = 0; i < 5; i++)
            {
                CB_Distances.AddItem(new ComboBoxMemberData( i.ToString(), "Select Distance"));
            }
            for (int i = 0; i < 50; i++)
            {
                CB_ExpLevel.AddItem(new ComboBoxMemberData("Level_" + i, "Select Level"));
            }
            foreach (ComboBoxMember member in member_distances)
            { member.Visible = false;}
            CB_Distances.DropDownOpened += distances_cb_Opened;
            foreach (ComboBoxMember member in member_levels)
            {
                member.Visible = false;
            }
            CB_ExpLevel.DropDownOpened += levels_cb_Opened;

            panel_ViewSetup.AddStackedItems(PBD_qv1, PBD_qv2, PBD_qv3);
            panel_ViewSetup.AddStackedItems(PBD_qv4, PBD_qv5, PBD_qv6);
            panel_ViewSetup.AddItem(PBD_setupqv);
            panel_Reelevate.AddItem(PBD_rehostelements);
            panel_Managers.AddItem(PBD_managerefs);
            panel_Managers.AddItem(PBD_managerevs);
            panel_Selections.AddStackedItems(PBD_memadd, PBD_memdel, PBD_memsel);
            panel_Selections.AddStackedItems(PBD_linkedId, PBD_allonlevel, PBD_selanno);
            panel_Annot.AddItem(PBD_mtag);

            PulldownButtonData QtData = new PulldownButtonData("Quicktools", "QuickTools");
            PulldownButton QtButtonGroup = panel_Qt.AddItem(QtData) as PulldownButton;

            QtButtonGroup.AddPushButton(qt1);
            QtButtonGroup.AddPushButton(qt2);
            QtButtonGroup.AddPushButton(qt3);
            QtButtonGroup.AddPushButton(qt4);
            QtButtonGroup.AddPushButton(qt5);
            QtButtonGroup.AddPushButton(qt6);
            QtButtonGroup.AddPushButton(qt7);
            QtButtonGroup.AddPushButton(qt8);
            QtButtonGroup.AddPushButton(qt9);
            QtButtonGroup.AddPushButton(qt10);
            QtButtonGroup.AddPushButton(qt11);
            QtButtonGroup.AddPushButton(qt12);
            QtButtonGroup.AddPushButton(qt13);
            //Remove stance name from button name//
            PushButtonData PBD_unitogglered = CreateButton("Universal Toggle Red OFF", "StoreExp.dll",
              "ToggleRed", off: true);
            PBD_unitogglered.ToolTip = "Universal Toggle Red OFF";
            PushButtonData PBD_unitogglegreen = CreateButton("Universal Toggle Green OFF", "StoreExp.dll",
             "ToggleGreen", off: true);
            PBD_unitogglegreen.ToolTip = "Universal Toggle Green OFF";
            PushButtonData PBD_unitoggleblue = CreateButton("Universal Toggle Blue OFF", "StoreExp.dll",
             "ToggleBlue", off: true);
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
                SetTextBox(item, "A", ":A:", "Universal Modifier - A", 80);
                SetTextBox(item, "B", ":B:", "Universal Modifier - B", 80);
                SetTextBox(item, "C", ":C:", "Universal Modifier - C", 80);
                SetTextBox(item, "1", ":1:", "Universal Modifier - 1", 50);
                SetTextBox(item, "2", ":2:", "Universal Modifier - 2", 50);
                SetTextBox(item, "3", ":3:", "Universal Modifier - 3", 50);
            }

            a.ApplicationClosing += a_ApplicationClosing;

            //Set Application to Idling
            a.Idling += a_Idling;

            return Result.Succeeded;
        }
        public List<string> shiftdistances(Document doc)
        {
            
            List<string> distances = new List<string>();

            string Range1 = "1/2\"";
            string Range2 = "1\"";
            string Range3 = "3\"";
            string Range4 = "1' 0\"";
            string Range5 = "3' 0\"";
            string Range6 = "10' 0\"";
            if (doc.DisplayUnitSystem == DisplayUnit.IMPERIAL)
            {
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "1/2\"", out StoreExp.vrOpt1);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "1\"", out StoreExp.vrOpt2);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "3\"", out StoreExp.vrOpt3);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "1'0\"", out StoreExp.vrOpt4);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "3'0\"", out StoreExp.vrOpt5);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "10'0\"", out StoreExp.vrOpt6);
            }
            else
            {
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "1 cm", out StoreExp.vrOpt1);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "2 cm", out StoreExp.vrOpt2);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "10 cm", out StoreExp.vrOpt3);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "30 cm", out StoreExp.vrOpt4);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "90 cm", out StoreExp.vrOpt5);
                UnitFormatUtils.TryParse(doc.GetUnits(), SpecTypeId.Length, "300 cm", out StoreExp.vrOpt6);
                Range1 = "1 cm"; Range2 = "2 cm"; Range3 = "10 cm";
                Range4 = "30 cm"; Range5 = "90 cm"; Range6 = "300 cm";
            }
            distances.Add(Range1); distances.Add(Range2); distances.Add(Range3);
            distances.Add(Range4); distances.Add(Range5); distances.Add(Range6);

            return distances;
        }
        void distances_cb_Opened(object sender, Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs e)
        {
            int c = 0;
            ComboBox cb = sender as ComboBox;
            IList<ComboBoxMember> members = cb.GetItems();
            foreach (ComboBoxMember member in members)
            {
                member.Visible = false;
            }
            Document doc = e.Application.ActiveUIDocument.Document;
            string version = doc.Application.VersionName;
            List<string> distances = new List<string>();
            if (version.Contains("202") && !version.Contains("2021") && !version.Contains("2020"))
            { distances = shiftdistances(doc); }
            else { distances = Versioned_methods.shiftdistances(doc); }
            foreach (string distance in distances)
            {
                members[c].Visible = true;
                members[c].ItemText = distance;
                c += 1;
            }
        }
        void levels_cb_Opened(object sender, Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs e)
        {
            int c = 0;
            ComboBox cb = sender as ComboBox;
            IList<ComboBoxMember> members = cb.GetItems();
            foreach (ComboBoxMember member in members)
            {
                member.Visible = false;
            }
            FilteredElementCollector allLevels = new FilteredElementCollector(e.Application.ActiveUIDocument.Document).OfClass(typeof(Level));
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
            return Result.Succeeded;
        }

    }
}