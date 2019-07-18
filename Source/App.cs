/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018 by Attila Kalina <attilakalina.arch@gmail.com>
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
        public static string level;
        public static Double vrOpt1; public static Double vrOpt2; public static Double vrOpt3;
        public static Double vrOpt4; public static Double vrOpt5; public static Double vrOpt6;
    }
    class App : IExternalApplication
    {
        public static UIControlledApplication UiCtrApp;
        public static Document doc;
        public static RibbonPanel panel_ViewSetup;
        public void SetTextBox(RibbonItem item,string name, string prompt,string tooltip,double width)
        {
            if (item.Name == name)
            {
                TextBox textBox = (TextBox)item;
                textBox.Width = width;
                textBox.PromptText = prompt; textBox.ToolTip = tooltip;
            }
        }
        public BitmapImage SetImage(string path)
        {
            Uri uriImage = new Uri(path);
            BitmapImage Image = new BitmapImage(uriImage);
            return Image;
        }
        public Result OnStartup(UIControlledApplication a)
        {
            UiCtrApp = a;
            try
            {
                // Register event for Syncronization 
               // a.ControlledApplication.DocumentSynchronizingWithCentral += new EventHandler
               //     <Autodesk.Revit.DB.Events.DocumentSynchronizingWithCentralEventArgs>(application_Sync);
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
            RibbonPanel panel_Managers = a.CreateRibbonPanel("Exp. Add-Ins", "Managers");
            RibbonPanel panel_Spec = a.CreateRibbonPanel("Exp. Add-Ins", "Specific");
            RibbonPanel panel_Qt = a.CreateRibbonPanel("Exp. Add-Ins", "Quick Tools");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            string root = thisAssemblyPath.Remove(thisAssemblyPath.Length - 11);

            PushButtonData button_DWGExport = new PushButtonData("Button_DWGExport", "Export DWG", root + "DWGExport.dll", "DWGExport.DWGExport");
            PushButtonData button_NavisExport = new PushButtonData("button_NavisExport", "Export NWC", root + "DWGExport.dll", "NavisExport.NavisExport");
            PushButtonData button_AllExport = new PushButtonData("Button_Navis_DWGExport", "Export All", root + "DWGExport.dll", "ExportAll.ExportAll");
            PushButton button_RevPrint = panel_Export.AddItem(new PushButtonData("Button_RevPrint", "Print Revision", root + "DWGExport.dll", "PrintRevision.PrintRevision")) as PushButton;

            button_DWGExport.ToolTip = "Selected Views/Sheets will be exported to .DWG. Tip: use Manage Save/Load selection for sets.";
            button_NavisExport.ToolTip = "Views with 'Export' as Title on Sheet will be exported to .NWC";
            button_AllExport.ToolTip = "Views with 'Export' as Title on Sheet will be exported to .DWG and .NWC";
            button_RevPrint.ToolTip = "Select and Print a certain Revision using the current print settings.";

            PushButtonData button_ShiftViewRange_Top_Up = new PushButtonData("Button_ShiftViewRange_Top_Up", "T+", root + "SetViewRange.dll",
                "SetViewRange.Shift_TU");
            PushButtonData button_ShiftViewRange_Top_Down = new PushButtonData("Button_ShiftViewRange_Top_Down", "T-", root + "SetViewRange.dll",
                "SetViewRange.Shift_TD");         
            PushButtonData button_ShiftViewRange_Bottom_Up = new PushButtonData("Button_ShiftViewRange_Bottom_Up", "B+", root + "SetViewRange.dll",
                "SetViewRange.Shift_BU");
            PushButtonData button_ShiftViewRange_Bottom_Down = new PushButtonData("Button_ShiftViewRange_Bottom_Down", "B-", root + "SetViewRange.dll",
                "SetViewRange.Shift_BD");
            PushButtonData button_SetViewRange = new PushButtonData("Button_SetViewRange", "Set View Range per 3D", root + "SetViewRange.dll",
                "SetViewRange.SetPer3D");

            button_ShiftViewRange_Top_Up.ToolTip = "Shifts Top of View Range by set value - Up.";
            button_ShiftViewRange_Top_Down.ToolTip = "Shifts Top of View Range by set value - Down.";
            button_ShiftViewRange_Bottom_Up.ToolTip = "Shifts Bottom of View Range by set value - Up.";
            button_ShiftViewRange_Bottom_Down.ToolTip = "Shifts Bottom of View Range by set value - Down.";
            button_SetViewRange.ToolTip = "Sets View Range of active Plan View to Top and Bottom planes of the Section Box used on identically named 3d View.";

            PushButtonData button_TL = new PushButtonData("ToggleLink", "Toggle Links", root + "MultiDWG.dll",
                "MultiDWG.ToggleLink");
            PushButtonData button_TPC = new PushButtonData("TogglePC", "Toggle PCs", root + "MultiDWG.dll",
                "MultiDWG.TogglePC");
            button_TL.ToolTip = "Toggles visibility of all Links in the active view";
            button_TPC.ToolTip = "Toggles visibility of all Point Clouds in the active view";

            PushButton button_RehostElements = panel_Reelevate.AddItem(new PushButtonData("Button_RehostElements", "Rehost Elements", root + "Rehost Elements.dll",
                "RehostElements.RehostElements")) as PushButton;
            PushButtonData button_AlignToBottom = new PushButtonData("Button_AlignToBottom", "MEP to Bottom", root + "AlignToBottom.dll",
                "AlignToBottom.AlignToBottom");
            PushButtonData button_AlignToTop = new PushButtonData("Button_AlignToTop", "MEP to Top", root + "AlignToBottom.dll",
                "AlignToBottom.AlignToTop");

            button_RehostElements.ToolTip = "Sets the Reference Level of selected elements to the active Plan View's Associated Level";
            button_AlignToBottom.ToolTip = "Aligns Bottom of MEP elements to the bottom of selected MEP element";
            button_AlignToTop.ToolTip = "Aligns Top of MEP elements to the top of selected MEP element";

            PushButtonData button_MultiDWG =new PushButtonData("Button_MultiDWG", "MultiDWG", root + "MultiDWG.dll",
               "MultiDWG.MultiDWG");
            PushButtonData button_CloneMeta = new PushButtonData("Button_CloneMeta", "Clone MetaData", root + "MultiDWG.dll",
               "MultiDWG.MetaData");
            PushButtonData button_TypeParam = new PushButtonData("Button_TypeParam", "TypeParam", root + "MultiDWG.dll",
               "MultiDWG.TypeParam");
            PushButtonData button_DuctSurfaceArea = new PushButtonData("Button_DSA", "Duct F. Unfolded", root + "MultiDWG.dll",
               "MultiDWG.DuctSurfaceArea");
            PushButtonData button_ConduitAngle = new PushButtonData("Button_CA", "Conduit Angles", root + "MultiDWG.dll",
               "MultiDWG.ConduitAngle");
            PushButtonData button_DuctId = new PushButtonData("Button_DI", "Duct Id", root + "MultiDWG.dll",
            "MultiDWG.DuctId");

            button_MultiDWG.ToolTip = "Specific: Loads all .DWG-s from selected folder. Sets LOD according to filename, temporarily hides medium and high LOD-s.";
            button_CloneMeta.ToolTip = "Specific: Copies Type-MetaData of selected instance into all type of the same family.";
            button_TypeParam.ToolTip = "Specific: Sets all instance parameters in Construction,Dimensions,and Visibility Parameter Groups. Might need to Run Multiple times";
            button_DuctSurfaceArea.ToolTip = "Specific: Inserts total surface area without connections into \"Duct Surface Area\" Project Parameter.";
            button_ConduitAngle.ToolTip = "Specific: Sums the angles of selected Conduit turns";
            button_DuctId.ToolTip = "Specific: Selected Ducts are grouped and Id-d based on System Abbreviation, Size and Length. Needs instance parameter for ducts: Manual ID ";

            panel_Spec.AddStackedItems(button_MultiDWG, button_CloneMeta,button_TypeParam);
            panel_Spec.AddStackedItems(button_DuctSurfaceArea,button_ConduitAngle,button_DuctId);

            PushButtonData toggle_Insulation = new PushButtonData("Toggle_Insulation", "Align to INS", root + "AlignToBottom.dll",
              "Toggle.Toggle");
            toggle_Insulation.ToolTip =
              "Aligns to Insulation surfaces, when present";

            PushButtonData button_Qv1 = new PushButtonData("Qv1",  "1", root + "SetViewRange.dll",
               "QuickViews.QuickView1");         
            PushButtonData button_Qv2 = new PushButtonData("Qv2", "2", root + "SetViewRange.dll",
               "QuickViews.QuickView2");
            PushButtonData button_Qv3 = new PushButtonData("Qv3", "3", root + "SetViewRange.dll",
               "QuickViews.QuickView3");
            PushButtonData button_Qv4 = new PushButtonData("Qv4", "4", root + "SetViewRange.dll",
               "QuickViews.QuickView4");
            PushButtonData button_Qv5 = new PushButtonData("Qv5", "5", root + "SetViewRange.dll",
               "QuickViews.QuickView5");
            PushButtonData button_Qv6 = new PushButtonData("Qv6", "6", root + "SetViewRange.dll",
               "QuickViews.QuickView6");

            button_Qv1.ToolTip = "Switch View to 'QuickView 1'"; button_Qv2.ToolTip = "Switch View to 'QuickView 2'";
            button_Qv3.ToolTip = "Switch View to 'QuickView 3'"; button_Qv4.ToolTip = "Switch View to 'QuickView 4'";
            button_Qv5.ToolTip = "Switch View to 'QuickView 5'"; button_Qv6.ToolTip = "Switch View to 'QuickView 6'";

            panel_ViewSetup.AddStackedItems(button_Qv1, button_Qv2, button_Qv3);
            panel_ViewSetup.AddStackedItems(button_Qv4, button_Qv5, button_Qv6);

            PushButton button_SetupQV = panel_ViewSetup.AddItem(new PushButtonData("Button_SetupQV", "Quick Views", root + "SetViewRange.dll",
           "QuickViews.QuickViews")) as PushButton;
            PushButtonData button_Dim2Grid = new PushButtonData("Dimtogrid", "RackDim", root + "AnnoTools.dll",
              "AnnoTools.RackDim");        
            PushButtonData button_Rack = new PushButtonData("Rack", "Rack", root + "AnnoTools.dll",
              "AnnoTools.Rack");     
            PushButtonData button_Lin = new PushButtonData("Lin", "Linear", root + "AnnoTools.dll",
             "AnnoTools.LinearAnnotation");
            string versioned = "AnnoTools";
            if (a.ControlledApplication.VersionName.Contains("2017"))
                { versioned = "Exp_apps_R17"; }
            PushButtonData button_MTag = new PushButtonData("MTag", "MultiTag", root + versioned + ".dll",
             versioned + ".MultiTag");
            PushButtonData button_ManageRefs = new PushButtonData("Button_ManageRefs", "Mng. Ref.Planes", root + "MultiDWG.dll",
             "MultiDWG.ManageRefPlanes");
            PushButtonData button_ManageRevs = new PushButtonData("Button_ManageRevs", "Mng. Revisions", root + "Revision_Editor.dll",
             "Revision_Editor.Revision_Editor");

            button_SetupQV.ToolTip = "Setup quick access to views.";
            button_Dim2Grid.ToolTip = "Create Dimension referring the selected element's centerlines and Grids.";
            button_Rack.ToolTip = "Create tag for Conduit Rack, listing conduits: left to right \\ top to bottom";
            button_Lin.ToolTip = "Create Dimension for objects with more distance in-between";
            button_MTag.ToolTip = "Create Tags by Category for multiple selected elements at once";
            button_ManageRefs.ToolTip = "Create Reference Planes from at the origins of 3 selected items, or Delete Ref.Planes";
            button_ManageRevs.ToolTip = "Manage Revisions";

            TextBoxData leftSpaceData = new TextBoxData("Left Space");
            TextBoxData rightSpaceData = new TextBoxData("Right Space");
            TextBoxData firstYData = new TextBoxData("First Y");
            TextBoxData stepYData = new TextBoxData("Step Y");
            TextBoxData splitPointData = new TextBoxData("Split Point");
            TextBoxData placementData = new TextBoxData("Placement");

            panel_Annot.AddStackedItems(leftSpaceData, rightSpaceData, firstYData);
            panel_Annot.AddStackedItems(stepYData, splitPointData,placementData);
            
            foreach (RibbonItem item in panel_Annot.GetItems())
            {
                SetTextBox(item, "Left Space", ":Left:", "Distance of TextBoxes on Left", 60);
                SetTextBox(item, "Right Space", ":Right:", "Distance of TextBoxes on Right", 60);
                SetTextBox(item, "First Y", ":Start:", "Height offset of TextBoxes", 60);
                SetTextBox(item, "Step Y", ":Step:", "Gap between TextBoxes", 60);
                SetTextBox(item, "Split Point", ":Split:", "Controls directional switch, and linebreaks of TextBoxes", 60);
                SetTextBox(item, "Placement", ":Place:", "Placement of annotation along the reference line", 60);
            }

            PulldownButtonData QtData = new PulldownButtonData("Quicktools","QuickTools");
            PulldownButton QtButtonGroup = panel_Qt.AddItem(QtData) as PulldownButton;

            PushButtonData qt1 = new PushButtonData("Filter Verticals", "Filter Vert", root + "MultiDWG.dll","MultiDWG.FindVert");
            PushButtonData qt2 = new PushButtonData("Filter Round Hosted", "Filter Round Hosted",root + "AnnoTools.dll", "AnnoTools.CheckTag");
            PushButtonData qt3 = new PushButtonData("Align Identicals", "Align Identicals", root + "AnnoTools.dll", "AnnoTools.Cleansheet");
            PushButtonData qt4 = new PushButtonData("Replace in Parameter", "Replace in Parameter", root + "MultiDWG.dll", "MultiDWG.ReplaceInParam");

            qt1.ToolTip = "Filters Vertical elements from selection (:Step: controls vertical sensitivity)";
            qt2.ToolTip = "Filter the selected tags that are hosted to Round duct - (type in :place: for 'Rectangular' filter)";
            qt3.ToolTip = "Merges selected tags with same content.(:Left: and :Right: controls sensitivity";
            qt4.ToolTip = "Replaces text in parameter of selection.(:Left: - Parameter name, :Right: - Original, :Start: - Replace)";

            string IconsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Icons\\");
            string[] files = Directory.GetFiles(IconsPath);

            string im_dwg = IconsPath + "Button_DWGExport.png"; string im_nwc = IconsPath + "Button_NavisExport.png";
            string im_all = IconsPath + "Button_AllExport.png"; string im_reh = IconsPath + "Button_RehostElements.png";
            string im_tup = IconsPath + "Button_T_UP.png"; string im_tdn = IconsPath + "Button_T_DN.png";
            string im_bup = IconsPath + "Button_B_UP.png"; string im_bdn = IconsPath + "Button_B_DN.png";
            string im_ins = IconsPath + "Button_Ins.png"; string im_pre = IconsPath + "Button_PrintRev.png";
            string im_pre_sm = IconsPath + "Button_PrintRev_sm.png"; string im_rak = IconsPath + "Button_Rack.png";
            string im_d2g_sm = IconsPath + "Button_Dim2Grid_sm.png"; string im_ref = IconsPath + "Button_Ref.png";
            string im_lin_sm = IconsPath + "Button_Lin_sm.png"; string im_mtag = IconsPath + "Button_MTag.png";
            string im_tl_sm = IconsPath + "Button_TL_sm.png"; string im_tpc_sm = IconsPath + "Button_TPC_sm.png";
            string im_qt1 = IconsPath + "Button_qt1.png"; string im_qt2 = IconsPath + "Button_qt2.png";
            string im_qt3 = IconsPath + "Button_qt3.png"; string im_qt4 = IconsPath + "Button_qt4.png";
            string im_rev = IconsPath + "Button_Rev.png";

            button_DWGExport.Image = SetImage(im_dwg);
            button_NavisExport.Image = SetImage(im_nwc);
            button_AllExport.Image = SetImage(im_all);
            button_ShiftViewRange_Top_Up.Image = SetImage(im_tup);
            button_ShiftViewRange_Top_Down.Image = SetImage(im_tdn);
            button_ShiftViewRange_Bottom_Up.Image = SetImage(im_bup);
            button_ShiftViewRange_Bottom_Down.Image = SetImage(im_bdn);
            button_RehostElements.Image = SetImage(im_reh);
            button_RehostElements.LargeImage = SetImage(im_reh);
            toggle_Insulation.Image = SetImage(im_ins);
            button_RevPrint.LargeImage = SetImage(im_pre);
            button_RevPrint.Image = SetImage(im_pre_sm);
            button_Rack.Image = SetImage(im_rak);
            button_Dim2Grid.Image = SetImage(im_d2g_sm);
            button_Lin.Image = SetImage(im_lin_sm);
            button_MTag.LargeImage = SetImage(im_mtag);
            button_ManageRefs.LargeImage = SetImage(im_ref);
            button_ManageRefs.Image = SetImage(im_ref);
            button_ManageRevs.LargeImage = SetImage(im_rev);
            button_ManageRevs.Image = SetImage(im_rev);
            button_TL.Image = SetImage(im_tl_sm);
            button_TPC.Image = SetImage(im_tpc_sm);

            qt1.LargeImage = SetImage(im_qt1);
            qt2.LargeImage = SetImage(im_qt2);
            qt3.LargeImage = SetImage(im_qt3);
            qt4.LargeImage = SetImage(im_qt4);

            QtButtonGroup.AddPushButton(qt1);
            QtButtonGroup.AddPushButton(qt2);
            QtButtonGroup.AddPushButton(qt3);
            QtButtonGroup.AddPushButton(qt4);

            ComboBoxData ShiftRange = new ComboBoxData("ShiftRange");

            panel_ViewSetup.AddStackedItems(button_ShiftViewRange_Bottom_Up, button_ShiftViewRange_Bottom_Down);
            panel_ViewSetup.AddStackedItems(button_ShiftViewRange_Top_Up, button_ShiftViewRange_Top_Down);
            panel_ViewSetup.AddStackedItems(button_TL, button_TPC);
            panel_Export.AddStackedItems(button_DWGExport, button_NavisExport, button_AllExport);
            panel_ViewSetup.AddStackedItems(button_SetViewRange,ShiftRange);
            panel_Reelevate.AddStackedItems(button_AlignToTop, button_AlignToBottom,toggle_Insulation);
            panel_Managers.AddItem(button_ManageRefs); panel_Managers.AddItem(button_ManageRevs);
            panel_Annot.AddStackedItems(button_Dim2Grid,button_Lin, button_Rack);
            panel_Annot.AddItem(button_MTag);

            a.ApplicationClosing += a_ApplicationClosing;

            //Set Application to Idling
            a.Idling += a_Idling;

            return Result.Succeeded;
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
        //public void application_Sync(object sender, DocumentSynchronizingWithCentralEventArgs args)
        //{
        //    Document doc = args.Document;
            //progress.Report(doc);
        //}
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