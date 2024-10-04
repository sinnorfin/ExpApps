using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows.Media.Imaging;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ToggleRed : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        GUI.togglebutton(uiapp, "Universal Modifiers", "Universal Toggle Red OFF", "Universal Toggle Red ON");
        return Result.Succeeded;
    }
}
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ToggleGreen : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        GUI.togglebutton(uiapp, "Universal Modifiers", "Universal Toggle Green OFF", "Universal Toggle Green ON");
        return Result.Succeeded;
    }
}
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ToggleBlue : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        GUI.togglebutton(uiapp, "Universal Modifiers", "Universal Toggle Blue OFF", "Universal Toggle Blue ON");
        return Result.Succeeded;
    }
}
public static class GUI
{
    public static void togglebutton(UIApplication uiapp, string panelname, string togglebutton_OFF, string togglebutton_ON)
    {
        RibbonPanel inputpanel = null;
        PushButton toggle = null;
        foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
        {
            if (panel.Name == panelname)
            {
                inputpanel = panel;
            }
        }
        foreach (RibbonItem item in inputpanel.GetItems())
        {
            if (item.Name == togglebutton_OFF)
            {
                toggle = (PushButton)item;
            }
        }
        string s = toggle.ToolTip;
        toggle.ToolTip = s.Equals(togglebutton_OFF) ? togglebutton_ON : togglebutton_OFF;
        //find and switch button and data 
        string IconsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Icons\\");
        string im_off = IconsPath + togglebutton_OFF + ".png"; string im_on = IconsPath + togglebutton_ON + ".png";
        if (toggle.ToolTip == togglebutton_OFF)
        {
            Uri uriImage = new Uri(im_off);
            BitmapImage Image = new BitmapImage(uriImage);
            toggle.Image = Image;
            toggle.ItemText = "OFF";
        }
        else
        {
            Uri uriImage = new Uri(im_on);
            BitmapImage Image = new BitmapImage(uriImage);
            toggle.Image = Image;
            toggle.ItemText = "ON";
        }
    }
}
public class StoreExp
    {
        public static List<View> quickViews = new List<View> { null, null, null, null, null, null };
        public static string level = "Active PlanView"; public static string ThreeDview = "Same Name";
        public static Double vrOpt1; public static Double vrOpt2; public static Double vrOpt3;
        public static Double vrOpt4; public static Double vrOpt5; public static Double vrOpt6;
        public static ICollection<ElementId> SelectionMemory = new List<ElementId> { };
        public static XYZ tag_shift = new XYZ(0,0,0); public static bool tag_leader = false;
        public static TagOrientation tag_orientation = TagOrientation.Horizontal;
        public static string Path_Insulation = null; public static string Path_IDs = null;
    public static string Path_BKVelocity = null;

    public static class Store
    //For storing values that are updated from the Ribbon
    {
        //public static Double menu_1 = 1;
        public static TextBox menu_1_Box = null;
        //public static Double menu_2 = 1;
        public static TextBox menu_2_Box = null;
        //public static Double menu_3 = 1;
        public static TextBox menu_3_Box = null;
        //public static Double menu_A = 1;
        public static TextBox menu_A_Box = null;
        //public static Double menu_B = 1;
        public static TextBox menu_B_Box = null;
        //public static Double menu_C = 1;
        public static TextBox menu_C_Box = null;
    }
    public static void GetMenuValue(UIApplication uiapp)
    {
        //Gets values out of Ribbon
        RibbonPanel inputpanel = null;
        foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
        {
            if (panel.Name == "Universal Modifiers")
            { inputpanel = panel; }
        }
        foreach (RibbonItem item in inputpanel.GetItems())
        {
            switch (item.Name)
            {
                case "1":
                    Store.menu_1_Box = (TextBox)item; break;
                case "2":
                    Store.menu_2_Box = (TextBox)item; break;
                case "3":
                    Store.menu_3_Box = (TextBox)item; break;
                case "A":
                    Store.menu_A_Box = (TextBox)item; break;
                case "B":
                    Store.menu_B_Box = (TextBox)item; break;
                case "C":
                    Store.menu_C_Box = (TextBox)item; break;
            }
        }
        //Double.TryParse(Store.menu_1_Box.Value as string, out Store.menu_1);
        //if (Store.menu_1 == 0) { Store.menu_1 = 0.5; }
        //Double.TryParse(Store.menu_2_Box.Value as string, out Store.menu_2);
        //Double.TryParse(Store.menu_3_Box.Value as string, out Store.menu_3);
    }
    public static Level GetLevel(Document doc, string levelName)
        {
            Level selectedLevel = null;
            foreach (Level level in new FilteredElementCollector(doc).OfClass(typeof(Level)))
            {
                if (level.Name == levelName)
                    selectedLevel = level;
            }
            return selectedLevel;
        }
        public static ComboBox GetComboBox(List<RibbonPanel> ribbonPanels, string panelName, string ribbonItemName)
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
        public static bool GetSwitchStance(UIApplication uiapp, string SwitchName)
        {
            //Gets stance of switches on of Ribbon
            RibbonPanel inputpanel = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "Universal Modifiers")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name.ToString().Contains(SwitchName))
                {
                    if (item.ItemText == "ON")
                    { return true; }
                    else
                    { return false; }
                }
            }
            return false;
        }
    }

