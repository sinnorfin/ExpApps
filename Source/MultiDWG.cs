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

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using System.Collections.Generic;
using System;
using System.Windows.Media.Imaging;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using _ExpApps;
//using Autodesk.Revit.DB.Mechanical;

namespace MultiDWG
{
    public static class GUI
    {
        public static void togglebutton(UIApplication uiapp,string panelname, string togglebutton_OFF, string togglebutton_ON)
        {
            RibbonPanel inputpanel = null;
            PushButton toggle = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == panelname)
                { inputpanel = panel;
                }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == togglebutton_OFF)
                { toggle = (PushButton)item;
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
    public static class Store
    //For storing values that are updated from the Ribbon
    {
        public static Double menu_1 = 1;
        public static TextBox menu_1_Box = null;
        public static Double menu_2 = 1;
        public static TextBox menu_2_Box = null;
        public static Double menu_3 = 1;
        public static TextBox menu_3_Box = null;
        public static Double menu_A = 1;
        public static TextBox menu_A_Box = null;
        public static Double menu_B = 1;
        public static TextBox menu_B_Box = null;
        public static Double menu_C = 1;
        public static TextBox menu_C_Box = null;
    }

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
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ToggleLink : IExternalCommand
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
            View acview = doc.ActiveView;
            foreach (Category cat in acview.Document.Settings.Categories)
            {
                if (cat.Name == "RVT Links")
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("Toggle All Links");
                        if (cat.get_Visible(acview))
                        { cat.set_Visible(acview, false); }
                        else { cat.set_Visible(acview, true); }
                        trans.Commit();
                    }
                }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TogglePC : IExternalCommand
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
            View acview = doc.ActiveView;
            foreach (Category cat in acview.Document.Settings.Categories)
            {
                if (cat.Name == "Point Clouds")
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("Toggle All Point Cloud");
                        if (cat.get_Visible(acview))
                        { cat.set_Visible(acview, false); }
                        else { cat.set_Visible(acview, true); }
                        trans.Commit();
                    }
                }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuctSurfaceArea : IExternalCommand
    {
        //Calculates Duct surface area, substracts connection surfaces
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Options geOpt = new Options();
            FindVert.GetMenuValue(uiapp);
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Calculate Duct Surface Area");
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    FamilyInstance fami = elem as FamilyInstance;
                    MEPModel mepModel = fami.MEPModel;
                    double connectorAreas = 0;
                    double surfaceAreas = 0;
                    int ConnectorCount = 0;
                    // siwtch to mepModel.ConnectorManager.Connectors.Size for counting connectors
                    foreach (Connector connector in mepModel.ConnectorManager.Connectors)
                    {
                        ConnectorCount += 1;
                        try {
                            connectorAreas += (Math.Pow(connector.Radius, 2) * Math.PI);

                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            connectorAreas += (connector.Width * connector.Height);
                        }
                    }
                    double parts = 0;
                    GeometryElement geoelem = elem.get_Geometry(geOpt);
                    foreach (GeometryInstance geoInst in geoelem)
                    {
                        foreach (var solid in geoInst.GetInstanceGeometry())
                        {
                            try
                            {
                                Solid ssolid = solid as Solid;
                                if ((ssolid.SurfaceArea > 0) && ((ssolid.SurfaceArea / ssolid.Volume) < 80))
                                {
                                    surfaceAreas += ssolid.SurfaceArea;
                                    parts += 1;
                                }
                            }
                            catch (System.NullReferenceException)
                            { continue; }
                        }
                    }
                    double total = surfaceAreas - connectorAreas * parts;
                    elem.LookupParameter("Duct Surface Area").Set(total);
                    // CHECK NUMBER OF PARTS
                    if (Store.menu_A_Box.Value.ToString() != "") { TaskDialog.Show("Report", "Parts counted: " + parts); }
                    // TO ONLY INCLUDE CONNECTOR SIZE ON ENDCAPS
                    if (ConnectorCount == 1) { elem.LookupParameter("Duct Surface Area").Set(connectorAreas); }
                    // TO USE BUILT-IN AREA FOR ROUND REDUCERS
                    if (parts == 5 || Store.menu_B_Box.Value.ToString() != "")
                    {
                        FamilySymbol famsim = doc.GetElement(elem.GetTypeId()) as FamilySymbol;
                        try { elem.LookupParameter("Duct Surface Area").Set(famsim.LookupParameter("Duct Area").AsDouble()); }
                        catch { elem.LookupParameter("Duct Surface Area").Set(elem.LookupParameter("Duct Area").AsDouble()); }
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MultiDWG : IExternalCommand
    {
        //Loads all DWG from a directory at once
        //adds DWS-s to detail level based on name
        //hides med and high details levels
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Load Multiple DWGs");
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {
                    DWGImportOptions ImportOptions = new DWGImportOptions
                    {
                        Placement = ImportPlacement.Origin,
                        ColorMode = ImportColorMode.BlackAndWhite
                    };
                    string filepath = dlg.FileName;
                    filepath = filepath.Remove(dlg.FileName.LastIndexOf(@"\"));
                    Array files = System.IO.Directory.GetFiles(filepath);
                    foreach (string e in files)
                    {
                        doc.Import(e, ImportOptions, doc.ActiveView, out ElementId elementid);
                        Element elem = doc.GetElement(elementid);
                        Parameter LOD = elem.get_Parameter(BuiltInParameter.GEOM_VISIBILITY_PARAM);
                        int LODvis = LOD.AsInteger();
                        if (e.Contains("low") || (e.Contains("LOW")))
                        {
                            LODvis = LODvis & ~(1 << 14);
                            LODvis = LODvis & ~(1 << 15);
                        }
                        else if (e.Contains("med") || (e.Contains("MED")))
                        {
                            LODvis = LODvis & ~(1 << 13);
                            LODvis = LODvis & ~(1 << 15);
                        }
                        else if (e.Contains("fine") || (e.Contains("FINE")))
                        {
                            LODvis = LODvis & ~(1 << 13);
                            LODvis = LODvis & ~(1 << 14);
                        }
                        LOD.Set(LODvis);
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MemoryAdd : IExternalCommand
    {
        //Stores Selection

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = _ExpApps.StoreExp.SelectionMemory;
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Store Selection in Memory");
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid) as Element;
                    try
                    {
                        newsel.Add(eid);
                    }
                    catch { }
                }
                _ExpApps.StoreExp.SelectionMemory = newsel;
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MemorySel : IExternalCommand
    {
        //Selects stored Selection

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            uidoc.Selection.SetElementIds(_ExpApps.StoreExp.SelectionMemory);
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MemoryDel : IExternalCommand
    {
        //Selects stored Selection

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            _ExpApps.StoreExp.SelectionMemory = new List<ElementId>();
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ReplaceInParam : IExternalCommand
    {
        //Replaces text in a parameter of selected elements

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Replace in Parameter");
                double c = 0;
                double x = 0;
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid) as Element;
                    Parameter para;
                    try
                    {
                        para = elem.LookupParameter(Store.menu_A_Box.Value.ToString()) as Parameter;
                        string original = para.AsString();
                        if (Store.menu_B_Box.Value.ToString() == "*add")
                        {
                            original += Store.menu_C_Box.Value.ToString();
                            para.Set(original); }
                        else if (Store.menu_B_Box.Value.ToString() == "add*")
                        { para.Set(Store.menu_C_Box.Value.ToString() + original); }
                        else
                        { para.Set(original.Replace(Store.menu_B_Box.Value.ToString(), Store.menu_C_Box.Value.ToString())); }
                        if (para.AsString() != original)
                        {
                            c += 1; newsel.Add(eid);
                        }
                    }
                    catch { x += 1; }
                }
                trans.Commit();
                uidoc.Selection.SetElementIds(newsel);
                string text = "Replaced '" + Store.menu_B_Box.Value.ToString()
                              + "' to '" + Store.menu_C_Box.Value.ToString()
                              + "' in " + c.ToString() + " elements";
                if (c == 0) { text = "No replacement occurred"; }
                if (x > 0) { text += Environment.NewLine + "No such parameter: " + x.ToString(); }
                TaskDialog.Show("Result", text);
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class InjectParam : IExternalCommand
    {
        //Inject parameter value to target parameter

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            FindVert.GetMenuValue(uiapp);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Inject Parameter");
                double c = 0;
                double x = 0;
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid) as Element;
                    Parameter targetPara;
                    Parameter sourcePara;
                    try
                    {
                        targetPara = elem.LookupParameter(Store.menu_B_Box.Value.ToString()) as Parameter;
                        sourcePara = elem.LookupParameter(Store.menu_A_Box.Value.ToString()) as Parameter;


                        if (Store.menu_C_Box.Value.ToString() == "S")
                        {
                            targetPara.Set(sourcePara.AsString());
                        }
                        else if (Store.menu_C_Box.Value.ToString() == "VS")
                        {
                            targetPara.Set(sourcePara.AsValueString());
                        }
                        else if (Store.menu_C_Box.Value.ToString() == "Num")
                        {
                            double orig;
                            Double.TryParse(sourcePara.AsValueString(), out orig);
                            orig = UnitUtils.Convert(orig, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                            targetPara.Set(orig);
                        }
                        else if (Store.menu_C_Box.Value.ToString() != "")
                        {
                            double orig;
                            double oper;
                            Double.TryParse(sourcePara.AsString(), out orig);
                            Double.TryParse(Store.menu_C_Box.Value.ToString(), out oper);
                            double sum = orig + oper;
                            targetPara.Set(sum.ToString());
                        }
                        c += 1;
                    }
                    catch { x += 1; }
                }
                trans.Commit();
                string text = "Replaced '" + Store.menu_A_Box.Value.ToString()
                              + "' to '" + Store.menu_B_Box.Value.ToString()
                              + "' in " + c.ToString() + " elements";
                if (c == 0) { text = "No replacement occurred"; }
                if (x > 0) { text += Environment.NewLine + "No such parameter: " + x.ToString(); }
                TaskDialog.Show("Result", text);
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RecessHeight : IExternalCommand
    {
        //Replaces text in a parameter of selected elements

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Recess Heights");
                double c = 0;
                double r = 0;
                double x = 0;
                Level HostLevel = null;
                bool setLevel = true;
                double HostLevelElevation = 0;
                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                ICollection<Element> lvlCollection = lvlCollector.OfClass(typeof(Level)).ToElements();

                if (Store.menu_1_Box.Value.ToString() != "")
                {
                    setLevel = false;
                    foreach (Element level in lvlCollection)
                    {
                        Level lvl = level as Level;
                        if (level.Name == StoreExp.level)
                        {
                            HostLevel = lvl;
                            HostLevelElevation = HostLevel.ProjectElevation; //Feet
                        }
                    }
                }
                foreach (ElementId eid in ids)
                {
                    //Custom Inject Parameter

                    // String TargetParaString = "Comments";
                    Double Resultelevation;
                    Element elem = doc.GetElement(eid);
                    //Get host Level's elevation
                    if (setLevel)
                    {
                        FamilyInstance fami = elem as FamilyInstance;
                        ElementId HostLevelId = fami.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).AsElementId();
                        HostLevel = doc.GetElement(HostLevelId) as Level;
                        HostLevelElevation = HostLevel.ProjectElevation; //Feet
                    }
                    //Get Element Geom Center elevation
                    LocationPoint ElemLocation = elem.Location as LocationPoint;
                    Double ElemCenterElevation = ElemLocation.Point.Z;
                    //Determine Center Elevation
                    Resultelevation = ElemCenterElevation - HostLevelElevation;
                    string familyName = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();

                    // Custom Inject Parameter

                    //if (Store.menu_3_Box.Value.ToString() != "")
                    //{
                    //    TargetParaString = Store.menu_3_Box.Value.ToString();
                    //}
                    try
                    {
                        Parameter para_inject;
                        //para_inject = elem.LookupParameter(TargetParaString) as Parameter;

                        if (familyName.Contains("Circ"))
                        {
                            para_inject = elem.LookupParameter("Ass. Level Elevation Center") as Parameter;
                            para_inject.Set(Resultelevation);
                            double recess_height = elem.LookupParameter("Recess Diameter").AsDouble();
                            double BottomResultelevation = Resultelevation - (recess_height / 2);
                            elem.LookupParameter("Ass. Level Elevation Bottom").Set(BottomResultelevation);
                            c += 1;
                        }
                        else if (familyName.Contains("Rect"))
                        {
                            double recess_height = elem.LookupParameter("Recess Height").AsDouble();
                            //recess_height = Math.Round(recess_height, 0);
                            elem.LookupParameter("Ass. Level Elevation Center").Set(Resultelevation);
                            double BottomResultelevation = Resultelevation - (recess_height / 2);
                            para_inject = elem.LookupParameter("Ass. Level Elevation Bottom") as Parameter;
                            para_inject.Set(BottomResultelevation);
                            r += 1;
                        }
                    }
                    catch { x += 1; newsel.Add(elem.Id); }
                }
                trans.Commit();
                uidoc.Selection.SetElementIds(newsel);
                string text = "Circular Recess updated: " + c.ToString()
                              + Environment.NewLine + "Rectangular Recess updated: " + r.ToString()
                              + Environment.NewLine + "Invalid Elements (selected) :" + x.ToString();
                TaskDialog.Show("Result", text);
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ConvertRCFtoFloorPlan : IExternalCommand
    {
        //Replaces text in a parameter of selected elements
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            FindVert.GetMenuValue(uiapp);
            List<List<ElementId>> sheetsandview = new List<List<ElementId>>();
            FamilySymbol fs = new FilteredElementCollector(doc)
           .OfClass(typeof(FamilySymbol))
           .OfCategory(BuiltInCategory.OST_TitleBlocks)
           .FirstOrDefault() as FamilySymbol;

            IEnumerable<ViewFamilyType> ret = new FilteredElementCollector(doc)
                            .WherePasses(new ElementClassFilter(typeof(ViewFamilyType), false))
                            .Cast<ViewFamilyType>();
            ViewFamilyType FamType = ret.Where(e => e.ViewFamily == ViewFamily.FloorPlan).First() as ViewFamilyType;

            List<Level> allLevels = new List<Level>();
            foreach (Level level in new FilteredElementCollector(doc).OfClass(typeof(Level)))
            {
                allLevels.Add(level);
            }
            allLevels = allLevels.OrderBy(level => level.Elevation).ToList();

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Convert Ceiling to Floor");
                foreach (ElementId eid in ids)
                {
                    View view = doc.GetElement(eid) as View;
                    if (view.ViewType == ViewType.CeilingPlan)
                    {
                        int count = 0;
                        Level currentLevel = view.GenLevel;
                        foreach (Level sellevel in allLevels)
                        {
                            if (sellevel.Name == currentLevel.Name)
                            { break; }
                            else { count++; }
                        }
                        Level level = allLevels[count];

                        View newview = ViewPlan.Create(doc, FamType.Id, level.Id);

                        newview.Name = view.Name.ToString() + " FP";

                        ViewPlan newViewPlan = newview as ViewPlan;
                        ViewPlan oldViewPlan = view as ViewPlan;
                        newview.LookupParameter("Title on Sheet").Set(view.LookupParameter("Title on Sheet").AsString());
                        newViewPlan.SetUnderlayOrientation(UnderlayOrientation.LookingDown);
                        Level genlevel = newViewPlan.GenLevel;
                        PlanViewRange oldVR = oldViewPlan.GetViewRange();
                        PlanViewRange newVR = newViewPlan.GetViewRange();
                        newVR.SetLevelId(PlanViewPlane.TopClipPlane, oldVR.GetLevelId(PlanViewPlane.ViewDepthPlane));
                        newVR.SetLevelId(PlanViewPlane.CutPlane, oldVR.GetLevelId(PlanViewPlane.TopClipPlane));
                        newVR.SetLevelId(PlanViewPlane.BottomClipPlane, oldVR.GetLevelId(PlanViewPlane.CutPlane));
                        newVR.SetLevelId(PlanViewPlane.ViewDepthPlane, oldVR.GetLevelId(PlanViewPlane.CutPlane));
                        newVR.SetOffset(PlanViewPlane.TopClipPlane, oldVR.GetOffset(PlanViewPlane.ViewDepthPlane));
                        newVR.SetOffset(PlanViewPlane.CutPlane, oldVR.GetOffset(PlanViewPlane.TopClipPlane));
                        newVR.SetOffset(PlanViewPlane.BottomClipPlane, oldVR.GetOffset(PlanViewPlane.CutPlane));
                        newVR.SetOffset(PlanViewPlane.ViewDepthPlane, oldVR.GetOffset(PlanViewPlane.CutPlane));
                        newViewPlan.SetViewRange(newVR);
                        newview.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId());
                        newview.ViewTemplateId = view.ViewTemplateId;
                        CopyAnnot(doc, view, newview);
                    }
                }
                trans.Commit();
            }

            return Result.Succeeded;
        }
        static void CopyAnnot(Document doc, View from, View to)
        {
            ICollection<ElementId> newsel = new List<ElementId>();
            FilteredElementCollector elementsInView = new FilteredElementCollector(doc, from.Id);
            ICollection<Element> elems = elementsInView.ToElements();
            string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
            foreach (Element e in elems)
            {
                try
                {
                    string categoryName = e.Category.Name.ToString();
                    if (e.Category.CategoryType == CategoryType.Annotation && categoryName != "Grids" && categoryName != "Section Boxes" && categoryName != "Reference Planes" && categoryName != "Levels")
                    {
                        newsel.Add(e.Id);
                    }
                }
                catch { }
            }
            try
            {
                CopyPasteOptions cp = new CopyPasteOptions();
                ElementTransformUtils.CopyElements(from, newsel, to, null, cp);
            }
            catch { }
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuplicateSheets : IExternalCommand
    {
        //Replaces text in a parameter of selected elements
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            List<List<ElementId>> sheetsandview = new List<List<ElementId>>();
            FamilySymbol fs = new FilteredElementCollector(doc)
           .OfClass(typeof(FamilySymbol))
           .OfCategory(BuiltInCategory.OST_TitleBlocks)
           .FirstOrDefault() as FamilySymbol;
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Duplicate Sheets");
                foreach (ElementId eid in ids)
                {
                    ViewSheet sheet = doc.GetElement(eid) as ViewSheet;
                    ViewSheet newsheet = ViewSheet.Create(doc, fs.Id);
                    string num_suffix = "-New";
                    string name_suffix = " - Duplication";
                    if (Store.menu_A_Box.Value != null)
                    {
                        if (Store.menu_A_Box.Value.ToString() != "")
                        { num_suffix = Store.menu_A_Box.Value.ToString(); }
                    }
                    if (Store.menu_B_Box.Value != null)
                    {
                        if (Store.menu_B_Box.Value.ToString() != "")
                        {
                            newsheet.Name = sheet.Name + Store.menu_B_Box.Value.ToString();
                        }
                    }
                    //FindVert.GetMenuValue(uiapp);
                    try
                    {
                        if (Store.menu_A_Box.Value.ToString() != "*auto*")
                        {
                            newsheet.SheetNumber = sheet.SheetNumber + num_suffix;
                        }
                    }
                    catch
                    {
                        TaskDialog error = new TaskDialog("Error");
                        error.MainInstruction = "Sheet Number already exists."
                            + Environment.NewLine + "Create as: " + Environment.NewLine
                            + newsheet.SheetNumber + "-" + newsheet.Name + " ?";
                        error.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.No;
                        TaskDialogResult response = error.Show();
                        if (response == TaskDialogResult.No)
                        {
                            doc.Delete(newsheet.Id);
                            trans.Commit();
                            return Result.Succeeded;
                        }
                    }

                    // //Copy custom parameters from sheet to duplicate Specific to project

                    //Parameter para = newsheet.LookupParameter("MOD_DESCRIPTION_A");
                    //para.Set(sheet.LookupParameter("MOD_DESCRIPTION_A").AsString());

                    // //End of project specific

                    ICollection<ElementId> delete = new List<ElementId>();
                    foreach (Element e in new FilteredElementCollector(doc).OwnedByView(newsheet.Id))
                    {
                        delete.Add(e.Id);
                    }
                    doc.Delete(delete);
                    ICollection<ElementId> copy = new List<ElementId>();
                    IList<Element> ElementsOnSheet = new List<Element>();
                    foreach (Element e in new FilteredElementCollector(doc).OwnedByView(sheet.Id))
                    {
                        ElementsOnSheet.Add(e);
                    }
                    foreach (Element el in ElementsOnSheet)
                    {
                        if (el is FamilyInstance)
                        {
                            copy.Add(el.Id);
                        }
                    }
                    foreach (ElementId portid in sheet.GetAllViewports())
                    {
                        ViewDuplicateOption d_Option = ViewDuplicateOption.WithDetailing;
                        if (Store.menu_C_Box.Value != null)
                        {
                            if (Store.menu_C_Box.Value.ToString() != "")
                            {
                                d_Option = ViewDuplicateOption.AsDependent;
                            }
                            if (Store.menu_C_Box.Value.ToString() == "E")
                            {
                                d_Option = ViewDuplicateOption.Duplicate;
                            }
                        }
                        Viewport vp = doc.GetElement(portid) as Viewport;
                        List<ElementId> newlist = new List<ElementId>();
                        Element viewelem = doc.GetElement(vp.ViewId);
                        View view = viewelem as View;
                        if (view.Title.Contains("Legend"))
                        {
                            newlist.Add(view.Id);
                        }
                        else
                        { View dview = doc.GetElement(view.Duplicate(d_Option)) as View;
                            newlist.Add(dview.Id);
                        }
                        newlist.Add(vp.Id);
                        newlist.Add(newsheet.Id);
                        sheetsandview.Add(newlist);
                    }
                    View from = sheet as View;
                    View to = newsheet as View;
                    CopyPasteOptions cp = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(from, copy, to, null, cp);
                }
                trans.Commit();
            }
            using (Transaction trans2 = new Transaction(doc))
            {
                trans2.Start("place view");
                foreach (List<ElementId> list in sheetsandview)
                {
                    Viewport vp = doc.GetElement(list[1]) as Viewport;
                    Viewport newvp = Viewport.Create(doc, list[2], list[0], vp.GetBoxCenter());
                    newvp.Rotation = vp.Rotation;
                    newvp.SetBoxCenter(vp.GetBoxCenter());
                    newvp.ChangeTypeId(vp.GetTypeId());
                }
                trans2.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SheetForLevel : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            List<List<ElementId>> sheetsandview = new List<List<ElementId>>();
            FamilySymbol fs = new FilteredElementCollector(doc)
           .OfClass(typeof(FamilySymbol))
           .OfCategory(BuiltInCategory.OST_TitleBlocks)
           .FirstOrDefault() as FamilySymbol;
            List<Level> allLevels = new List<Level>();
            foreach (Level level in new FilteredElementCollector(doc).OfClass(typeof(Level)))
            {
                allLevels.Add(level);
            }
            allLevels = allLevels.OrderBy(level => level.Elevation).ToList();
            Double distance;
            if (Store.menu_1_Box.Value.ToString() == "")
            { distance = allLevels[17].Elevation - allLevels[16].Elevation; }
            else
            {
                distance = UnitUtils.ConvertToInternalUnits(Double.Parse(Store.menu_1_Box.Value.ToString()), DisplayUnitType.DUT_MILLIMETERS);
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Sheets for levels");
                foreach (ElementId eid in ids)
                {
                    ViewSheet sheet = doc.GetElement(eid) as ViewSheet;
                    ViewSheet newsheet = ViewSheet.Create(doc, fs.Id);
                    string num_suffix = "-New";
                    string name_suffix = " - Duplication";
                    if (Store.menu_A_Box.Value != null)
                    {
                        if (Store.menu_A_Box.Value.ToString() != "")
                        { num_suffix = Store.menu_A_Box.Value.ToString(); }
                    }
                    if (Store.menu_B_Box.Value != null)
                    {
                        if (Store.menu_B_Box.Value.ToString() != "")
                        {
                            newsheet.Name = sheet.Name + Store.menu_B_Box.Value.ToString();
                        }
                    }
                    try
                    {
                        if (Store.menu_A_Box.Value.ToString() != "*auto*")
                        {
                            newsheet.SheetNumber = sheet.SheetNumber + num_suffix;
                        }
                    }
                    catch
                    {
                        TaskDialog error = new TaskDialog("Error");
                        error.MainInstruction = "Sheet Number already exists."
                            + Environment.NewLine + "Create as: " + Environment.NewLine
                            + newsheet.SheetNumber + "-" + newsheet.Name + " ?";
                        error.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.No;
                        TaskDialogResult response = error.Show();
                        if (response == TaskDialogResult.No)
                        {
                            doc.Delete(newsheet.Id);
                            trans.Commit();
                            return Result.Succeeded;
                        }
                    }

                    // //Copy custom parameters from sheet to duplicate Specific to project

                    //Parameter para = newsheet.LookupParameter("MOD_DESCRIPTION_A");
                    //para.Set(sheet.LookupParameter("MOD_DESCRIPTION_A").AsString());

                    // //End of project specific

                    ICollection<ElementId> delete = new List<ElementId>();
                    foreach (Element e in new FilteredElementCollector(doc).OwnedByView(newsheet.Id))
                    {
                        delete.Add(e.Id);
                    }
                    doc.Delete(delete);
                    ICollection<ElementId> copy = new List<ElementId>();
                    IList<Element> ElementsOnSheet = new List<Element>();
                    foreach (Element e in new FilteredElementCollector(doc).OwnedByView(sheet.Id))
                    {
                        ElementsOnSheet.Add(e);
                    }
                    foreach (Element el in ElementsOnSheet)
                    {
                        if (el is FamilyInstance)
                        {
                            copy.Add(el.Id);
                        }
                    }
                    foreach (ElementId portid in sheet.GetAllViewports())
                    {
                        ViewDuplicateOption d_Option = ViewDuplicateOption.WithDetailing;
                        if (Store.menu_C_Box.Value != null)
                        {
                            if (Store.menu_C_Box.Value.ToString() != "")
                            {
                                d_Option = ViewDuplicateOption.AsDependent;
                            }
                            if (Store.menu_C_Box.Value.ToString() == "E")
                            {
                                d_Option = ViewDuplicateOption.Duplicate;
                            }
                        }
                        Viewport vp = doc.GetElement(portid) as Viewport;
                        List<ElementId> newlist = new List<ElementId>();
                        Element viewelem = doc.GetElement(vp.ViewId);
                        View view = viewelem as View;
                        if (view.Title.Contains("Legend"))
                        {
                            newlist.Add(view.Id);
                        }
                        if (view.ViewType == ViewType.Section)
                        {
                            View dview = doc.GetElement(view.Duplicate(d_Option)) as View;
                            BoundingBoxXYZ cropregion = dview.CropBox;
                            XYZ cropmin = cropregion.Min;
                            XYZ cropmax = cropregion.Max;
                            XYZ diff = new XYZ(0, distance, 0);
                            XYZ newcropmin = cropmin.Add(diff);
                            XYZ newcropmax = cropmax.Add(diff);
                            cropregion.Max = newcropmax;
                            cropregion.Min = newcropmin;
                            dview.CropBox = cropregion;
                            newlist.Add(dview.Id);
                        }
                        if (view.ViewType == ViewType.ThreeD)
                        {
                            View3D dview = doc.GetElement(view.Duplicate(d_Option)) as View3D;
                            BoundingBoxXYZ cropregion = dview.CropBox;
                            BoundingBoxXYZ sectionbox = dview.GetSectionBox();
                            XYZ cropmin = cropregion.Min;
                            XYZ cropmax = cropregion.Max;
                            XYZ diff = new XYZ(0, distance, 0);
                            XYZ newcropmin = cropmin.Add(diff);
                            XYZ newcropmax = cropmax.Add(diff);
                            cropregion.Max = newcropmax;
                            cropregion.Min = newcropmin;
                            dview.CropBox = cropregion;

                            XYZ scropmin = sectionbox.Min;
                            XYZ scropmax = sectionbox.Max;
                            XYZ sdiff = new XYZ(0, 0, distance);
                            XYZ newscropmin = scropmin.Add(sdiff);
                            XYZ newscropmax = scropmax.Add(sdiff);
                            sectionbox.Max = newscropmax;
                            sectionbox.Min = newscropmin;
                            dview.SetSectionBox(sectionbox);
                            newlist.Add(dview.Id);
                        }
                        if (view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan)
                        {
                            int count = 0;
                            Level currentLevel = view.GenLevel;
                            foreach (Level level in allLevels)
                            {
                                if (level.Name == currentLevel.Name)
                                { break; }
                                else { count++; }
                            }
                            int newLevel = 1;
                            if (Store.menu_2_Box.Value.ToString() != "")
                            {
                                newLevel = Int32.Parse(Store.menu_2_Box.Value.ToString());
                            }
                            Level nextlevel = allLevels[count + newLevel];
                            View newview = ViewPlan.Create(doc, view.GetTypeId(), nextlevel.Id);
                            newview.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId());
                            //newview.CropBox = view.CropBox;
                            newview.ViewTemplateId = view.ViewTemplateId;
                            newlist.Add(newview.Id);
                        }
                        newlist.Add(vp.Id);
                        newlist.Add(newsheet.Id);
                        sheetsandview.Add(newlist);
                    }
                    View from = sheet as View;
                    View to = newsheet as View;
                    CopyPasteOptions cp = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(from, copy, to, null, cp);
                }
                trans.Commit();
            }
            using (Transaction trans2 = new Transaction(doc))
            {
                trans2.Start("place view");
                foreach (List<ElementId> list in sheetsandview)
                {
                    Viewport vp = doc.GetElement(list[1]) as Viewport;
                    Viewport newvp = Viewport.Create(doc, list[2], list[0], vp.GetBoxCenter());
                    if (newvp.Rotation != vp.Rotation)
                    {
                        newvp.Rotation = vp.Rotation;
                        newvp.SetBoxCenter(vp.GetBoxCenter());
                    }
                    newvp.ChangeTypeId(vp.GetTypeId());
                }
                trans2.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Rotatesheet : IExternalCommand
    {
        //Rotates selected sheet clockwise
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            //FindVert.GetMenuValue(uiapp);
            List<List<ElementId>> sheetsandview = new List<List<ElementId>>();

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Rotate Sheets");
                foreach (ElementId eid in ids)
                {
                    ViewSheet sheet = doc.GetElement(eid) as ViewSheet;
                    IList<Element> ElementsOnSheet = new List<Element>();
                    foreach (Element e in new FilteredElementCollector(doc).OwnedByView(sheet.Id))
                    {
                        ElementsOnSheet.Add(e);
                    }
                    foreach (Element el in ElementsOnSheet)
                    {
                        if (el is FamilyInstance)
                        {
                            FamilyInstance famel = el as FamilyInstance;
                            XYZ point = new XYZ(0, 0, 0);
                            XYZ point1 = new XYZ(0, 0, -10);
                            Line axis = Line.CreateBound(point, point1);
                            double angle = Math.PI / 2;
                            ElementTransformUtils.RotateElement(doc, famel.Id, axis, angle);
                        }
                    }
                    foreach (ElementId portid in sheet.GetAllViewports())
                    {

                        Viewport vp = doc.GetElement(portid) as Viewport;
                        XYZ center = vp.GetBoxCenter();
                        XYZ newcenter = new XYZ(center.Y, center.X * (-1), center.Z);
                        vp.Rotation = ViewportRotation.Clockwise;
                        vp.SetBoxCenter(newcenter);

                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Rotatelegend : IExternalCommand
    {
        //Rotates selected sheet clockwise
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Rotate Legends");
                foreach (ElementId eid in ids)
                {
                    ViewSheet sheet = doc.GetElement(eid) as ViewSheet;
                    foreach (ElementId portid in sheet.GetAllViewports())
                    {
                        Viewport vp = doc.GetElement(portid) as Viewport;
                        Element viewelem = doc.GetElement(vp.ViewId);
                        View view = viewelem as View;
                        if (view.Title.Contains("Legend"))
                        {
                            vp.Rotation = ViewportRotation.Clockwise;
                        }
                        //XYZ center = vp.GetBoxCenter();
                        //XYZ newcenter = new XYZ(center.Y, center.X * (-1), center.Z);
                        //vp.Rotation = ViewportRotation.Clockwise;
                        //vp.SetBoxCenter(newcenter);

                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SwitchTeeTap : IExternalCommand
    {
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            string RouteType = "No Change";
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Switch Tap/Tee Routing preference");
                foreach (ElementId eid in ids)
                {
                    Duct duct = doc.GetElement(eid) as Duct;
                    if (duct.DuctType.RoutingPreferenceManager.PreferredJunctionType == PreferredJunctionType.Tap)
                    { duct.DuctType.RoutingPreferenceManager.PreferredJunctionType = PreferredJunctionType.Tee;
                        RouteType = "Tee";
                    }
                    else { duct.DuctType.RoutingPreferenceManager.PreferredJunctionType = PreferredJunctionType.Tap;
                        RouteType = "Tap";
                    }
                    break;
                }
                trans.Commit();
                TaskDialog.Show("Info", "Switched routing to: " + RouteType);
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class NotOnSheet : IExternalCommand
    {
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
            foreach (ElementId eid in ids)
            {
                Element view = doc.GetElement(eid) as Element;
                string SheetNum = view.LookupParameter("Sheet Number").AsString();
                if (SheetNum == "---")
                { newsel.Add(eid);
                    ViewNames += view.LookupParameter("View Name").AsString() + Environment.NewLine;
                }
            }
            ViewNames += "Press Delete after closing this window to delete.";
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select views not on sheet");
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
                TaskDialog.Show("List", ViewNames);
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RemoveProjectParameter : IExternalCommand
    {
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            FindVert.GetMenuValue(uiapp);
            string parametername = Store.menu_A_Box.Value.ToString();
            IEnumerable<ParameterElement> _params = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ParameterElement))
                    .Cast<ParameterElement>();
            List<ParameterElement> ProjectParameters = new List<ParameterElement>();
            foreach (ParameterElement pElem in _params)
            {
                if (pElem.GetDefinition().Name.Contains(parametername))
                {
                    ProjectParameters.Add(pElem);
                }
            }
            using (Transaction t = new Transaction(doc, "remove projectparameter containing -A- "))
            {
                t.Start();
                foreach (ParameterElement parameterElement in ProjectParameters)
                {
                    doc.Delete(parameterElement.Id);
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SectionsNotOnCurrentSheet : IExternalCommand
    {
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            //UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            //Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            //Document doc = uidoc.Document;
            //FindVert.GetMenuValue(uiapp);
            //string parametername = Store.menu_A_Box.Value.ToString();
            //IEnumerable<ParameterElement> _params = new FilteredElementCollector(doc)
            //        .WhereElementIsNotElementType()
            //        .OfClass(typeof(ParameterElement))
            //        .Cast<ParameterElement>();
            //List<ParameterElement> ProjectParameters = new List<ParameterElement>();
            //foreach (ParameterElement pElem in _params)
            //{
            //    if (pElem.GetDefinition().Name.Contains(parametername))
            //    {
            //        ProjectParameters.Add(pElem);
            //    }
            //}
            //using (Transaction t = new Transaction(doc, "remove projectparameter containing -A- "))
            //{
            //    t.Start();
            //    foreach (ParameterElement parameterElement in ProjectParameters)
            //    {
            //        doc.Delete(parameterElement.Id);
            //    }
            //    t.Commit();
            //}
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectAllOnLevel : IExternalCommand
    {
        //Returns elements that are referred to a link/import
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            string LevelName = "X";
            // Type in 'A' to select WITHOUT annotation instead
            if (Store.menu_A_Box.Value != null)
            {
                if (Store.menu_A_Box.Value.ToString() != "")
                { LevelName = Store.menu_A_Box.Value.ToString(); }
            }
            if (uidoc.Selection.GetElementIds().Count == 0)
            {
                FilteredElementCollector elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                ICollection<Element> elems = elementsInView.ToElements();

                string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
                foreach (Element e in elems)
                {
                    try {
                        string check = "Y";
                        if (e.LookupParameter("Level") != null) { check = e.LookupParameter("Level").AsValueString(); }
                        else if (e.LookupParameter("Reference Level") != null) { check = e.LookupParameter("Reference Level").AsValueString(); }

                        if (check == LevelName)
                            newsel.Add(e.Id);
                    }

                    catch { }
                }
            }
            else
            {
                foreach (ElementId eid in uidoc.Selection.GetElementIds())
                {
                    Element e = doc.GetElement(eid);
                    try
                    {
                        string check = "Y";
                        if (e.LookupParameter("Level") != null) { check = e.LookupParameter("Level").AsValueString(); }
                        else if (e.LookupParameter("Reference Level") != null) { check = e.LookupParameter("Reference Level").AsValueString(); }

                        if (check == LevelName)
                        {
                            newsel.Add(e.Id);
                        }
                    }
                    catch { }
                }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select all annotation in view");
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ClearOverrides : IExternalCommand
    {
        //Clears all overrides in view
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            OverrideGraphicSettings newOverride = new OverrideGraphicSettings();
            if (uidoc.Selection.GetElementIds().Count == 0)
            {
                FilteredElementCollector elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                ICollection<Element> elems = elementsInView.ToElements();
                foreach (Element e in elems)
                {
                    try
                    {
                    newsel.Add(e.Id); 
                    }
                    catch { }
                }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Remove overrides in view");
                foreach (ElementId eid in newsel)
                { doc.ActiveView.SetElementOverrides(eid, newOverride); }
                newsel = new List<ElementId>();
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class HighlightIds : IExternalCommand
    {
        //Overrides graphics of elements grouped by level on active view
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            FilteredElementCollector elementsInView = null;
            Random rnd = new Random();
            List<string> Ids = new List<string>();
            using (TextReader fileids = File.OpenText(@"C:\pROJEKT\POM\IDS.txt"))
            {
                string ID;
                while ((ID = fileids.ReadLine()) != null)
                { Ids.Add(ID); }
            }

            FillPatternElement SolidPattern = null;
            FilteredElementCollector allpatterns = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement));
            FilteredElementCollector alllevels = new FilteredElementCollector(doc).OfClass(typeof(Level));
            OverrideGraphicSettings newOverride = new OverrideGraphicSettings();
            OverrideGraphicSettings transparent = new OverrideGraphicSettings();
            foreach (FillPatternElement pattern in allpatterns)
            {
                if (pattern.GetFillPattern().IsSolidFill)
                { SolidPattern = pattern; }
            }
            newOverride.SetSurfaceBackgroundPatternId(SolidPattern.Id);
            newOverride.SetSurfaceForegroundPatternId(SolidPattern.Id);

            Byte R = (byte)rnd.Next(0, 255);
            Byte G = (byte)rnd.Next(0, 255);
            Byte B = (byte)rnd.Next(0, 255);
            newOverride.SetSurfaceBackgroundPatternColor(new Color(R, G, B));
            newOverride.SetSurfaceForegroundPatternColor(new Color(R, G, B));
            newOverride.SetProjectionLineColor(new Color(R, G, B));
            newOverride.SetSurfaceTransparency(0);
            transparent.SetSurfaceTransparency(100);
            if (uidoc.Selection.GetElementIds().Count == 0)
            {
                elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                ICollection<Element> elems = elementsInView.ToElements();
                foreach (Element e in elems)
                {
                    try
                    {
                        foreach (string id in Ids)
                        {
                            if (e.Id.ToString() == id)
                            { newsel.Add(e.Id); }
                        }
                    }
                    catch { }
                }
            }


            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Highlight IDs in view");
                foreach (Element elem in elementsInView)
                { doc.ActiveView.SetElementOverrides(elem.Id, transparent); }
                foreach (ElementId eid in newsel)
                { doc.ActiveView.SetElementOverrides(eid, newOverride); }
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
            }
            return Result.Succeeded;
        } 
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OverrideAllOnLevel : IExternalCommand
    {
        //Overrides graphics of elements grouped by level on active view
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            Random rnd = new Random();
            
            FillPatternElement SolidPattern = null;
            FilteredElementCollector allpatterns = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement));
            FilteredElementCollector alllevels = new FilteredElementCollector(doc).OfClass(typeof(Level));
            OverrideGraphicSettings newOverride = new OverrideGraphicSettings();
            foreach (FillPatternElement pattern in allpatterns)
            { if (pattern.GetFillPattern().IsSolidFill)
                {SolidPattern = pattern; }
            }
            newOverride.SetSurfaceBackgroundPatternId(SolidPattern.Id);
            newOverride.SetSurfaceForegroundPatternId(SolidPattern.Id);
            foreach (Level level in alllevels)
            {
                string LevelName = level.Name;
                Byte R = (byte)rnd.Next(0, 255);
                Byte G = (byte)rnd.Next(0, 255);
                Byte B = (byte)rnd.Next(0, 255);
                newOverride.SetSurfaceBackgroundPatternColor(new Color(R, G, B));
                newOverride.SetSurfaceForegroundPatternColor(new Color(R, G, B));
                newOverride.SetProjectionLineColor(new Color(R, G, B));

                if (uidoc.Selection.GetElementIds().Count == 0)
                {
                    FilteredElementCollector elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                    ICollection<Element> elems = elementsInView.ToElements();
                    foreach (Element e in elems)
                    {
                        try
                        {
                            string check = "Y";
                            if (e.LookupParameter("Level") != null) { check = e.LookupParameter("Level").AsValueString(); }
                            else if (e.LookupParameter("Reference Level") != null) { check = e.LookupParameter("Reference Level").AsValueString(); }

                            if (check == LevelName)
                                newsel.Add(e.Id);
                        }
                        catch { }
                    }
                }
                else
                {
                    foreach (ElementId eid in uidoc.Selection.GetElementIds())
                    {
                        Element e = doc.GetElement(eid);
                        try
                        {
                            string check = "Y";
                            if (e.LookupParameter("Level") != null) { check = e.LookupParameter("Level").AsValueString(); }
                            else if (e.LookupParameter("Reference Level") != null) { check = e.LookupParameter("Reference Level").AsValueString(); }

                            if (check == LevelName)
                            {
                                newsel.Add(e.Id);
                            }
                        }
                        catch { }
                    }
                }

                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("Override Per level in view");
                    foreach (ElementId eid in newsel)
                    {doc.ActiveView.SetElementOverrides(eid, newOverride); }
                    newsel = new List<ElementId>();
                    trans.Commit();
                }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OverrideAllByParameter : IExternalCommand
    {
        static UniqueValue checkUV(List<UniqueValue> uniqueValues, string name)
        {
            foreach (UniqueValue UV in uniqueValues)
            {
                if (UV.name == name)
                { return UV; }
            }
            return null;
        }
    static List<Byte> Rndcolor()
        {
            List<Byte> colors = new List<byte>();
            Random rnd = new Random();
            colors.Add((byte)rnd.Next(0, 255));
            colors.Add((byte)rnd.Next(0, 255));
            colors.Add((byte)rnd.Next(0, 255));
            return colors;  }
        class UniqueValue
        { public string name;
            public Byte r;
            public Byte g;
            public Byte b;
            public OverrideGraphicSettings newOverride;
            public UniqueValue(string Name, Byte R, Byte G, Byte B,OverrideGraphicSettings OVGS)
            {
                name = Name;
                r=R; g = G; b = B;
                newOverride = new OverrideGraphicSettings();
                newOverride.SetSurfaceBackgroundPatternId(OVGS.SurfaceBackgroundPatternId);
                newOverride.SetSurfaceForegroundPatternId(OVGS.SurfaceForegroundPatternId);
                newOverride.SetSurfaceBackgroundPatternColor(new Color(R, G, B));
                newOverride.SetSurfaceForegroundPatternColor(new Color(R, G, B));
                newOverride.SetProjectionLineColor(new Color(R, G, B));
            }
        }
        //Overrides graphics of elements grouped by parameter on active view
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);

            string parameterName = null;
            string Menu_B = null;
            bool ValueStringSwitch = true;
            try
            {
                parameterName = Store.menu_A_Box.Value.ToString();
                Menu_B = Store.menu_B_Box.Value.ToString(); }
            catch{}
            if (Menu_B == null || Menu_B == "") { ValueStringSwitch = false; }
            List<UniqueValue> uniqueValues = new List<UniqueValue>();
            Random rnd = new Random();
            
            FillPatternElement SolidPattern = null;
            FilteredElementCollector allpatterns = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement));
            OverrideGraphicSettings newOverride = new OverrideGraphicSettings();
            foreach (FillPatternElement pattern in allpatterns)
            {
                if (pattern.GetFillPattern().IsSolidFill)
                { SolidPattern = pattern; }
            }
            newOverride.SetSurfaceBackgroundPatternId(SolidPattern.Id);
            newOverride.SetSurfaceForegroundPatternId(SolidPattern.Id);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Override Per parameter in view");
                if (uidoc.Selection.GetElementIds().Count == 0)
                {
                    FilteredElementCollector elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                    ICollection<Element> elems = elementsInView.ToElements();
                    foreach (Element e in elems)
                    {
                       
                        try
                        {
                            if (e.LookupParameter(parameterName) != null)
                            {
                                    string value;
                                    if (ValueStringSwitch) { value = e.LookupParameter(parameterName).AsValueString(); }
                                    else { value = e.LookupParameter(parameterName).AsString();}
                                    UniqueValue checkedUV = checkUV(uniqueValues, value);
                                
                                if (checkedUV != null)
                                { 
                                    doc.ActiveView.SetElementOverrides(e.Id, checkedUV.newOverride);
                                }
                                else
                                {
                                UniqueValue newUV = new UniqueValue(value, 
                                    (byte)rnd.Next(0, 255), (byte)rnd.Next(0, 255), (byte)rnd.Next(0, 255), newOverride);
                                uniqueValues.Add(newUV);
                                doc.ActiveView.SetElementOverrides(e.Id, newUV.newOverride);
                                }
                            } 
                        }
                        catch { }
                    }
                }
                else
                {
                    foreach (ElementId eid in uidoc.Selection.GetElementIds())
                    {
                        Element e = doc.GetElement(eid);
                        try
                        {
                            
                            if (e.LookupParameter(parameterName) != null)
                            {
                                UniqueValue checkedUV = checkUV(uniqueValues, e.LookupParameter(parameterName).AsString());
                                if (checkedUV != null)
                                {
                                    doc.ActiveView.SetElementOverrides(e.Id, checkedUV.newOverride);
                                }
                                else
                                {
                                    UniqueValue newUV = new UniqueValue(e.LookupParameter(parameterName).AsString(),
                                (byte)rnd.Next(0, 255), (byte)rnd.Next(0, 255), (byte)rnd.Next(0, 255), newOverride);
                                    uniqueValues.Add(newUV);
                                    doc.ActiveView.SetElementOverrides(e.Id, newUV.newOverride);
                                }
                            }
                        }
                        catch { }
                    }
                }
                    trans.Commit();
                }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OnlyAnnotation : IExternalCommand
    {
        //Returns elements that are referred to a link/import
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            bool Annot = true;
            // Type in 'A' to select WITHOUT annotation instead
            if (Store.menu_A_Box.Value != null)
            {
                if (Store.menu_A_Box.Value.ToString() != "")
                { Annot = false; }
            }
            if (uidoc.Selection.GetElementIds().Count == 0)
            {
                FilteredElementCollector elementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
                ICollection<Element> elems = elementsInView.ToElements();

                string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
                foreach (Element e in elems)
                {
                    if (Annot)
                    {
                        try
                        {
                            if (e.Category.CategoryType == CategoryType.Annotation && e.Category.Name.ToString() != "Grids" && e.Category.Name.ToString() != "Section Boxes" && e.Category.Name.ToString() != "Reference Planes" && e.Category.Name.ToString() != "Levels")
                            {
                                newsel.Add(e.Id);
                            }
                        }
                        catch { }
                    }
                    else {
                        try
                        {
                            if (e.Category.CategoryType != CategoryType.Annotation && e.Category.Name.ToString() != "Grids")
                            {
                                newsel.Add(e.Id);
                            }
                        }
                        catch { }
                    }
                }
            }
            else
            {
                foreach (ElementId eid in uidoc.Selection.GetElementIds())
                    {
                    if (Annot)
                    {
                        try
                        {
                            Element e = doc.GetElement(eid);
                            if (e.Category.CategoryType == CategoryType.Annotation && e.Category.Name.ToString() != "Grids")
                            {
                                newsel.Add(e.Id);
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            Element e = doc.GetElement(eid);
                            if (e.Category.CategoryType != CategoryType.Annotation && e.Category.Name.ToString() != "Grids")
                            {
                                newsel.Add(e.Id);
                            }
                        }
                        catch { }
                    }
                }
            }
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("Select all annotation in view");
                    uidoc.Selection.SetElementIds(newsel);
                    trans.Commit();
                }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class IdofLinkedElement : IExternalCommand
    {
        //Returns Id of selected elements in link.
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObj = uidoc.Selection;
            string linkedid = SelectedObj.PickObject(ObjectType.LinkedElement).LinkedElementId.ToString();
            TaskDialog.Show("Id of selected object","Copied to Clipboard: "+ linkedid);
            System.Windows.Forms.Clipboard.SetText(linkedid);
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RefertoRefPlane : IExternalCommand
    {
        //Returns elements that are referred to a link/import
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
            foreach (ElementId eid in ids)
            {
                Dimension dimension = doc.GetElement(eid) as Dimension;
                foreach (Reference reference in dimension.References)
                {
                    if ((doc.GetElement(reference.ElementId) is ImportInstance) || (doc.GetElement(reference.ElementId) is RevitLinkInstance))
                    {
                        newsel.Add(eid);
                    }
                }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Move reference of Dimensionline to refplane");
                if (newsel.Count > 0) { TaskDialog.Show("x", "Reffered to Link!"); }
                uidoc.Selection.SetElementIds(newsel);

                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ReferredToImport : IExternalCommand
    {
        //Returns elements that are referred to a link/import
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            string ViewNames = "Views not on Sheet currently selected:" + Environment.NewLine;
            foreach (ElementId eid in ids)
            {
                Dimension dimension = doc.GetElement(eid) as Dimension;
                foreach (Reference reference in dimension.References)
                {
                    if ((doc.GetElement(reference.ElementId) is ImportInstance) || (doc.GetElement(reference.ElementId) is RevitLinkInstance))
                    {
                        newsel.Add(eid);
                    }
                }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select views referred to an import");
                if (newsel.Count > 0) { TaskDialog.Show("x", "Reffered to Link!"); }
                uidoc.Selection.SetElementIds(newsel);
                
                trans.Commit();
            }
            return Result.Succeeded;
            }
        }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BkFlowCheck : IExternalCommand
    {
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            FindVert.GetMenuValue(uiapp);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("BK Flow Check/Set");
                foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                FamilyInstance faminst = elem as FamilyInstance;
                MEPModel mepmod = faminst.MEPModel;
                Element connectedduct = null;
                foreach ( Connector connector in mepmod.ConnectorManager.Connectors)
                {
                    foreach (Connector connected in connector.AllRefs)
                    { connectedduct = connected.Owner; }
                    break;
                }
                string Bmcflow = elem.LookupParameter("BMC_Flow").AsValueString();
                string Realflow = connectedduct.LookupParameter("Flow").AsValueString();
                if (Bmcflow != Realflow)
                {
                    newsel.Add(eid);
                    //Type in 'A' for option to copy real flow to BK
                    //OPTION
                    if (Store.menu_A_Box.Value != null)
                    {
                            Parameter para_bmcflow = elem.LookupParameter("BMC_Flow") as Parameter;
                        para_bmcflow.Set(connectedduct.LookupParameter("Flow").AsDouble());
                    }
                    TaskDialog.Show("Wrong","BMC: "+Bmcflow + Environment.NewLine + "Real: " + Realflow); }
            }
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExplodeMEP : IExternalCommand
    {
        public List<ElementId> findconnected (List<Connector> connectors)
        {
            List<ElementId> connectedelements = new List<ElementId>();
            foreach (Connector connector in connectors)
            {
                if (connector.IsConnected)
                {
                    foreach (Connector connected in connector.AllRefs)
                    {
                        if (connector.IsConnected)
                        { connectedelements.Add(connected.Owner.Id); }
                    }
                }
            }
            return connectedelements;
        }
        public List<Connector> findconnectors (Document doc, ICollection<ElementId> ids)
        {
            
            List<Connector> outerconnectors = new List<Connector>();
            foreach (ElementId eid in ids)
            {
                ConnectorSet connectors;
                Element elem = doc.GetElement(eid);
                if (elem.Category.Name == "Ducts" || elem.Category.Name == "Pipes")
                {
                    MEPCurve mepcurve = elem as MEPCurve;
                    connectors = mepcurve.ConnectorManager.Connectors;
                }
                else
                {
                    FamilyInstance faminst = elem as FamilyInstance;
                    MEPModel mepmod = faminst.MEPModel;
                    connectors = mepmod.ConnectorManager.Connectors;
                }
                foreach (Connector connector in connectors)
                {
                    if (connector.IsConnected)
                    {
                        TaskDialog.Show("x", connector.Angle.ToString());
                        outerconnectors.Add(connector);
                    }
                }
            }
            return outerconnectors;
        }
        public Result Execute(
           ExternalCommandData commandData,
           ref string message,
           ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            
            List<Connector> outerconnectors = findconnectors(doc, ids);
            List<ElementId> connectedelements = findconnected(outerconnectors);
            List<ElementId> uniqueelements = connectedelements.Distinct().ToList();
            List<Connector> uniqueconnectors = findconnectors(doc, uniqueelements);
            TaskDialog.Show("x", connectedelements.Count.ToString() + "ALL");
            TaskDialog.Show("x", uniqueelements.Count.ToString() + "UE");
            TaskDialog.Show("x", uniqueconnectors.Count.ToString() + "UC");

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Explode MEP");
                foreach (Connector connector in uniqueconnectors)
                {
                    foreach (Connector connected in connector.AllRefs)
                    {
                        if (!ids.Contains(connector.Owner.Id))
                        { connector.DisconnectFrom(connected); }
                    }
                }
                uidoc.Selection.SetElementIds(uniqueelements);
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FindVert : IExternalCommand
    {
        //Filters out non-vertical elements from selection based on
        //Bounding box height, checks against value provided on Ribbon
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
            Double.TryParse(Store.menu_1_Box.Value as string, out Store.menu_1);
            if (Store.menu_1 == 0) { Store.menu_1 = 0.5; }
            Double.TryParse(Store.menu_2_Box.Value as string, out Store.menu_2);
            Double.TryParse(Store.menu_3_Box.Value as string, out Store.menu_3);
        }
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            GetMenuValue(uiapp);
            foreach (ElementId eid in ids)
            {
                Element fami = doc.GetElement(eid) as Element;
                BoundingBoxXYZ BB = fami.get_BoundingBox(null);
                if ((BB.Max.Z - BB.Min.Z) > Store.menu_1)
                { newsel.Add(eid); }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select vertical");
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
            }
            return Result.Succeeded;
        }
    } 
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ManageRefPlanes : IExternalCommand
    {
        UIDocument uidoc = null;
        Document doc = null;
        List<ReferencePlane> allRefs = null;
        System.Windows.Forms.Form menu = new System.Windows.Forms.Form();
        System.Windows.Forms.ComboBox selectRef = new System.Windows.Forms.ComboBox();

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            menu.MaximizeBox = false; menu.MinimizeBox = false;
            menu.MaximumSize = new System.Drawing.Size(315, 120);
            menu.MinimumSize = new System.Drawing.Size(315, 120);
            menu.Text = " Create or Delete Reference Planes";
            selectRef.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            selectRef.Location = new System.Drawing.Point(10, 10);
            selectRef.Size = new System.Drawing.Size(280, 20);
            selectRef.TabIndex = 0;
            selectRef.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            selectRef.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            System.Windows.Forms.Button delete_Button = new System.Windows.Forms.Button();
            System.Windows.Forms.Button create_Button = new System.Windows.Forms.Button();
            delete_Button.Text = " Delete "; create_Button.Text = " Create ";
            delete_Button.Location = new System.Drawing.Point(10, 40);
            create_Button.Location = new System.Drawing.Point(100, 40);
            allRefs = new List<ReferencePlane>();
            foreach (ReferencePlane refe in new FilteredElementCollector(doc).OfClass(typeof(ReferencePlane)))
            {
                if (refe.Name != "Reference Plane")
                {
                    allRefs.Add(refe);
                    selectRef.Items.Add(refe.Name);
                }
            }
            menu.Controls.Add(selectRef);
            menu.Controls.Add(create_Button); menu.Controls.Add(delete_Button);
            create_Button.Click += new EventHandler(create_Click); delete_Button.Click += new EventHandler(delete_Click);
            menu.ShowDialog(); 
        
            return Result.Succeeded;
        }
        private void delete_Click(object sender, System.EventArgs e)
        {
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Delete Ref.Plane");
                foreach (Element elem in allRefs)
                {
                    if (elem.Name == selectRef.Text)
                    {
                        doc.Delete(elem.Id); selectRef.Items.Remove(selectRef.SelectedItem);allRefs.Remove(elem as ReferencePlane); break;
                    }
                }
                trans.Commit();
            }
        }
        private void create_Click(object sender, System.EventArgs e)
        {
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<XYZ> points = new List<XYZ>();
            try
            {
                foreach (ElementId eid in ids)
                {
                    LocationPoint loc = doc.GetElement(eid).Location as LocationPoint;
                    points.Add(loc.Point);
                }
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("Create Ref.Plane");
                    XYZ bubbleEnd = points[0];
                    XYZ freeEnd = points[1];
                    XYZ thirdPnt = points[2];

                    if (!selectRef.Items.Contains(selectRef.Text))
                    {
                        selectRef.Items.Add(selectRef.Text);
                        ReferencePlane refPlane = doc.Create.NewReferencePlane2(bubbleEnd, freeEnd, thirdPnt, doc.ActiveView);
                        refPlane.Name = selectRef.Text; allRefs.Add(refPlane); trans.Commit();
                    }
                    else { TaskDialog.Show("Naming error", "Reference Plane with this name already exists."); trans.RollBack(); }
                    
                    menu.Close();
                }
            }
            catch (System.ArgumentOutOfRangeException)
            {
                TaskDialog.Show("Cannot create Ref.Plane", "Select 3 Items for creating the Ref.Plane.");
            }
            catch (System.NullReferenceException) { TaskDialog.Show("Cannot create Ref.Plane", "Select 3 Valid Items for creating the Ref.Plane."); }
        } 
    }
    //[Transaction(TransactionMode.Manual)]
    //[Regeneration(RegenerationOption.Manual)]
    //public class ConduitAngle : IExternalCommand
    //{
    //    // Adding up angles of conduits and checking if against value
    //    public Result Execute(
    //        ExternalCommandData commandData,
    //        ref string message,
    //        ElementSet elements)
    //    {
    //        UIApplication uiapp = commandData.Application;
    //        UIDocument uidoc = uiapp.ActiveUIDocument;
    //        Application app = uiapp.Application;
    //        Document doc = uidoc.Document;
    //        Selection SelectedObjs = uidoc.Selection;
    //        ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
    //        Double TotalAngle = 0;
    //        foreach (ElementId eid in ids)
    //        { Element elem = doc.GetElement(eid);
    //            if (elem.Category.Name == "Conduit Fittings")
    //            {
    //                string anglestring = elem.LookupParameter("Angle").AsValueString();
    //                string formatted = anglestring.Remove(anglestring.Length - 4, 4);
    //                Double.TryParse(formatted, out double AddAngle);
    //                TotalAngle += AddAngle;
    //            }
    //        }
    //        if (TotalAngle > 360) { TaskDialog.Show("Check Total Angle:", "Total Angle: "
    //           + TotalAngle + Environment.NewLine + "Needs Junction Box!"); }
    //        else { TaskDialog.Show("Check Total Angle:", "Total Angle: "
    //            + TotalAngle + Environment.NewLine + "Fine!"); }
    //        return Result.Succeeded;
    //    }
    //}
}