﻿/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018,2019 by Attila Kalina <attilakalina.arch@gmail.com>
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
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using _ExpApps;
//using Autodesk.Revit.DB.Mechanical;

namespace MultiDWG
{
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
    public class ConduitAngle : IExternalCommand
    {
        // Adding up angles of conduits and checking if against value
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Double TotalAngle = 0;
            foreach (ElementId eid in ids)
            { Element elem = doc.GetElement(eid);
                if (elem.Category.Name == "Conduit Fittings")
                {
                    string anglestring = elem.LookupParameter("Angle").AsValueString();
                    string formatted = anglestring.Remove(anglestring.Length-4, 4);
                    Double.TryParse(formatted,out double AddAngle);
                    TotalAngle += AddAngle;
                    }
            } 
            if ( TotalAngle > 360) { TaskDialog.Show("Check Total Angle:", "Total Angle: "
                + TotalAngle + Environment.NewLine + "Needs Junction Box!"); }
            else { TaskDialog.Show("Check Total Angle:", "Total Angle: " 
                + TotalAngle + Environment.NewLine + "Fine!"); }
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
                        catch ( Autodesk.Revit.Exceptions.InvalidOperationException)
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
                    double total = surfaceAreas - connectorAreas *  parts ;
                    elem.LookupParameter("Duct Surface Area").Set(total);
                    // CHECK NUMBER OF PARTS
                    if (Store.menu_A_Box.Value.ToString() != "") { TaskDialog.Show("Report", "Parts counted: " + parts); }
                    // TO ONLY INCLUDE CONNECTOR SIZE ON ENDCAPS
                    if (ConnectorCount == 1 ) { elem.LookupParameter("Duct Surface Area").Set(connectorAreas); }
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
                    catch {}
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
                            { para.Set(original + Store.menu_C_Box.Value.ToString()); }
                            else if (Store.menu_B_Box.Value.ToString() == "add*")
                            { para.Set(Store.menu_C_Box.Value.ToString() + original); }

                            {para.Set(original.Replace(Store.menu_B_Box.Value.ToString(), Store.menu_C_Box.Value.ToString()));}
                            if (para.AsString() != original)
                            {
                                c += 1; newsel.Add(eid);
                            }
                        }
                        catch { x += 1; }
                    }
                trans.Commit();
                uidoc.Selection.SetElementIds(newsel);
                string text = "Replaced '" +  Store.menu_B_Box.Value.ToString() 
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
                        if (Store.menu_C_Box.Value.ToString() != "")
                        {
                            double orig;
                            double oper;
                            Double.TryParse(sourcePara.AsValueString(), out orig);
                            Double.TryParse(Store.menu_C_Box.Value.ToString(), out oper);
                            double sum = orig + oper;
                            targetPara.Set(sum.ToString());

                        }
                        else
                        {
                            targetPara.Set(sourcePara.AsValueString());
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
                double HostLevelElevation = 0;
                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                ICollection<Element> lvlCollection = lvlCollector.OfClass(typeof(Level)).ToElements();

                if (Store.menu_1_Box.Value.ToString() != "")
                {
                    
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
                    String TargetParaString = "Comments";
                    Double Resultelevation;
                    Element elem = doc.GetElement(eid);
                    //Get host Level's elevation
                    if (HostLevel == null)
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
                    Resultelevation = ElemCenterElevation - HostLevelElevation ;
                    Resultelevation = UnitUtils.ConvertFromInternalUnits(Resultelevation, DisplayUnitType.DUT_MILLIMETERS);
                    Resultelevation = Math.Round(Resultelevation, 0);
                    string familyName = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                    if (Store.menu_3_Box.Value.ToString() != "")
                    {
                        TargetParaString = Store.menu_3_Box.Value.ToString();
                    }
                    try
                    {
                        Parameter para_inject;
                        para_inject = elem.LookupParameter(TargetParaString) as Parameter;
                   
                        if (familyName.Contains("Circ"))
                        {
                            para_inject.Set(Resultelevation.ToString());
                            c += 1;
                        }
                        if (familyName.Contains("Rect"))
                        {
                            double recess_height = UnitUtils.ConvertFromInternalUnits(elem.LookupParameter("Recess Height").AsDouble(), DisplayUnitType.DUT_MILLIMETERS);
                            recess_height = Math.Round(recess_height, 0);
                            Resultelevation -= (recess_height / 2);
                            para_inject.Set(Resultelevation.ToString());
                            r += 1;
                        }
                    }
                    catch { x += 1; newsel.Add(elem.Id); }
                }
                trans.Commit();
                uidoc.Selection.SetElementIds(newsel);
                string text = "Circular Recess updated: " + c.ToString()
                              + Environment.NewLine +  "Rectangular Recess updated: " + r.ToString()
                              + Environment.NewLine + "Invalid Elements (selected) :" + x.ToString();
                TaskDialog.Show("Result", text);
            }
            return Result.Succeeded;
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
                    ElementTransformUtils.CopyElements(from,copy,to,null,cp);
                }
                trans.Commit();
            }
            using (Transaction trans2 = new Transaction(doc))
            {
                trans2.Start("place view");
                    foreach (List<ElementId> list in sheetsandview)
                    {
                    Viewport vp = doc.GetElement(list[1]) as Viewport;
                    Viewport newvp = Viewport.Create(doc,list[2],list[0], vp.GetBoxCenter());
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
            //string ViewNames = "ID - different flow:" + Environment.NewLine;
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
                string Reaflow = connectedduct.LookupParameter("Flow").AsValueString();
                if (Bmcflow != Reaflow) { newsel.Add(eid);
                TaskDialog.Show("Wrong","BMC: "+Bmcflow + Environment.NewLine + "Real: " + Reaflow); }
            }
            
            //ViewNames += "Press Delete after closing this window to delete.";
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select Different");
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
                //TaskDialog.Show("List", ViewNames);
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
                if (panel.Name == "Annotation")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "1")
                { Store.menu_1_Box = (TextBox)item; }
                if (item.Name == "2")
                { Store.menu_2_Box = (TextBox)item; }
                if (item.Name == "3")
                { Store.menu_3_Box = (TextBox)item; }
                if (item.Name == "A")
                { Store.menu_A_Box = (TextBox)item; }
                if (item.Name == "B")
                { Store.menu_B_Box = (TextBox)item; }
                if (item.Name == "C")
                { Store.menu_C_Box = (TextBox)item; }
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
}