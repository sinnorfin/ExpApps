/**
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
//using Autodesk.Revit.DB.Mechanical;

namespace MultiDWG
{
    public static class Store
        //For storing values that are updated from the Ribbon
    {
        public static Double mod_stepy = 1;
        public static TextBox stepy_ib = null;
        public static Double mod_left = 1;
        public static TextBox left_ib = null;
        public static Double mod_right = 1;
        public static TextBox right_ib = null;
        public static Double mod_firsty = 1;
        public static TextBox firsty_ib = null;
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuctId : IExternalCommand
    {
        private string Correct(string system)
            //Searches String for varies typos, corrects to preset standard
        {
            string abr = null;
            if (system.Contains("EXH")) { abr = "EA"; }
            else if (system.Contains("RET")) { abr = "RA"; }
            else if (system.Contains("SUP")) { abr = "SA"; }
            else if (system.Contains("FLUE")) { abr = "F"; }
            else if (system.Contains("CAI")) { abr = "CAI"; }
            else { abr = system; }
            return abr;
        }
        private List<List<Element>> SortParam_String(List<List<Element>> toSort, string param)
        {
            //Sorts by String
            List<List<Element>> Sorted = new List<List<Element>>();
            foreach (List<Element> sortable in toSort)
            {
                List<List<Element>> Sortation = new List<List<Element>>();
                Sortation = sortable.GroupBy(s => s.LookupParameter(param).AsString()).Select(x => x.ToList()).ToList();
                Sorted.AddRange(Sortation);
            }
            return Sorted;
        }
        private List<List<Element>> SortParam_VString(List<List<Element>> toSort, string param)
        {
            //Sorts by ValueString
            //Should be merged to the above
            //Should be standardized sortation mechanism
            List<List<Element>> Sorted = new List<List<Element>>();
            foreach (List<Element> sortable in toSort)
            {
                List<List<Element>> Sortation = new List<List<Element>>();
                Sortation = sortable.GroupBy(s => s.LookupParameter(param).AsValueString()).Select(x => x.ToList()).ToList();
                Sorted.AddRange(Sortation);
            }
            return Sorted;
        }
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
            List<Element> Ducts = new List<Element>();
            List<List<Element>> Systems = new List<List<Element>>();
            // Could be standardised, replace variables with settings from inside the program
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                //FamilyInstance fami = elem as FamilyInstance;
                //MechanicalFitting mep = fami.MEPModel as MechanicalFitting;
                //if (mep.PartType == PartType.Elbow)
                //{  }
                if (elem.Category.Name == "Ducts") // Change elem.Category.Name = "" as Category requires
                {
                    Ducts.Add(elem);
                }
            }
            Systems = Ducts.GroupBy(s => s.LookupParameter("System Abbreviation").AsString()).Select(x => x.ToList()).ToList();
            List<List<Element>> Family = SortParam_String(Systems, "Family Name");
            List<List<Element>> Sizes = SortParam_String(Family,"Size");
            List<List<Element>> Length = SortParam_VString(Sizes, "Length");
            // Add sortation as needed, refer previous, replace parameter name as needed
            //List<List<Element>> Off1 = SortParam_VString(Length, "Actual Transition Offset 1");
            //List<List<Element>> Off2 = SortParam_VString(Off1, "Actual Transition Offset 2");
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Duct ID");
                Int32 count = -1; // set to start ID number -1
                List<string> SortSys = new List<string>();
                foreach (List<Element> listElem in Length)  // replace 'in ...' with last sorted list
                {
                    if (!SortSys.Contains(Correct(listElem[0].LookupParameter("System Abbreviation").AsString())))
                    {
                        SortSys.Add(Correct(listElem[0].LookupParameter("System Abbreviation").AsString()));
                        count = 0;
                    }
                    else
                    {
                        count += 1;
                    }
                    foreach (Element elem in listElem)
                    {
                        string abr = Correct(elem.LookupParameter("System Abbreviation").AsString());
                        elem.LookupParameter("Manual ID").Set(abr + "-" + count.ToString("000")); // Change format accordingly
                    }
                }
                trans.Commit();
            }
            TaskDialog.Show("Created Manual IDs", "Sys:" + Systems.Count.ToString() + " Size:" + Sizes.Count.ToString() + " Len:" + Length.Count.ToString());
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
    public class PlaceAccessory : IExternalCommand
    {
        // Not working, Accessory needs to be aligned
        // pipe has to manually recreated and connected back on both ends
        // Pipe has to connect to previous connector + new connector on fitting
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
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector.OfClass(typeof(FamilySymbol))
                                                       .OfCategory(BuiltInCategory.OST_PipeAccessory)
                                                       .ToElements();
            IEnumerator<Element> symbolItor = collection.GetEnumerator();
            symbolItor.MoveNext();
            using (Transaction tx = new Transaction(doc))
            {
                FamilyInstance instance = null;
                LocationCurve pipeloccurve = null;
                tx.Start("Place Accesory on pipe");
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    if (elem.Category.Name == "Pipes")
                    {
                        MEPCurve pipecurve = elem as MEPCurve;
                            pipeloccurve = elem.Location as LocationCurve;
                            Level level = pipecurve.ReferenceLevel as Level;
                            FamilySymbol symbol = symbolItor.Current as FamilySymbol;

                            XYZ location = pipeloccurve.Curve.Evaluate(0.5, true);
                            if (location == null) { TaskDialog.Show("error", "location null"); }
                            if (symbol == null) { TaskDialog.Show("error", "symbol null"); }
                            instance = doc.Create.NewFamilyInstance(location, symbol, level, level, StructuralType.NonStructural);
                            
                    }
                }
                tx.Commit();
                XYZ ac_refline = XYZ.Zero;
                Reference instaref = instance.GetReferences(FamilyInstanceReferenceType.CenterLeftRight).FirstOrDefault();
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Temporary Sketch Plane");
                    SketchPlane sk = SketchPlane.Create(doc, instaref);
                    if (null != sk)
                    {
                        Plane pl = sk.GetPlane();
                        ac_refline = pl.Normal;
                    }
                    t.RollBack();
                }
                Line pi_refline = pipeloccurve.Curve as Line;
                if (ac_refline != null) { TaskDialog.Show("Accessory Direction:", ac_refline.ToString()); }
                TaskDialog.Show("Pipe Direction:", pi_refline.Direction.ToString());
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
                { Double AddAngle = 0;
                    string anglestring = elem.LookupParameter("Angle").AsValueString();
                    string formatted = anglestring.Remove(anglestring.Length-4, 4);
                    Double.TryParse(formatted,out AddAngle);
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
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Calculate Duct Surface Area");
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    FamilyInstance fami = elem as FamilyInstance;
                    MEPModel mepModel = fami.MEPModel;
                    Double connectorAreas = 0; 
                    foreach (Connector connector in mepModel.ConnectorManager.Connectors)
                    {
                        try { 
                            connectorAreas += Math.Pow(connector.Radius, 2) * Math.PI;
                        }
                        catch ( Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            connectorAreas += connector.Width * connector.Height;
                        }
                    }
                    GeometryElement geoelem = elem.get_Geometry(geOpt);
                    foreach (GeometryInstance geoInst in geoelem)
                    {
                        foreach (Solid solid in geoInst.GetInstanceGeometry())
                        {
                            if (solid.SurfaceArea > 0)
                            {
                                elem.LookupParameter("Duct Surface Area").Set(solid.SurfaceArea-connectorAreas);
                            }
                        }
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
        //Loads multiple DWG at once
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
                        DWGImportOptions ImportOptions = new DWGImportOptions();
                        ImportOptions.Placement = ImportPlacement.Origin;
                        ImportOptions.ColorMode = ImportColorMode.BlackAndWhite;
                        string filepath = dlg.FileName;
                        filepath = filepath.Remove(dlg.FileName.LastIndexOf(@"\"));
                        Array files = System.IO.Directory.GetFiles(filepath);
                        foreach (string e in files)
                        {
                            ElementId elementid = null;
                            doc.Import(e, ImportOptions, doc.ActiveView, out elementid);
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
                                doc.ActiveView.HideElementTemporary(elementid);
                            }
                            else if (e.Contains("fine") || (e.Contains("FINE")))
                            {
                                LODvis = LODvis & ~(1 << 13);
                                LODvis = LODvis & ~(1 << 14);
                                doc.ActiveView.HideElementTemporary(elementid);
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
    public class SwapMultiDWG : IExternalCommand
    {
        // for swapping existing dwg-s with loading
        // probably decrepit
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            //ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            //Options geOpt = new Options();
            //Element style = doc.GetElement(ids.FirstOrDefault());
            //TaskDialog.Show("asd", style.Name);
            //ElementId grap = style.get_Geometry(geOpt).GraphicsStyleId;
            using (Transaction tx = new Transaction(doc))
            {
                List<Element> imports = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).ToList();
                string importName = doc.PathName.Remove(doc.PathName.Length-doc.Title.Length-16)+"DWGs\\";
                //TaskDialog.Show("PATH", importName);
                tx.Start("Swap Multiple DWGs");
                foreach (Element elem in imports)
                {
                //    GeometryElement impi = elem.get_Geometry(geOpt);
                //    foreach (GeometryInstance geoInst in impi)
                //    {
                //        Curve curve = null;
                //        foreach (var item in geoInst.GetSymbolGeometry())
                //        {
                //            curve = item as Curve;
                //            if (curve != null)
                //            { curve.SetGraphicsStyleId(grap);
                //                //TaskDialog.Show("A", "A");
                //            }
                //        }
                //    }
                //    //{ curve.SetGraphicsStyleId(grap); }                  
                  doc.Delete(elem.Id);
                }
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = importName;
                    DWGImportOptions ImportOptions = new DWGImportOptions();
                    ImportOptions.Placement = ImportPlacement.Origin;
                    ImportOptions.ColorMode = ImportColorMode.BlackAndWhite;
                    string filepath = dlg.FileName;
                    filepath = filepath.Remove(dlg.FileName.LastIndexOf(@"\"));
                    Array files = System.IO.Directory.GetFiles(filepath);
                    foreach (string e in files)
                    {
                        ElementId elementid = null;
                        doc.Import(e, ImportOptions, doc.ActiveView, out elementid);
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
                            doc.ActiveView.HideElementTemporary(elementid);
                        }
                        else if (e.Contains("fine") || (e.Contains("FINE")))
                        {
                            LODvis = LODvis & ~(1 << 13);
                            LODvis = LODvis & ~(1 << 14);
                            doc.ActiveView.HideElementTemporary(elementid);
                        }
                        LOD.Set(LODvis);
                    }
                tx.Commit();
            }
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
                foreach (ElementId eid in ids)
                    {
                        Element elem = doc.GetElement(eid) as Element;
                        Parameter para = elem.LookupParameter(Store.left_ib.Value.ToString()) as Parameter;
                        string original = para.AsString();
                        if (original != null)
                        {
                            para.Set(original.Replace(Store.right_ib.Value.ToString(), Store.firsty_ib.Value.ToString()));
                            if (para.AsString() != original) { c += 1; newsel.Add(eid); }
                        }
                    }
                trans.Commit();
                uidoc.Selection.SetElementIds(newsel);
                TaskDialog.Show("Result", "Replaced: " + Store.right_ib.Value.ToString() + " to: " + Store.firsty_ib.Value.ToString() + " in: " + c.ToString() + " elements");
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
                if (item.Name == "Step Y")
                { Store.stepy_ib = (TextBox)item; }
                if (item.Name == "Left Space")
                { Store.left_ib = (TextBox)item; }
                if (item.Name == "Right Space")
                { Store.right_ib = (TextBox)item; }
                if (item.Name == "First Y")
                { Store.firsty_ib = (TextBox)item; }
            }
            Double.TryParse(Store.stepy_ib.Value as string, out Store.mod_stepy);
            if (Store.mod_stepy == 0) { Store.mod_stepy = 0.5; }
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
                if ((BB.Max.Z - BB.Min.Z) > Store.mod_stepy)
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
    public class MetaData : IExternalCommand
    {
        // Simple parameter filler
        // No real use for this anymore
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
            foreach (ElementId eid in ids)
            {
                FamilyInstance fami = doc.GetElement(eid) as FamilyInstance;
                string keynote = "";
                string model = "";
                string manufacturer = "";
                string url = "";
                string description = "";
                string uniformat = "";
                Family fam = fami.Symbol.Family;
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("Set Parameters");
                    foreach (ElementId famsym in fam.GetFamilySymbolIds())
                    {
                        FamilySymbol famtype = doc.GetElement(famsym) as FamilySymbol;
                        famtype.get_Parameter(BuiltInParameter.KEYNOTE_PARAM).Set(keynote);
                        famtype.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).Set(model);
                        famtype.get_Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).Set(manufacturer);
                        famtype.get_Parameter(BuiltInParameter.ALL_MODEL_URL).Set(url);
                        famtype.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).Set(description);
                        famtype.get_Parameter(BuiltInParameter.UNIFORMAT_CODE).Set(uniformat);
                    }
                    trans.Commit();
                }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TypeParam : IExternalCommand
    {
        //Converts all type parameters to instance parameters
        //Should be tested, might be useful
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            FamilyManager famman = doc.FamilyManager;
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Set to Type");
                foreach (FamilyParameter fampar in famman.Parameters)
                {
                    if (fampar.Definition.ParameterGroup == BuiltInParameterGroup.PG_IDENTITY_DATA)
                    {
                       famman.Set(fampar,""); 
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeleteDetailItem : IExternalCommand
    {
        // No use for this anymore
        public Result Execute(
              ExternalCommandData commandData,
              ref string message,
              ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Array load_dirs = System.IO.Directory.GetDirectories("O:\\2_Project_Data\\22_Model\\223_Families\\01_Wall Systems");
            List<string> arch_load = Mullion.GetAllDir(load_dirs, "Architectural");
            List<string> fl_arch_load = Mullion.Searchrvt(arch_load, "v02.rvt");
            TextWriter tw_arch_load = new StreamWriter("E:\\ARCH_DDI.txt");
            IFamilyLoadOptions famoption = new Mullion.OverwriteFamilyLoadOptions();
            foreach (String s in fl_arch_load)
                tw_arch_load.WriteLine(s);
            tw_arch_load.Close();
            int count = 0;
            foreach (string path in fl_arch_load)
            {
                //if (count > -1 && count < 3) //|| (count == 6) || (count == 8) || (count == 12) || (count == 56) || (count == 66)) // PROJECT
              //  {
                    uiapp.OpenAndActivateDocument(path);
                    doc.Close();
                    doc = uiapp.ActiveUIDocument.Document;
                    List<ElementId> prof_ids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProfileFamilies).ToElementIds().ToList();
                //TaskDialog.Show("sd", prof_ids.Count().ToString());
                foreach (ElementId profile in prof_ids)
                {
                    //TaskDialog.Show("asd", profile.Name);
                    FamilySymbol faminst = doc.GetElement(profile) as FamilySymbol;
                    Family fam = faminst.Family;
                    if (!fam.Name.Contains("OSC"))
                    {
                        Document editFamily = doc.EditFamily(fam);
                        using (Transaction trans = new Transaction(editFamily))
                        {
                            trans.Start("Remove Detail items");
                            List<ElementId> detailItems = new FilteredElementCollector(editFamily).OfCategory(BuiltInCategory.OST_DetailComponents).ToElementIds().ToList();
                            //TaskDialog.Show("asd", detailItems.Count().ToString());// != 0)
                            {
                                foreach (ElementId detailId in detailItems)
                                {
                                    try
                                    {
                                        editFamily.Delete(detailId);
                                    }
                                    catch { }

                                }
                            }
                            trans.Commit();
                        }
                        editFamily.LoadFamily(doc, famoption);
                    }
                }
                    count += 1;
                //}
               // else { count += 1; }
            }
            TaskDialog.Show("Finish", "Finished");
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class EngToArch : IExternalCommand
    {
        //No use for this now
        public static List<ElementId> ElementsNamed(Document doc, List<ElementId> elementids, string name)
        {
            List<ElementId> result = new List<ElementId>();
            foreach (ElementId eid in elementids)
            {
                if (doc.GetElement(eid).Name.Contains(name))
                {
                    result.Add(eid);
                }
            }
            return result;
        }
        public Result Execute(
              ExternalCommandData commandData,
              ref string message,
              ElementSet elements)
        {
            //No USE for this
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Array load_dirs = System.IO.Directory.GetDirectories("O:\\2_Project_Data\\22_Model\\223_Families\\01_Wall Systems");
            List<string> eng_load = Mullion.GetAllDir(load_dirs,"Engineering");
            List<string> arch_load = Mullion.GetAllDir(load_dirs,"Architectural");
            List<string> fl_eng_load = Mullion.Searchrvt(eng_load, "v02.rvt");
            List<string> fl_arch_load = Mullion.Searchrvt(arch_load, ".rvt");
            TextWriter tw_eng_load = new StreamWriter("E:\\ENG.txt");
            foreach (String s in fl_eng_load)
                tw_eng_load.WriteLine(s);
            tw_eng_load.Close();
            TextWriter tw_arch_load = new StreamWriter("E:\\ARCH.txt");
            foreach (String s in fl_arch_load)
                tw_arch_load.WriteLine(s);
            tw_arch_load.Close();
            int count = 0;
            string inpath = null;
            foreach (string path in fl_eng_load)
            {
                if (count > 48 && count < 50) //|| (count == 6) || (count == 8) || (count == 12) || (count == 56) || (count == 66)) // PROJECT
                {
                    uiapp.OpenAndActivateDocument(path);
                    doc.Close();
                    doc = uiapp.ActiveUIDocument.Document;
                    
                    foreach (string insertpath in fl_arch_load)
                    {
                    string pathname = insertpath.Remove(0, insertpath.LastIndexOf('\\')+1);
                    string name = pathname.Remove(pathname.Length - 4, 4);
                        //TaskDialog.Show("asd", name + " VS " + doc.Title.Remove(doc.Title.Length - 4, 4));
                    // if (insertpath.Contains(doc.Title.Remove(doc.Title.Length - 4, 4)))
                    if (name == doc.Title.Remove(doc.Title.Length-4,4))
                        {
                        inpath = insertpath;
                        }
                    }
                    List<ElementId> prof_ids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProfileFamilies).ToElementIds().ToList();
                    List<ElementId> mull_ids = new List<ElementId>();
                    List<Element> mulls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CurtainWallMullions).Where(q => q.Location == null).ToList();
                    foreach (Element elem in mulls)
                    { mull_ids.Add(elem.Id); }
                    List<ElementId> wall_ids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).ToElementIds().ToList();
                    List<ElementId> tocopy = ElementsNamed(doc, prof_ids, "XXX");
                    
                    tocopy.AddRange(ElementsNamed(doc, mull_ids, "XXX"));
                    //foreach ( ElementId wall in wall_ids) {if (doc.GetElement(wall).get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() == "Curtain Wall" && doc.GetElement(wall).Location == null) { tocopy.Add(wall); } }

                    Document targetDoc = app.OpenDocumentFile(inpath);
                    using (Transaction trans = new Transaction(targetDoc))
                    {
                        trans.Start("Copy Families");
                        ElementTransformUtils.CopyElements(doc, tocopy, targetDoc, null, new CopyPasteOptions());
                        trans.Commit();
                        targetDoc.SaveAs(inpath.Remove(inpath.Length-4,4)+"_v02.rvt");
                    }

                    // // COPY CURTAIN WALLS CONTAINING 'XXX' // // 

                    //List<Element> CWcopy = new List<Element>();
                    //List<Element> curtainwalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).ToElements().ToList();
                    //TaskDialog.Show("ad", curtainwalls.Count().ToString());
                    //foreach (Element elem in curtainwalls)
                    //{
                    //    try
                    //    {
                    //        Wall wall = elem as Wall;
                    //        ICollection<ElementId> mullions = wall.CurtainGrid.GetMullionIds();
                    //        foreach (ElementId eid in mullions)
                    //        {
                    //            if (doc.GetElement(eid).Name.Contains("OSC-L"))
                    //            {
                    //                CWcopy.Add(elem);
                    //                TaskDialog.Show("FOUND", "Found");
                    //                break;
                    //            }
                    //        }
                    //    }
                    //    catch { }
                    //}
                    //TaskDialog.Show("asd", CWcopy.Count().ToString());

                    // // END OF SECTION // //

                    count += 1;
                }
                else { count += 1; }
            }              
                //IFamilyLoadOptions famoption = new OverwriteFamilyLoadOptions();
                //ReloadFam(doc, fl_insert, famoption);
                return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Mullion : IExternalCommand
    {
        // No use for this
        public void Swap(Document doc, List<string> InsertPaths)
        {
            bool s = true;
            List<string> swapped = new List<string>();
            FilteredElementCollector profiles = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProfileFamilies);
            FilteredElementCollector mullions = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CurtainWallMullions);
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Load Circle");
                Family circle = null;
                FamilySymbol cirtyp = null;
                doc.LoadFamily("E:\\Family1.rfa", out circle);
                foreach (ElementId eid in circle.GetFamilySymbolIds())
                {
                    cirtyp = doc.GetElement(eid) as FamilySymbol;
                    break;
                }
                trans.Commit();

                foreach (Element pr in profiles)
                {
                    s = true;
                    string name = pr.Name;
                    foreach (string path in InsertPaths)
                    {
                        if (s == false) { break; }

                        string type = path.Remove(0, path.LastIndexOf("\\") + 1);
                        string typename = type.Remove(type.Length - 4, 4);
                        if (path.Contains(name) && swapped.Contains(name) != true)
                        {
                            //TaskDialog.Show("asd", "Megvan");
                            //if (doc.PathName.Contains("Architectural") && path.Contains("Engineering")) { }
                            //if (doc.PathName.Contains("Engineering") && path.Contains("Architectural")) { }
                            //else
                            //{
                            swapped.Add(name);
                            foreach (Element mu in mullions)
                            {
                                if (s == true)
                                {
                                    try
                                    {
                                        FamilyInstance famins = mu as FamilyInstance;
                                        if (s == true)
                                        {
                                            foreach (ElementId famsym in famins.Symbol.Family.GetFamilySymbolIds())
                                            {
                                                FamilySymbol famtype = doc.GetElement(famsym) as FamilySymbol;
                                                if (famtype.get_Parameter(BuiltInParameter.MULLION_PROFILE).AsValueString() == typename + " : " + typename)
                                                {
                                                    trans.Start("Swap family");
                                                    famtype.get_Parameter(BuiltInParameter.MULLION_PROFILE).Set(cirtyp.Id);
                                                    FamilySymbol prsym = pr as FamilySymbol;
                                                    //TaskDialog.Show("asd", "Cseretipus törlésébe fagyok");
                                                    doc.Delete(prsym.Family.Id);
                                                    //TaskDialog.Show("asd", "Toltesbe fagyok");
                                                    doc.LoadFamily(path);
                                                    //TaskDialog.Show("asd", "Toltve");
                                                    profiles = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProfileFamilies);
                                                    foreach (Element prof in profiles)
                                                    {
                                                        if (prof.Name == typename)
                                                        {
                                                            //TaskDialog.Show("asd", "Cserébe fagyok");
                                                            famtype.get_Parameter(BuiltInParameter.MULLION_PROFILE).Set(prof.Id);
                                                            //TaskDialog.Show("asd", "Csere" + Environment.NewLine + "PATH:" + path + Environment.NewLine
                                                            //  + "PROFILE NAME :" + prof.Name);
                                                            s = false; trans.Commit(); break;
                                                            //  }
                                                        }
                                                    }
                                                    trans.Commit();
                                                }
                                            }
                                        }
                                        else { break; }
                                    }
                                    catch
                                    { //TaskDialog.Show("asd", "HIBA"); 
                                    }
                                }
                                else { break; }
                                //}
                            }
                        }
                    }
                }
                trans.Start("Delete Circle");
                doc.Delete(circle.Id);
                trans.Commit();
            }
        }
        public void ReloadFam(Document doc, List<string> InsertPaths,IFamilyLoadOptions loadOptions)
        {
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Reload");
                int count = 0;
                foreach (string path in InsertPaths)
                {
                    if (count == 0 || count == 2 || count == 4 || count == 6 || count == 8 || count == 10 || count == 12 || count == 14) // FAMILY
                    {
                        doc.LoadFamily(path,loadOptions,out Family fam);
                        count += 1;
                    }
                    else count += 1;
                }
                trans.Commit();
            }
        }
        public static List<string> GetAllDir(Array Idir, string dirname)
        {
            // Should test, extracts all folders under one folder
            List<string> engdir = new List<string>();
            foreach (string dir in Idir)
            {
                if (dir.Contains(dirname))
                { engdir.Add(dir); }
                else { engdir.AddRange(GetAllDir(System.IO.Directory.GetDirectories(dir),dirname)); }
            }
            if (engdir != null) return engdir;
            else return new List<string>();
        }
        public List<string> EndFolders(Array Idir)
        {
            // Returns only the deepest folders of folder structure
            List<string> enddir = new List<string>();
            foreach (string dir in Idir)
            {
                if (System.IO.Directory.GetDirectories(dir).Length == 0 && dir.Contains("Engineering"))
                { enddir.Add(dir); }
                else { enddir.AddRange(EndFolders(System.IO.Directory.GetDirectories(dir))); }
            }
            if (enddir != null) return enddir;
            else return new List<string>();
        }
        public class OverwriteFamilyLoadOptions : IFamilyLoadOptions
            //NO use for this
        {
            #region IFamilyLoadOptions Members

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
 
                    overwriteParameterValues = true;
              
                    return true;
                
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                    source = FamilySource.Family;
                    overwriteParameterValues = true;
                    return true;
            }
            #endregion
        }
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Array load_dirs = System.IO.Directory.GetDirectories("O:\\2_Project_Data\\22_Model\\223_Families\\01_Wall Systems");
            Array insert_dirs = System.IO.Directory.GetDirectories("O:\\3_Deliverables\\02_Entrances");
            List<string> load = GetAllDir(load_dirs,"Engineering");
            List<string> fl_load = Searchrvt(load,"v02.rvt");
            List<string> insert = EndFolders(insert_dirs);
            List<string> fl_insert = Searchrvt(insert, ".rfa");
            TextWriter tw_load = new StreamWriter("E:\\SavedLoad.txt");
            foreach (String s in fl_load)
                tw_load.WriteLine(s);
            tw_load.Close();
            TextWriter tw_insert = new StreamWriter("E:\\SavedInsert.txt");
            foreach (String s in fl_insert)
                tw_insert.WriteLine(s);
            tw_insert.Close();
            //return Result.Succeeded;
            int count = 0;
            foreach (string path in fl_load)
            {
                if ( count == 29 || count == 31 || count == 33 || count == 35) //|| (count == 6) || (count == 8) || (count == 12) || (count == 56) || (count == 66)) // PROJECT
                {
                    uiapp.OpenAndActivateDocument(path);
                    doc = uiapp.ActiveUIDocument.Document;
                    //Swap(doc, OscPaths);
                    IFamilyLoadOptions famoption = new OverwriteFamilyLoadOptions();
                    ReloadFam(doc, fl_insert,famoption);
                    count += 1;
                }
                else
                { count += 1; }
            }
            TaskDialog.Show("Finish","Files were reloaded");
            return Result.Succeeded;
        }
        public static List<string> Searchrvt(List<string> dirs,string fileend)
        {
            // no use
            List<string> files = new List<string>();
            if (dirs != null)
            {
                foreach (string dirpath in dirs)
                {
                    foreach (string path in System.IO.Directory.GetFiles(dirpath))
                    {
                        if (path.EndsWith(fileend) && path.Contains(".00") != true)
                        { files.Add(path); }
                    }
                }
            }
            return files;
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