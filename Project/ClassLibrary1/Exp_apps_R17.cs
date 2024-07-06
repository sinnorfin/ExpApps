/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2019 by Attila Kalina <attilakalina.arch@gmail.com>
 *
 * This file is part of Experimental Apps.
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
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using System;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using System.Text.RegularExpressions;
using System.Linq;
using static StoreExp;

public class Versioned_methods
{
    public static Element GetTaggedLocalElements(IndependentTag tag)
    { return tag.GetTaggedLocalElement(); }
    public static List<string> shiftdistances(Document doc)
    {
        List<string> shiftdistances = new List<string>();
        string Range1 = "1/2\"";
        string Range2 = "1\"";
        string Range3 = "3\"";
        string Range4 = "1' 0\"";
        string Range5 = "3' 0\"";
        string Range6 = "10' 0\"";
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
            Range1 = "1 cm"; Range2 = "2 cm"; Range3 = "10 cm";
            Range4 = "30 cm"; Range5 = "90 cm"; Range6 = "300 cm";
        }
        shiftdistances.Add(Range1); shiftdistances.Add(Range2); shiftdistances.Add(Range3);
        shiftdistances.Add(Range4); shiftdistances.Add(Range5); shiftdistances.Add(Range6);
        return shiftdistances;
    }
}
namespace Exp_apps_R17
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MultiTag : IExternalCommand
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
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Tags");
                bool start = true;
                XYZ tagpos = new XYZ(0, 0, 0);
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    Reference tagref = new Reference(elem);
                    IndependentTag tag = null;
                    if (elem.Category.Name.Contains("Tags"))
                    {
                        tag = elem as IndependentTag;
                        if (start)
                        {
                            tagpos = tag.TagHeadPosition;
                            start = false;
                        }
                        else { tag.TagHeadPosition = tagpos; }
                    }
                    else
                    {
                        LocationCurve refline = elem.Location as LocationCurve;
                        // R2021 tag = doc.Create.NewTag(doc.ActiveView, elem, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, refline.Curve.Evaluate(0.5, true));
                 
                        if (start)
                        {
                            tagpos = tag.TagHeadPosition;
                            start = false;
                        }
                        else { tag.TagHeadPosition = tagpos; }
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ApplyInsulations : IExternalCommand
    {
        class Rule
        {
            public string Name { get; set; }
            public double MinS { get; set; }
            public double MaxS { get; set; }
            public double Thickness { get; set; }

            public Rule(string name, double minS, double maxS, double thickness)
            {
                Name = name;
                MinS = minS;
                MaxS = maxS;
                Thickness = thickness;
            }
        }
        static MEPSystem GetNextSystem(ICollection<ElementId> ids, Document doc)
        {
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                if (elem is MEPCurve mepCurve)
                {
                    return mepCurve.MEPSystem;
                }
            }
            return null;
        }
        static void RemoveInsulations(MEPSystem system, Document doc)
        {
            if (system is PipingSystem pipes)
            {
                foreach (Element elem in pipes.PipingNetwork)
                { if (elem is PipeInsulation) doc.Delete(elem.Id); }
            }
            if (system is MechanicalSystem ducts)
            {
                foreach (Element elem in ducts.DuctNetwork)
                { if (elem is DuctInsulation) doc.Delete(elem.Id); }
            }
        }
        static ICollection<ElementId> RemoveElementsfromList(MEPSystem system, ICollection<ElementId> ids, Document doc)
        {
            ICollection<ElementId> newList = new List<ElementId>(ids);
            foreach (ElementId eid in ids)
            {
                Element element = doc.GetElement(eid);
                if (element is MEPCurve mepCurve)
                {
                    if (mepCurve.MEPSystem.Name == system.Name)
                    { newList.Remove(eid); }
                }
                else if (element is FamilyInstance faminst)
                {
                    if (faminst.LookupParameter("System Name").AsString() == system.Name)
                    { newList.Remove(eid); }
                }
            }
            return newList;
        }
        //Applies insulations to pipe/fitting/accessory based on rules regarding
        //System, Min/max sizes
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
            List<MEPSystem> systems = new List<MEPSystem>();
            MEPSystem nextSystem;
            string pattern = @"\d+";
            Regex regex = new Regex(pattern);
            List<Rule> rules = new List<Rule>()
            {
                //new Rule("WW", 20, 25, 20),
                //new Rule("WW", 32, 50, 25),
                //new Rule("KW", 20, 25, 15),
                //new Rule("KW", 32, 40, 30),
                //new Rule("KW", 50, 60, 35)
                new Rule("CVA", 15, 15, 25),
                new Rule("CVA", 20, 25, 30),
                new Rule("CVA", 32, 100, 40),
                new Rule("CVA", 125, 250, 50),
                new Rule("CVR", 15, 15, 25),
                new Rule("CVR", 20, 25, 30),
                new Rule("CVR", 32, 100, 40),
                new Rule("CVR", 125, 250, 50),
                new Rule("KWA", 15, 15, 9),
                new Rule("KWA", 20, 25, 13),
                new Rule("KWA", 32, 50, 19),
                new Rule("KWA", 65, 125, 25),
                new Rule("KWA", 150, 250, 32),
                new Rule("KWR", 15, 15, 9),
                new Rule("KWR", 20, 25, 13),
                new Rule("KWR", 32, 50, 19),
                new Rule("KWR", 65, 125, 25),
                new Rule("KWR", 150, 250, 32),
             
            //Add only visible option
            };
            foreach (Rule rule in rules)
            { rule.Thickness = UnitUtils.ConvertToInternalUnits(rule.Thickness, DisplayUnitType.DUT_MILLIMETERS); }
            Rule selectedRule = null;

            Dictionary<string, List<Rule>> ruleDictionary = rules.GroupBy(r => r.Name)
            .ToDictionary(g => g.Key, g => g.ToList());
            //if (ruleDictionary.TryGetValue(externalName, out List<Rule> selectedList))
            //{
            //    selectedRule = selectedList.FirstOrDefault(r => externalValue >= r.MinS && externalValue <= r.MaxS);
            //}
            //TaskDialog.Show("The selected Rule", selectedRule.Name + " " + selectedRule.MinS + " " + selectedRule.MaxS + " " + selectedRule.Thickness);

            double insulationThickness;
            UnitFormatUtils.TryParse(doc.GetUnits(), UnitType.UT_Length, "20", out insulationThickness);
            while ((nextSystem = GetNextSystem(ids, doc)) != null)
            {
                systems.Add(nextSystem);
                ids = RemoveElementsfromList(nextSystem, ids, doc);
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Remove Insulations");
                foreach (MEPSystem system in systems)
                {
                    RemoveInsulations(system, doc);
                }
                trans.Commit();
                trans.Start("Auto-Apply Insulations");

                foreach (MEPSystem system in systems)
                {
                    if (system is PipingSystem pipingSystem)
                    {
                        foreach (Element elem in pipingSystem.PipingNetwork)
                        {
                            if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_PipeCurves
                                || (BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_FlexPipeCurves)
                            {
                                //Parameter diameterParam = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                                //double convertedDiameter = UnitUtils.ConvertFromInternalUnits(diameterParam.AsDouble(), UnitTypeId.Millimeters);
                                Element systemtype = doc.GetElement(system.GetTypeId());
                                Match match = regex.Match(elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString());
                                //string[] parts = elem.LookupParameter("Size").AsString().Split(' ');

                                double.TryParse(match.Value, out double number);
                                if (ruleDictionary.TryGetValue(systemtype.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM).AsString(), out List<Rule> selectedList))
                                {
                                    selectedRule = selectedList.FirstOrDefault(r => number >= r.MinS && number <= r.MaxS);
                                }
                                //TaskDialog.Show("The selected Rule", selectedRule.Name + " " + selectedRule.MinS + " " + selectedRule.MaxS + " " + selectedRule.Thickness);
                                try
                                {
                                    PipeInsulation.Create(doc, elem.Id, ElementId.InvalidElementId, selectedRule.Thickness);
                                }
                                catch { }
                            }
                            else if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_PipeFitting)
                                //|| (BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_PipeAccessory)
                            {


                                // Find the first match in the input string
                                Match match = regex.Match(elem.LookupParameter("Size").AsString());
                                //string[] parts = elem.LookupParameter("Size").AsString().Split(' ');
                                double.TryParse(match.Value, out double number);


                                Element systemtype = doc.GetElement(system.GetTypeId());
                                if (ruleDictionary.TryGetValue(systemtype.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM).AsValueString(), out List<Rule> selectedList))
                                {
                                    selectedRule = selectedList.FirstOrDefault(r => number >= r.MinS && number <= r.MaxS);
                                }
                                //TaskDialog.Show("The selected Rule", selectedRule.Name + " " + selectedRule.MinS + " " + selectedRule.MaxS + " " + selectedRule.Thickness);
                                try { PipeInsulation.Create(doc, elem.Id, ElementId.InvalidElementId, selectedRule.Thickness); }
                                catch { }
                            }
                        }
                    }
                    else if (system is MechanicalSystem mechanicalSystem)
                    {
                        foreach (Element elem in mechanicalSystem.DuctNetwork)
                        {
                            if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_DuctCurves
                                || (BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_DuctFitting
                                || (BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_DuctAccessory
                                || (BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_FlexDuctCurves)
                            { DuctInsulation.Create(doc, elem.Id, ElementId.InvalidElementId, insulationThickness); }
                        }
                    }
                }
                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CheckTag : IExternalCommand
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
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            GetMenuValue(uiapp);
            string contains = "Round";
            if (Store.menu_3_Box.ToString() != "")
            { contains = "Rectangular"; }
            foreach (ElementId eid in ids)
            {
                IndependentTag tag = doc.GetElement(eid) as IndependentTag;
                string familyname = tag.GetTaggedLocalElement().get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                if (familyname.Contains(contains))
                { newsel.Add(eid); }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Select tags");
                uidoc.Selection.SetElementIds(newsel);
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RehostElements : IExternalCommand
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
            Selection SelectedObjs = uidoc.Selection;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ComboBox selectedlevel = StoreExp.GetComboBox(uiapp.GetRibbonPanels("Exp. Add-Ins"), "View Tools", "ExpLevel");
            Level targetLevel = null;
            Double targetLevelElev = 0;
            Double ElemLevelElev = 0;
            int FaceHosted = 0;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Re-Host Elements");
                
                targetLevel = StoreExp.GetLevel(doc, selectedlevel.Current.ItemText);
                if (targetLevel != null)
                {
                    targetLevelElev = targetLevel.Elevation;
                }
                if (targetLevel == null)
                {
                    if (!(doc.ActiveView is ViewPlan viewPlan))
                    {
                        TaskDialog.Show("Please select Plan View or set level in QuickViews", "Level not selected, or Active view must be a plan view.");
                        tx.Commit();
                        return Result.Succeeded;
                    }
                    Parameter associated = doc.ActiveView.LookupParameter("Associated Level");
                    targetLevel = StoreExp.GetLevel(doc, "Associated");
                    targetLevelElev = targetLevel.Elevation;
                }
                if (SelectedObjs != null)
                {
                    foreach (ElementId eid in ids)
                    {
                        Element e = doc.GetElement(eid);
                        try
                        {
                            Level elemLvl = doc.GetElement(e.LevelId) as Level;
                            if (elemLvl == null)
                            { elemLvl = doc.GetElement(e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()) as Level; }
                            ElemLevelElev = elemLvl.Elevation;
                            if (e.GetType() == typeof(Ceiling))
                            {
                                e.get_Parameter(BuiltInParameter.LEVEL_PARAM).Set(targetLevel.Id);
                                Double currentoffset = e.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).Set(fulloffset);
                            }
                            else if (e.GetType() == typeof(Floor))
                            {
                                e.get_Parameter(BuiltInParameter.LEVEL_PARAM).Set(targetLevel.Id);
                                Double currentoffset = e.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(fulloffset);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Structural Framing")
                            {
                                Double currentoffsetStart = e.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION).AsDouble();
                                Double dehostoffsetStart = currentoffsetStart + 1.0;
                                e.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION).Set(dehostoffsetStart);
                                e.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION).Set(currentoffsetStart);
                                e.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).Set(targetLevel.Id);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Columns")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(targetLevel.Id);
                                e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(fulloffset);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Walls")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).Set(targetLevel.Id);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Roofs")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.ROOF_BASE_LEVEL_PARAM).Set(targetLevel.Id);
                            }
                            else if ((e.LookupParameter("Category").AsValueString() == "Doors") || (e.LookupParameter("Category").AsValueString() == "Windows"))
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(targetLevel.Id);
                            }
                            else if (e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) != null)
                            {
                                FamilyInstance family = e as FamilyInstance;
                                if ((family.HostFace != null) || (e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsString() == "Orphaned") ||
                                    (e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsString() == "<not associated>"))
                                { FaceHosted = FaceHosted + 1; }
                          
                                else
                                {
                                    e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(targetLevel.Id);
                                    Double currentoffset = e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                                    Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                    e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(fulloffset);
                                }
                            }
                            else
                            {
                                e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).Set(targetLevel.Id);
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException) { FaceHosted += 1; }
                    }
                    if (FaceHosted > 0) { TaskDialog.Show("Incompatible", "There were " + FaceHosted + " Incompatible elements selected."); }
                    tx.Commit();
                }
                return Result.Succeeded;
            }
        }
    }
}
