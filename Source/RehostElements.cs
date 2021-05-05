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
using _ExpApps;

namespace RehostElements
{
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
            ElementId targetLevel = null;
            Double targetLevelElev =  0;
            Double ElemLevelElev = 0;
            int FaceHosted = 0;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Re-Host Elements");
                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                ICollection<Element> lvlCollection = lvlCollector.OfClass(typeof(Level)).ToElements();
                foreach (Element level in lvlCollection)
                {
                    Level lvl = level as Level;
                    if (level.Name == StoreExp.level)
                    {
                        targetLevel = lvl.Id;
                        targetLevelElev = lvl.Elevation;
                    }
                    else if (StoreExp.level == "Active PlanView")
                    {
                        targetLevel = null;
                    }
                }
                if (targetLevel == null)
                {
                    if (!(doc.ActiveView is ViewPlan viewPlan ))
                    {
                        TaskDialog.Show("Please select Plan View or set level in QuickViews", "Level not selected, or Active view must be a plan view.");
                        tx.Commit();
                        return Result.Succeeded;
                    }
                    Parameter associated = doc.ActiveView.LookupParameter("Associated Level");
                    foreach (Element level in lvlCollection)
                    {
                        Level lvl = level as Level;
                        if (lvl.Name == associated.AsString())
                        {
                            targetLevel = lvl.Id;
                            targetLevelElev = lvl.Elevation;
                        }
                    }
                }
                if (SelectedObjs != null)
                {
                    foreach (ElementId eid in ids)
                {Element e = doc.GetElement(eid);
                        try
                        {
                            foreach (Element level in lvlCollection)
                            {
                                Level lvl = level as Level;
                                if (lvl.Id == e.LevelId)
                                {
                                    ElemLevelElev = lvl.Elevation;
                                }
                            }
                            if (e.GetType() == typeof(Ceiling))
                            {
                                e.get_Parameter(BuiltInParameter.LEVEL_PARAM).Set(targetLevel);
                                Double currentoffset = e.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).Set(fulloffset);
                            }
                            else if (e.GetType() == typeof(Floor))
                            {
                                e.get_Parameter(BuiltInParameter.LEVEL_PARAM).Set(targetLevel);
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
                                e.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).Set(targetLevel);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Columns")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(targetLevel);
                                e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(fulloffset);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Walls")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).Set(targetLevel);
                            }
                            else if (e.LookupParameter("Category").AsValueString() == "Roofs")
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.ROOF_BASE_LEVEL_PARAM).Set(targetLevel);
                            }
                            else if ((e.LookupParameter("Category").AsValueString() == "Doors") || (e.LookupParameter("Category").AsValueString() == "Windows"))
                            {
                                Double currentoffset = e.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).AsDouble();
                                Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                e.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(fulloffset);
                                e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(targetLevel);
                            }
                            else if (e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) != null)
                            {
                                FamilyInstance family = e as FamilyInstance;
                                if ((family.HostFace != null) || (e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsString() == "Orphaned") ||
                                    (e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsString() == "<not associated>"))
                                { FaceHosted = FaceHosted + 1; }
                                else if (e.Name.Contains("YeahSo"))
                                {
                                    e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(targetLevel);
                                    Double currentoffset = e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                                    Double Old_AnchorElevation = e.LookupParameter("AnchorElevation").AsDouble();
                                    Double Old_CenterElevation = e.LookupParameter("CenterElevation").AsDouble();
                                    e.LookupParameter("AnchorElevation").Set(ElemLevelElev - targetLevelElev + currentoffset + Old_AnchorElevation);
                                    e.LookupParameter("CenterElevation").Set(ElemLevelElev - targetLevelElev + currentoffset + Old_CenterElevation);
                                    e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(0);
                                }
                                else
                                {
                                    e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(targetLevel);
                                    Double currentoffset = e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                                    Double fulloffset = ElemLevelElev - targetLevelElev + currentoffset;
                                    e.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(fulloffset);
                                }
                            }
                            else
                            {
                                e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).Set(targetLevel);
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
