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

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using System.Collections.Generic;
using System;

namespace AlignToBottom
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignToBottom : IExternalCommand
    {
        public static double GetHeight(Element elem)
        {
            double height = 0;
            if (elem.LookupParameter("Category").AsValueString() == "Conduits")
            {
                height = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM).AsDouble() / 2;
            }
            else if (elem.LookupParameter("Category").AsValueString() == "Ducts")
            {
                if (elem.LookupParameter("Family").AsValueString() != "Round Duct")
                { height = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() / 2; }
                else { height = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsDouble() / 2; }
            }
            else if (elem.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) != null)
            {
                height = elem.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() / 2;
            }
            else
            {
                TaskDialog.Show("Error", "Invalid Target. Select Pipe,Duct or Conduit");
            }
            return height;
        }
        public static double GetElev(PushButton toggle_Insulation, Element elem)
        {
            double elev = 0;
            if (elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM) != null)
            {
                if (toggle_Insulation.ItemText == "Align to INS")
                {
                    if (elem.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS) != null)
                    {
                        double ins_thickness = elem.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
                        elev = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble() + ins_thickness;
                    }
                    else { elev = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble(); }
                }
                else { elev = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble(); }
            }
            return elev;
        }
        public static PushButton GetButton(UIApplication uiapp,string panelname, string itemname)
        {
            RibbonPanel inputpanel = null;
            PushButton toggle_Insulation = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == panelname)
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == itemname)
                { toggle_Insulation = (PushButton)item; }
            }
            return toggle_Insulation;
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
            int errorcount = 0;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Align To Bottom");
                {
                    if (ids.Count == 0)
                    {
                        TaskDialog.Show("Error", "No Selection");
                        return Result.Succeeded;
                    }
                    Double ins_thickness = 0;
                    Reference clicked = SelectedObjs.PickObject(ObjectType.Element);
                    Element targetelem = doc.GetElement(clicked);
                    PushButton toggle_Insulation = GetButton(uiapp, "Re-Elevate", "Toggle_Insulation");
                    double targetelev = GetElev(toggle_Insulation, targetelem);
                    double targetR = GetHeight(targetelem);
                    if (targetR == 0) return Result.Succeeded;
                    foreach (ElementId eid in ids)
                    {
                        if (eid != clicked.ElementId)
                        {
                            Element e = doc.GetElement(eid);
                            double currentR = GetHeight(e);
                            if (currentR == 0) errorcount++;
                            if (e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM) != null)
                            {
                                if (toggle_Insulation.ItemText == "Align to INS")
                                {
                                    if (e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS) != null)
                                    {
                                        ins_thickness = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
                                        e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + currentR - targetR + ins_thickness);
                                    }
                                    else { e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + currentR - targetR); }
                                }
                                else { e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + currentR - targetR); }
                            }
                        }
                    }
                }
                if (errorcount > 0)
                    { TaskDialog.Show("There were errors", errorcount + " Invalid Element(s) left unchanged"); }
                    tx.Commit();
                return Result.Succeeded;
            }
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignToTop : IExternalCommand
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
            int errorcount = 0;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Align to Top");
                {
                    if (ids.Count == 0)
                    {
                        TaskDialog.Show("Error", "No Selection");
                        return Result.Succeeded;
                    }
                    Reference clicked = SelectedObjs.PickObject(ObjectType.Element);
                    Element Targetelem = doc.GetElement(clicked);
                    PushButton toggle_Insulation = AlignToBottom.GetButton(uiapp, "Re-Elevate", "Toggle_Insulation");
                    double targetelev = AlignToBottom.GetElev(toggle_Insulation,Targetelem);
                    double targetR = AlignToBottom.GetHeight(Targetelem);
                    if (targetR == 0) return Result.Succeeded;
                    foreach (ElementId eid in ids)
                    {
                        if (eid != clicked.ElementId)
                        {
                            Element e = doc.GetElement(eid);
                            double currentR = AlignToBottom.GetHeight(e);
                            if (currentR == 0) errorcount++;
                            if (e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM) != null)
                            {
                                if (toggle_Insulation.ItemText == "Align to INS")
                                {
                                    if (e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS) != null)
                                    {
                                        double ins_thickness = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
                                        e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + targetR - currentR - ins_thickness);
                                    }
                                    else { e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + targetR - currentR); }
                                }
                                else { e.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(targetelev + targetR - currentR); }
                            }
                        }
                    }
                }
                if (errorcount > 0)
                { TaskDialog.Show("There were errors", errorcount + " Invalid Element(s) left unchanged"); }
                tx.Commit();
                return Result.Succeeded;
            }
        }
    }
}