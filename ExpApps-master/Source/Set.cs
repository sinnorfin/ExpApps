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

using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using System;
using _ExpApps;
using System.Collections.Generic;

namespace SetViewRange
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SetPer3D : IExternalCommand
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

            if (!(doc.ActiveView is ViewPlan viewPlan))
            {
                TaskDialog.Show("Please select Plan view", "Select Plan view to change it's View Range");
                return Result.Succeeded;
            }
            Level level = viewPlan.GenLevel;
            View3D view3d = null;
            string source3d = StoreExp.ThreeDview;
            if (source3d == "Same Name")
            {
                source3d = viewPlan.Name;
            }
            try
            {
                view3d = (from v in new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>() where v.Name == source3d select v).First();
            }
            catch
            {
                TaskDialog.Show("Please rename 3D view or Select in Options", "Rename 3D view to match:" + Environment.NewLine + "'" + viewPlan.Name +
                                "'" + Environment.NewLine + " or, Select source 3D view in Options");
                return Result.Succeeded;
            }
            BoundingBoxXYZ bbox = view3d.GetSectionBox();
            Transform transform = bbox.Transform;
            double bboxOriginZ = transform.Origin.Z;
            double minZ = bbox.Min.Z + bboxOriginZ;
            double maxZ = bbox.Max.Z + bboxOriginZ;

            PlanViewRange viewRange = viewPlan.GetViewRange();

            viewRange.SetLevelId(PlanViewPlane.TopClipPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.CutPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.BottomClipPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.ViewDepthPlane, level.Id);
            if (viewPlan.ViewType == ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.CutPlane, minZ - level.Elevation);
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, maxZ - level.Elevation);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, maxZ - level.Elevation);               
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, minZ - level.Elevation);
            }
            if (viewPlan.ViewType != ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.CutPlane, maxZ - level.Elevation);
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, minZ - level.Elevation);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, minZ - level.Elevation);
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, maxZ - level.Elevation);
            }
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Set View Range");
                viewPlan.SetViewRange(viewRange);
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Shift_BD : IExternalCommand
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
            if (!(doc.ActiveView is ViewPlan viewPlan))
            {
                TaskDialog.Show("Please select Plan view", "Select Plan view to change it's View Range");
                return Result.Succeeded;
            }
            PlanViewRange viewRange = viewPlan.GetViewRange();
            double CCut = viewRange.GetOffset(PlanViewPlane.CutPlane);
            double CBot = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double CVd = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);
            RibbonPanel inputpanel = null;
            ComboBox inputbox = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "View Tools")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "ShiftRange")
                { inputbox = (ComboBox)item; }
            }
            List<Double> vrOpts = new List<Double> { StoreExp.vrOpt1, StoreExp.vrOpt2, StoreExp.vrOpt3,
                StoreExp.vrOpt4, StoreExp.vrOpt5, StoreExp.vrOpt6 };
            double mod = vrOpts[Int32.Parse(inputbox.Current.Name)];

            if (viewPlan.ViewType != ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, CBot - mod);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, CBot - mod);
            }
            if (viewPlan.ViewType == ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.CutPlane, CCut - mod);
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, CCut - mod);
            }
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                try
                {
                    t.Start();
                    viewPlan.SetViewRange(viewRange);
                    t.Commit();
                }
                catch
                { TaskDialog.Show("Error", "Cannot shift ViewRange this way."); }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CopyCrop : IExternalCommand
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
            var view = uidoc.ActiveView;
            
            View selectPlan = null;
            
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                if (elem.Category.Name == "Views")
                {
                    selectPlan = elem as View;
                    TaskDialog.Show("Copying crop from", selectPlan.Title);
                }
            }
            BoundingBoxXYZ box = new BoundingBoxXYZ();
            XYZ minx = new XYZ(selectPlan.CropBox.Min.X, selectPlan.CropBox.Min.Y, selectPlan.CropBox.Min.Z);
            XYZ maxx = new XYZ(selectPlan.CropBox.Max.X, selectPlan.CropBox.Max.Y, selectPlan.CropBox.Max.Z);
            box.Min = minx;
            box.Max = maxx;
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                t.Start();
                view.CropBoxActive = true;
                view.CropBox = box;
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CopyVR : IExternalCommand
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
            ViewPlan viewPlan = uidoc.ActiveView as ViewPlan;
            PlanViewRange VR = null;
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                if (elem.Category.Name == "Views")
                {
                    ViewPlan selectPlan = elem as ViewPlan;
                    VR = selectPlan.GetViewRange();
                    TaskDialog.Show("Copying VR from", selectPlan.Title);
                }
            }
            if (VR == null)
            { TaskDialog.Show("error", "error"); }
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                t.Start();
                viewPlan.SetViewRange(VR);
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Shift_BU : IExternalCommand
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

            if (!(doc.ActiveView is ViewPlan viewPlan))
            {
                TaskDialog.Show("Please select Plan view", "Select Plan view to change it's View Range");
                return Result.Succeeded;
            }
            PlanViewRange viewRange = viewPlan.GetViewRange();
            double CCut = viewRange.GetOffset(PlanViewPlane.CutPlane);
            double CBot = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double CTop = viewRange.GetOffset(PlanViewPlane.TopClipPlane);

            RibbonPanel inputpanel = null;
            ComboBox inputbox = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "View Tools")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "ShiftRange")
                { inputbox = (ComboBox)item; }
            }
            List<Double> vrOpts = new List<Double> { StoreExp.vrOpt1, StoreExp.vrOpt2, StoreExp.vrOpt3,
                StoreExp.vrOpt4, StoreExp.vrOpt5, StoreExp.vrOpt6 };
            double mod = vrOpts[Int32.Parse(inputbox.Current.Name)];
            double rangedepth = CBot + mod + 0.999999999999962;
            if (viewPlan.ViewType != ViewType.CeilingPlan)
            {
                if (CBot + mod + 0.04 >= CCut)
                {
                    viewRange.SetOffset(PlanViewPlane.TopClipPlane, rangedepth);
                    viewRange.SetOffset(PlanViewPlane.CutPlane, rangedepth);
                }
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, CBot + mod);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, CBot + mod);

            }
            if (viewPlan.ViewType == ViewType.CeilingPlan)
            {
                if (CCut + mod + 0.04 >= CTop)
                {
                    viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, rangedepth);
                    viewRange.SetOffset(PlanViewPlane.TopClipPlane, rangedepth);
                }
                viewRange.SetOffset(PlanViewPlane.CutPlane, CCut + mod);
                viewRange.SetOffset(PlanViewPlane.BottomClipPlane, CCut + mod);
            }
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                t.Start();
                viewPlan.SetViewRange(viewRange);
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Shift_TD : IExternalCommand
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

            if (!(doc.ActiveView is ViewPlan viewPlan))
            {
                TaskDialog.Show("Please select Plan view", "Select Plan view to change it's View Range");
                return Result.Succeeded;
            }
            PlanViewRange viewRange = viewPlan.GetViewRange();
            double CBot = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double CCut = viewRange.GetOffset(PlanViewPlane.CutPlane);
            double CTop = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            RibbonPanel inputpanel = null;
            ComboBox inputbox = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "View Tools")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "ShiftRange")
                { inputbox = (ComboBox)item; }
            }
            List<Double> vrOpts = new List<Double> { StoreExp.vrOpt1, StoreExp.vrOpt2, StoreExp.vrOpt3,
                StoreExp.vrOpt4, StoreExp.vrOpt5, StoreExp.vrOpt6 };
            double mod = vrOpts[Int32.Parse(inputbox.Current.Name)];
            double rangedepth = CCut - mod - 0.999999999999962;

            if (viewPlan.ViewType != ViewType.CeilingPlan)
            {
                if (CCut - mod - 0.04 <= CBot)
                {
                    viewRange.SetOffset(PlanViewPlane.BottomClipPlane, rangedepth);
                    viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, rangedepth);
                }
                viewRange.SetOffset(PlanViewPlane.CutPlane, CCut - mod);
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, CCut - mod);
            }
            if (viewPlan.ViewType == ViewType.CeilingPlan)
            {
                if (CTop - mod - 0.04 <= CCut)
                {
                    viewRange.SetOffset(PlanViewPlane.BottomClipPlane, rangedepth);
                    viewRange.SetOffset(PlanViewPlane.CutPlane, rangedepth);
                }
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, CTop - mod);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, CTop - mod);
            }
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                {
                    t.Start();
                    viewPlan.SetViewRange(viewRange);
                    t.Commit();
                }
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Shift_TU : IExternalCommand
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

            if (!(doc.ActiveView is ViewPlan viewPlan ))
            {
                TaskDialog.Show("Please select Plan view", "Select Plan view to change it's View Range");
                return Result.Succeeded;
            }
            PlanViewRange viewRange = viewPlan.GetViewRange();
            double CCut = viewRange.GetOffset(PlanViewPlane.CutPlane);
            double CTop = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            RibbonPanel inputpanel = null;
            ComboBox inputbox = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "View Tools")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "ShiftRange")
                { inputbox = (ComboBox)item; }
            }
            List<Double> vrOpts = new List<Double> { StoreExp.vrOpt1, StoreExp.vrOpt2, StoreExp.vrOpt3,
                StoreExp.vrOpt4, StoreExp.vrOpt5, StoreExp.vrOpt6 };
            double mod = vrOpts[Int32.Parse(inputbox.Current.Name)];

            if (viewPlan.ViewType != ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.CutPlane, CCut + mod);
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, CCut + mod);
            }
            if (viewPlan.ViewType == ViewType.CeilingPlan)
            {
                viewRange.SetOffset(PlanViewPlane.TopClipPlane, CTop + mod);
                viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, CTop + mod);
            }
            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                try
                {
                    t.Start();
                    viewPlan.SetViewRange(viewRange);
                    t.Commit();
                }
                catch
                { TaskDialog.Show("Error", "Cannot shift ViewRange this way."); }
            }
            return Result.Succeeded;
        }
    }
}
