﻿/**
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
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.Reflection;


namespace AnnoTools
{
    public static class StoreAnno
    {
        public static Double mod_place = 0.5; public static Double mod_left = 1; public static Double mod_right = 1;
        public static Double mod_firsty = 1; public static Double mod_stepy = 1; public static int mod_split = 0;
        public static TextBox left_ib = null; public static TextBox right_ib = null; public static TextBox firsty_ib = null;
        public static TextBox stepy_ib = null; public static TextBox split_ib = null; public static TextBox place_ib = null;
        public static Double scaleMod = 1;
        public static Options Dimop(View ActiveView)
        {
            Options dimop = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = true,
                View = ActiveView
            };
            return dimop;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    // Aligns tags in selection if they contain the same text and are close enough.
    // Exceptions not added.
    // Uses the same menu values that AnnoTools.RackDim uses.

    public class Cleansheet : IExternalCommand
    {
        public Result Execute(
               ExternalCommandData commandData,
               ref string message,
               ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<IndependentTag> prevtags = new List<IndependentTag>();

            // Menu Values have to be acquired from Ribbon.
            RackDim.GetMenuValues(uiapp);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Merge Tags");
                bool start = true;
                bool fits = false;
                XYZ tagpos = new XYZ(0, 0, 0);
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    IndependentTag tag = elem as IndependentTag;
                    if (start)
                    {
                        tagpos = tag.TagHeadPosition;
                        start = false;
                        prevtags.Add(tag);
                    }
                    else
                    {
                        fits = false;
                        foreach (IndependentTag ptag in prevtags)
                        {
                            if (fits == false &&
                                ptag.TagText == tag.TagText &&
                                Math.Abs(ptag.TagHeadPosition.X - tag.TagHeadPosition.X) < StoreAnno.mod_left &&
                                Math.Abs(ptag.TagHeadPosition.Y - tag.TagHeadPosition.Y) < StoreAnno.mod_right)
                            {
                                tag.TagHeadPosition = ptag.TagHeadPosition;
                                fits = true;
                            }
                        }
                        if (!fits)
                        {
                            prevtags.Add(tag);
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
    public class RehostMultiCategorytoNearest : IExternalCommand
    {
        public Result Execute(
               ExternalCommandData commandData,
               ref string message,
               ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            List<IndependentTag> tags = new List<IndependentTag>();
            List<(Element,XYZ)> hosts = new List<(Element,XYZ)>();
            using (Transaction trans = new Transaction(doc))
            {
                foreach (ElementId eid in ids)
                {
                    if (doc.GetElement(eid) is IndependentTag tag)
                    { if (tag.GetTaggedReferences().Count == 0) 
                        { tags.Add(tag); }
                    }
                    else 
                    {
                        Element elem = doc.GetElement(eid);
                        if (elem.Location is LocationCurve refcurve)
                        {
                            hosts.Add((elem,refcurve.Curve.Evaluate(0.5, true)));
                        }
                        else if (elem.Location is LocationPoint refpoint)
                        {
                            hosts.Add((elem, refpoint.Point));
                        }
                    }
                }
                trans.Start("Rehost Tags to Nearest");
                foreach (IndependentTag tag in tags)
                {
                    Element closestElem = hosts.OrderBy(p => p.Item2.DistanceTo(tag.TagHeadPosition)).First().Item1;
                    Reference newref = new Reference(closestElem);
                    IndependentTag newtag = IndependentTag.Create(doc,tag.GetTypeId(), doc.ActiveView.Id, newref, false,
                                TagOrientation.Horizontal,
                                tag.TagHeadPosition);
                    doc.Delete(tag.Id);
                    newsel.Add(newtag.Id);
                }
                trans.Commit();

            }
            uidoc.Selection.SetElementIds(newsel);
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
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ICollection<ElementId> newsel = new List<ElementId>();
            string contains = "Round";
            if (StoreExp.GetSwitchStance(uiapp, "Red"))
            { contains = "Rectangular"; }
            foreach (ElementId eid in ids)
            {
                IndependentTag tag = doc.GetElement(eid) as IndependentTag;
                string familyname = tag.GetTaggedLocalElements().First().get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
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
    public class RoomLeaderOff : IExternalCommand
    {
        public Result Execute(
               ExternalCommandData commandData,
               ref string message,
               ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> newSel = new List<ElementId>();
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<SpatialElementTag> tags = new List<SpatialElementTag>();
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                SpatialElementTag tag;
                if (elem.Category.Name.Contains("Tags"))
                {
                    tag = elem as SpatialElementTag;
                    tags.Add(tag);
                }
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Remove Room Tag-Leaders (same pos) ");
               
                    foreach (SpatialElementTag tag in tags)
                    {
                    try
                    {
                        XYZ position = tag.TagHeadPosition;
                        tag.HasLeader = false;
                        tag.TagHeadPosition = position;
                    }
                    catch { newSel.Add(tag.Id); }
                }
                uidoc.Selection.SetElementIds(newSel);
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FixFreeLeaders : IExternalCommand
    {
        public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Fix Free leader Tags");
                foreach (ElementId eid in ids)
                {
                    Element elem = doc.GetElement(eid);
                    if (elem is IndependentTag tag)
                    {
                         Element firstreference = tag.GetTaggedLocalElements().First();
                        LocationPoint locpoint = firstreference.Location as LocationPoint;
                        
                         tag.LeaderEndCondition = LeaderEndCondition.Free;
                         tag.SetLeaderEnd(tag.GetTaggedReferences().First(), locpoint.Point);
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
[Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MultiTag : IExternalCommand
    {
        public Element GetTaggedLocalElements(IndependentTag tag)
        { return tag.GetTaggedLocalElements().First(); }
        public void ReadTag(Document doc, UIDocument uidoc, UIApplication uiapp)
        {
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                IndependentTag tag;
                XYZ shift = new XYZ(0, 0, 0);
                if (elem.Category.Name.Contains("Tags"))
                {
                    tag = elem as IndependentTag;
                    string version = doc.Application.VersionName;
                    Element element;
                    if (version.Contains("202") && !version.Contains("2021") && !version.Contains("2020"))
                    { element = GetTaggedLocalElements(tag); }
                    else { element = Versioned_methods.GetTaggedLocalElements(tag); }
                    
                    if (element.Location is LocationCurve)
                    {
                        LocationCurve refcurve = element.Location as LocationCurve;
                        XYZ refpoint = refcurve.Curve.Evaluate(0.5, true);
                        shift = tag.TagHeadPosition.Subtract(refpoint);
                    }
                    else if (element.Location is LocationPoint)
                    {
                        LocationPoint refpoint = element.Location as LocationPoint;
                        shift = tag.TagHeadPosition.Subtract(refpoint.Point);
                    }
                    StoreExp.tag_shift = shift;       
                    StoreExp.tag_leader = tag.HasLeader;
                    StoreExp.tag_orientation = tag.TagOrientation;
                }
                break;
            }
            GUI.togglebutton(uiapp, "Universal Modifiers", "Universal Toggle Green OFF", "Universal Toggle Green ON");
        }
        public Result Execute(
               ExternalCommandData commandData,
               ref string message,
               ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            string version = doc.Application.VersionName;
            ICollection<ElementId> newSel = new List<ElementId>();
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (StoreExp.GetSwitchStance(uiapp, "Green"))
                { ReadTag(doc, uidoc, uiapp);
                return Result.Succeeded;}
            bool mod_NoMerge = StoreExp.GetSwitchStance(uiapp, "Red");
            bool mod_SavedRelation = StoreExp.GetSwitchStance(uiapp, "Blue");
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
                    if (elem.Category.Name.Contains("Tags") && !mod_NoMerge)
                    {
                        tag = elem as IndependentTag;
                        if (start)
                        {
                            tagpos = tag.TagHeadPosition;
                            start = false;
                        }
                        else { tag.TagHeadPosition = tagpos;}
                    }
                    else if (elem.Category.Name.Contains("Tags") && mod_NoMerge)
                    {
                        Element element;
                        XYZ refpoint = new XYZ(0,0,0);
                        tag = elem as IndependentTag;
                        
                            if (version.Contains("202") && !version.Contains("2021") && !version.Contains("2020"))
                            { element = GetTaggedLocalElements(tag); }
                            else { element = Versioned_methods.GetTaggedLocalElements(tag); }

                            if (element.Location is LocationCurve)
                            {
                                LocationCurve refcurve = element.Location as LocationCurve;
                                refpoint = refcurve.Curve.Evaluate(0.5, true);
                            }
                            else if (element.Location is LocationPoint)
                            {
                                LocationPoint locpoint = element.Location as LocationPoint;
                                refpoint = locpoint.Point;
                            }
                            tag.TagOrientation = StoreExp.tag_orientation;
                            tag.TagHeadPosition = refpoint;
                           
                        if (mod_SavedRelation)
                        {
                            tag.TagHeadPosition = tag.TagHeadPosition.Add(StoreExp.tag_shift);
                        }
                        
                    }
                    else if ( !elem.Category.Name.Contains("Tags"))
                    {
                        if (elem.Location is LocationCurve)
                        {
                            LocationCurve refline = elem.Location as LocationCurve;
                            tag = IndependentTag.Create(doc, doc.ActiveView.Id, tagref, false,
                                TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal,
                                refline.Curve.Evaluate(0.5, true));
                        }
                        else if (elem.Location is LocationPoint)
                        {
                            LocationPoint refpoint = elem.Location as LocationPoint;
                            tag = IndependentTag.Create(doc, doc.ActiveView.Id, tagref, false,
                                TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal,
                                refpoint.Point);
                        }
                            if (start && !mod_NoMerge)
                        {
                            tagpos = tag.TagHeadPosition;
                            start = false;
                        }
                        else if (!mod_NoMerge){ tag.TagHeadPosition = tagpos;}
                        tag.HasLeader = StoreExp.tag_leader;
                        if (mod_NoMerge && mod_SavedRelation)
                        {
                            tag.TagOrientation = StoreExp.tag_orientation;
                            tag.TagHeadPosition = tag.TagHeadPosition.Add(StoreExp.tag_shift);
                        }
                    }
                    try
                    {newSel.Add(tag.Id);}
                    catch { TaskDialog.Show("Error", "Not applicable to selection"); }
                }
                uidoc.Selection.SetElementIds(newSel);
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RackDim : IExternalCommand
    {
        public static void GetMenuValues(UIApplication uiapp)
        {
            double scale = uiapp.ActiveUIDocument.ActiveView.Scale;
            StoreAnno.scaleMod =  96 /scale ;
            RibbonPanel inputpanel = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "Universal Modifiers")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "A")
                { StoreAnno.left_ib = (TextBox)item; }
                else if (item.Name == "B")
                { StoreAnno.right_ib = (TextBox)item; }
                else if (item.Name == "C")
                { StoreAnno.firsty_ib = (TextBox)item; }
                else if (item.Name == "1")
                { StoreAnno.stepy_ib = (TextBox)item; }
                else if (item.Name == "2")
                { StoreAnno.split_ib = (TextBox)item; }
                else if (item.Name == "3")
                { StoreAnno.place_ib = (TextBox)item; }
            }
            Double.TryParse(StoreAnno.left_ib.Value as string,out StoreAnno.mod_left);
            Double.TryParse(StoreAnno.right_ib.Value as string, out StoreAnno.mod_right);
            Double.TryParse(StoreAnno.firsty_ib.Value as string, out StoreAnno.mod_firsty);
            Double.TryParse(StoreAnno.stepy_ib.Value as string, out StoreAnno.mod_stepy);
            int.TryParse(StoreAnno.split_ib.Value as string, out StoreAnno.mod_split);
            Double.TryParse(StoreAnno.place_ib.Value as string, out StoreAnno.mod_place);
            double mod = 1 / StoreAnno.scaleMod;
            if (StoreAnno.mod_left == 0) { StoreAnno.mod_left = mod; } else { StoreAnno.mod_left /= StoreAnno.scaleMod; }
            if (StoreAnno.mod_right == 0) { StoreAnno.mod_right = mod; } else { StoreAnno.mod_right /= StoreAnno.scaleMod; }
            if (StoreAnno.mod_firsty == 0) { StoreAnno.mod_firsty = mod; } else { StoreAnno.mod_firsty /= StoreAnno.scaleMod; }
            if (StoreAnno.mod_stepy == 0) { StoreAnno.mod_stepy =  mod; } else { StoreAnno.mod_stepy /= StoreAnno.scaleMod; }
        }
        public static Line GetPerpendicular(Line dimDir, Double mod_place)
        {
            XYZ cent = dimDir.Evaluate(mod_place, true);
            XYZ dir = new XYZ(0, 0, 1);
            XYZ cross = dimDir.Direction.Normalize().CrossProduct(dir);
            XYZ end = cent + cross;
            return Line.CreateBound(cent, end);
        }
        public static Line GetLineOfGeom(GeometryElement mep)
        {
            Line refLine = null; 
            foreach (var item in mep)
            {
                refLine = item as Line;
                if (refLine != null)
                {
                    return refLine;
                }
            }
            return refLine;
        }
        public static Curve GetRotatedToVert(ElementId original, Document doc, XYZ rot = null)
        {
            Transform toVert; Line origLine;
            LocationCurve origCurve = doc.GetElement(original).Location as LocationCurve;
            XYZ startP = origCurve.Curve.GetEndPoint(0);
            XYZ endP = origCurve.Curve.GetEndPoint(1);
            origLine = Line.CreateBound(new XYZ(startP.X, startP.Y, startP.Z),
            new XYZ(endP.X, endP.Y, endP.Z));
            double Angle = origLine.Direction.AngleTo(new XYZ(0, 1, 0));
            toVert = Transform.CreateRotation(new XYZ(0, 0, 1), Angle);
            if (rot == null && Angle <= Math.PI / 2) { Angle = Math.PI - Angle; toVert = Transform.CreateRotation(new XYZ(0, 0, -1), Angle); }
            else if (rot != null && Angle <= Math.PI / 2) { toVert = Transform.CreateRotationAtPoint(new XYZ(0, 0, -1), Angle,rot); }
            else if (rot != null && Angle > Math.PI/2) { toVert = Transform.CreateRotationAtPoint(new XYZ(0, 0, 1), Angle,rot); }
            return origLine.CreateTransformed(toVert);
        }
        public static void CreateLinearDim(Document doc, Line dimDir, Line dimDirCross, ReferenceArray dimto)
        {
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Linear Dimension");
                Dimension dim = doc.Create.NewDimension(doc.ActiveView, dimDirCross, dimto);
                int count = 0;
                if (dim.Segments.Size == 0) { tx.Commit(); return; }
                foreach (DimensionSegment segment in dim.Segments)
                {
                    if (count == 0 || count == dim.Segments.Size - 1)
                    {
                        if (segment.ValueString.Length * 0.3 > segment.Value * StoreAnno.scaleMod)
                        {
                            if (count == 0) { segment.TextPosition = segment.TextPosition
                                    .Subtract(dimDirCross.Direction.Normalize()
                                .Multiply(StoreAnno.mod_left * (segment.ValueString.Length / 2.5))); }
                            else { segment.TextPosition = segment.TextPosition
                                    .Add(dimDirCross.Direction.Normalize()
                                .Multiply(StoreAnno.mod_right * (segment.ValueString.Length / 2.5))); }
                        }
                    }
                    else
                    {
                        DimensionSegment lastdim = dim.Segments.get_Item(count - 1);
                        DimensionSegment nextdim = dim.Segments.get_Item(count + 1);
                        if (segment.ValueString.Length * 0.45 < segment.Value * StoreAnno.scaleMod) { }
                        else if (lastdim.ValueString.Length * 0.6 < lastdim.Value * StoreAnno.scaleMod &&
                                    nextdim.ValueString.Length * 0.6 < nextdim.Value * StoreAnno.scaleMod) { }
                        else
                        {
                            if (segment.ValueString == lastdim.ValueString)
                            { segment.LeaderEndPosition = lastdim.LeaderEndPosition; }
                            else
                            {
                                segment.LeaderEndPosition = lastdim.LeaderEndPosition
                                    .Subtract(dimDir.Direction.Normalize()
                                    .Multiply(StoreAnno.mod_stepy * 1.4));
                            }
                        }
                    }
                    count += 1;}
                tx.Commit();}
        }
        public static void CreateRackDim(Document doc,Line dimDir,Line dimDirCross, ReferenceArray dimto)
        {
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Rack Dimension");
                Dimension dim = doc.Create.NewDimension(doc.ActiveView, dimDirCross, dimto);
                int count = 0;
                if (dim.Segments.Size == 0) { tx.Commit(); return; }
                if (StoreAnno.mod_split == 0) { StoreAnno.mod_split = dim.Segments.Size / 2; }
                else if (StoreAnno.mod_split == 1) { StoreAnno.mod_split = 0; }
                double modF = StoreAnno.mod_stepy * 1.4;
                double modL = (1.4 * StoreAnno.mod_stepy * (dim.Segments.Size - StoreAnno.mod_split - 1));
                int fHeightCount = 0;
                DimensionSegment FirstDim = dim.Segments.get_Item(0);
                DimensionSegment LastDim = dim.Segments.get_Item(dim.Segments.Size - 1);
                FirstDim.LeaderEndPosition = FirstDim.LeaderEndPosition
                    .Subtract(dimDirCross.Direction.Normalize()
                    .Multiply(StoreAnno.mod_left * 3.5));
                FirstDim.LeaderEndPosition = FirstDim.LeaderEndPosition
                    .Subtract(dimDir.Direction.Normalize()
                    .Multiply(StoreAnno.mod_firsty * 2.8));
                LastDim.LeaderEndPosition = LastDim.LeaderEndPosition
                    .Add(dimDirCross.Direction.Normalize()
                    .Multiply(StoreAnno.mod_right * 3.5));
                LastDim.LeaderEndPosition = LastDim.LeaderEndPosition
                    .Subtract(dimDir.Direction.Normalize()
                    .Multiply(StoreAnno.mod_firsty * 2.8));
                foreach (DimensionSegment segment in dim.Segments)
                {
                    if (count < StoreAnno.mod_split)
                    {
                        if (count != 0)
                        {
                            if (segment.ValueString == dim.Segments.get_Item(count - 1).ValueString)
                            { segment.LeaderEndPosition = dim.Segments.get_Item(count - 1).LeaderEndPosition; }
                            else
                            {
                                segment.LeaderEndPosition = FirstDim.LeaderEndPosition.Subtract(dimDir.Direction.Normalize()
                                    .Multiply(modF));
                                modF += StoreAnno.mod_stepy * 1.4;
                            }
                        }
                        count += 1;
                    }
                    else
                    {
                        if (count != StoreAnno.mod_split && segment.ValueString == dim.Segments.get_Item(count - 1).ValueString)
                        { segment.LeaderEndPosition = dim.Segments.get_Item(count - 1).LeaderEndPosition; fHeightCount += 1; }
                        else
                        {
                            segment.LeaderEndPosition = LastDim.LeaderEndPosition.Subtract(dimDir.Direction.Normalize()
                                .Multiply(modL));
                            modL -= StoreAnno.mod_stepy * 1.4;
                        }
                        count += 1;
                    }
                }
                count = 0;
                if (fHeightCount != 0)
                    {
                    foreach (DimensionSegment segment in dim.Segments)
                    {
                        if (count >= StoreAnno.mod_split)
                        { segment.LeaderEndPosition = segment.LeaderEndPosition.Add(dimDir.Direction.Normalize()
                            .Multiply(StoreAnno.mod_stepy * 1.4 * fHeightCount)); }
                        count += 1;}
                }
                tx.Commit();}
        }
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Line dimDirCross = null; Line dimDir = null;
            int selectdim = 0;
            bool updateMode = false;
            ElementArray dimensions = new ElementArray();
            ReferenceArray dimto = new ReferenceArray();
            GetMenuValues(uiapp);
            foreach (ElementId eid in ids)
            {
                if (doc.GetElement(eid).GetType() == typeof(Grid))
                {
                    Grid grid = doc.GetElement(eid) as Grid;
                    dimto.Append(new Reference(grid));
                }
                else if (doc.GetElement(eid).GetType() == typeof(Dimension))
                {
                    updateMode = true;
                    dimensions.Append(doc.GetElement(eid));
                }
                else {
                    GeometryElement geom = doc.GetElement(eid).get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                    Line refLine = GetLineOfGeom(geom);
                    if (selectdim == 0)
                    {
                        dimDir = Line.CreateBound(refLine.GetEndPoint(0), refLine.GetEndPoint(1));
                        dimDirCross = RackDim.GetPerpendicular(dimDir, StoreAnno.mod_place);
                        selectdim = 1;
                    }
                    dimto.Append(refLine.Reference);
                }
            }
            if (updateMode)
            {
                foreach (Element elem in dimensions)
                {
                    selectdim = 0;
                    Dimension dim = elem as Dimension;
                    ReferenceArray updatedref = new ReferenceArray();
                    foreach (Reference refe in dim.References)
                    {
                        if (doc.GetElement(refe.ElementId).GetType() == typeof(Grid))
                        {
                            Grid grid = doc.GetElement(refe.ElementId) as Grid;
                            updatedref.Append(new Reference(grid));
                        }
                        else
                        {                                 
                            GeometryElement geom = doc.GetElement(refe.ElementId).get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                            Line refLine = GetLineOfGeom(geom);
                            if (selectdim == 0)
                            {
                                dimDir = Line.CreateBound(refLine.GetEndPoint(0), refLine.GetEndPoint(1));
                                dimDirCross = RackDim.GetPerpendicular(dimDir, StoreAnno.mod_place);
                                selectdim = 1;
                            }
                            updatedref.Append(refLine.Reference);
                        }
                    }
                    RackDim.CreateRackDim(doc,dimDir, dimDirCross,updatedref);
                    using (Transaction deltrans = new Transaction(doc))
                    {
                        deltrans.Start("Update Dimension");
                        doc.Delete(elem.Id);
                        deltrans.Commit();
                    }
                }
            }
            else { RackDim.CreateRackDim(doc, dimDir, dimDirCross, dimto); }
            return Result.Succeeded;
        }      
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Rack : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            ElementId defTextType = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
            List<double> cDiams = new List<double>();
            Line LineDir = null; bool updateRack = false;
            List<string> Conduits = new List<string>();
            TextNote toUpdate = null;
            RackDim.GetMenuValues(uiapp);
            StoreAnno.mod_split = 30;
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                if (elem.GetType() == typeof(TextNote))
                    { updateRack = true; toUpdate = elem as TextNote;
                }
                else
                {
                    GeometryElement geom = elem.get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                    cDiams.Add(elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble());
                    LineDir = RackDim.GetLineOfGeom(geom);
                }
            }
            if (toUpdate != null) { ids.Remove(toUpdate.Id); }
            double dMax = cDiams.Max();
            foreach (string s in Conduits) { TaskDialog.Show("asd", s); }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Text");
                List<ElementId> Rows = new List<ElementId>();
                List<List<ElementId>> allRows = SeparateRows(ids,doc,dMax);
                string diameters = ""; Double Distance = 0;
                List<double> leftXs = new List<double>(); List<ElementId> lefts = new List<ElementId>();
                List<double> rightXs = new List<double>(); List<ElementId> rights = new List<ElementId>();
                foreach (int i in SortRows(allRows, doc))
                {
                    string prev = ""; int typcount = 1; string diam;
                    if (i == 0 && allRows.Count > 1) { diameters += "ABOVE: "; }
                    else if (i == allRows.Count - 1 && allRows.Count != 1) { diameters += '\n' + "BELOW: "; }
                    else if (allRows.Count > 1) { diameters += '\n' + "MIDDLE: "; }
                    List<ElementId> SortedRows = SortX(allRows[i], doc);
                    leftXs.Add(RackDim.GetRotatedToVert(SortedRows[0], doc).GetEndPoint(0).X); lefts.Add(SortedRows[0]);
                    rightXs.Add(RackDim.GetRotatedToVert(SortedRows[SortedRows.Count - 1], doc).GetEndPoint(0).X); rights.Add(SortedRows[SortedRows.Count - 1]);
                    foreach (ElementId eid in SortedRows)
                    {
                        diam = doc.GetElement(eid).get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsValueString();

                        if (diam == prev)
                        {
                            typcount += 1;
                        }
                        else
                        {
                            if (typcount != 1)
                            {
                                diameters = diameters + typcount + "-" + prev + "C, ";
                                diameters = BreakLine(diameters, StoreAnno.mod_split);
                                typcount = 1;
                            }
                            else if (prev != "")
                            {
                                diameters = diameters + prev + "C, ";
                                diameters = BreakLine(diameters, StoreAnno.mod_split);
                            }
                            prev = diam;
                        }
                    }
                    if (typcount != 1) { diameters = diameters + typcount + "-" + prev + "C"; }
                    else { diameters = diameters + prev + "C"; }
                }
                ElementId firstEid = lefts[leftXs.IndexOf(leftXs.Max())];
                ElementId lastEid = rights[rightXs.IndexOf(rightXs.Min())];
                LocationCurve firstP = doc.GetElement(firstEid).Location as LocationCurve;
                XYZ rotP = firstP.Curve.GetEndPoint(0);
                Distance = Math.Abs(RackDim.GetRotatedToVert(firstEid,doc,rotP).GetEndPoint(0).X - RackDim.GetRotatedToVert(lastEid,doc,rotP).GetEndPoint(0).X);
                Line firstL = firstP.Curve as Line;
                if (updateRack)
                {
                    toUpdate.Text = diameters;
                }
                else
                {
                    Line texRot = Line.CreateBound(firstP.Curve.GetEndPoint(0), firstP.Curve.GetEndPoint(1));
                    Double textAngle = texRot.Direction.AngleTo(new XYZ(0, -1, 0));
                    bool fix = false;
                    if (textAngle >= Math.PI / 2)
                    {
                        textAngle = Math.PI - textAngle;
                        fix = true;
                    }
                    TextNoteOptions textRotate = new TextNoteOptions
                    {
                        TypeId = defTextType, 
                        Rotation = textAngle
                    };
                    TextNote text = TextNote.Create(doc, doc.ActiveView.Id, LineDir.Evaluate(StoreAnno.mod_place, true), diameters, textRotate);
                    text.get_Parameter(BuiltInParameter.TEXT_ALIGN_HORZ).Set((Int32)TextAlignFlags.TEF_ALIGN_BOTTOM);
                    if (fix)
                    {
                        text.Coord = firstP.Curve.Evaluate(StoreAnno.mod_place, true)
                            .Add(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(Distance + (4* StoreAnno.mod_firsty)));
                    }
                    else
                    {
                        text.Coord = firstP.Curve.Evaluate(StoreAnno.mod_place, true)
                            .Subtract(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(Distance + (4* StoreAnno.mod_firsty)));
                    }
                    text.Coord = new XYZ(text.Coord.X, text.Coord.Y + (1.4* StoreAnno.mod_stepy), text.Coord.Z);   
                    text.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R); text.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R);
                    IList<Leader> leaders = text.GetLeaders();
                    leaders[0].End = firstP.Curve.Evaluate(StoreAnno.mod_place, true);
                    if (firstL.Direction.AngleTo(new XYZ(0, 1, 0)) >= Math.PI / 2)
                    { leaders[1].End = leaders[0].End.Subtract(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                        .Direction.Multiply(Distance)); }
                    else { leaders[1].End = leaders[0].End.Add(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                        .Direction.Multiply(Distance)); }
                    if (fix)
                    {
                        leaders[0].Elbow = leaders[0].End.Subtract(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(1.4 * StoreAnno.mod_left));
                        leaders[1].Elbow = leaders[1].End.Add(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(1.4 * StoreAnno.mod_left));
                    }
                    else {
                        leaders[0].Elbow = leaders[0].End.Add(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(1.4 * StoreAnno.mod_left));
                        leaders[1].Elbow = leaders[1].End.Subtract(RackDim.GetPerpendicular(firstL, StoreAnno.mod_place)
                            .Direction.Multiply(1.4 * StoreAnno.mod_left));
                    }
                }
                tx.Commit(); }
            return Result.Succeeded;
        }
        public List<List<ElementId>> SeparateRows(ICollection<ElementId> sRows, Document doc, double dMax)
        {
            List<List<ElementId>> sortedRows = SortZ(sRows, doc, dMax);
            List<List<ElementId>> rows = new List<List<ElementId>>();
            if (sortedRows[2].Count == 0)
            {
                List<List<ElementId>> AboveRows = SeparateRows(sortedRows[0], doc, dMax);
                List<List<ElementId>> BelowRows = SeparateRows(sortedRows[1], doc, dMax);
                foreach (List<ElementId> row in AboveRows) { rows.Add(row); }
                foreach (List<ElementId> row in BelowRows) { rows.Add(row); }
            }
            else { rows.Add(sortedRows[2]); }
            return rows;
        }
        public List<int> SortRows(List<List<ElementId>> rows,Document doc)
        {
            int countElem ;
            List<int> rowIndex = new List<int>();
            foreach (List<ElementId> row in rows)
            { countElem = rows.Count;
                foreach (List<ElementId> compRow in rows)
                {
                    if (row != compRow)
                    {   LocationCurve lRow = doc.GetElement(row[0]).Location as LocationCurve;
                        LocationCurve lCompRow = doc.GetElement(compRow[0]).Location as LocationCurve;
                        if (lRow.Curve.Evaluate(0.5,true).Z > lCompRow.Curve.Evaluate(0.5, true).Z) { countElem -= 1; } }
                }
                rowIndex.Add(countElem -1 );
            }
            return rowIndex;
        }
        public string BreakLine(string str, int mod_split)
        {
            int numLines = str.Split('\n').Length;
            if ((str.Length+numLines-1) > (mod_split * numLines))
            { return str + '\n'; }
            else return str;
        }
        public List<ElementId> SortX(ICollection<ElementId>toSort,Document doc)
        {
            List<ElementId> xSorted = new List<ElementId>();
            int countElem = toSort.Count;
            foreach (int _ in Enumerable.Range(0, countElem))
            { xSorted.Add(null); }
            foreach (ElementId eid in toSort)
            {
                countElem = toSort.Count;
                foreach (ElementId compEid in toSort)
                {
                    if (eid.IntegerValue != compEid.IntegerValue)
                    {
                        if (RackDim.GetRotatedToVert(eid, doc).GetEndPoint(0).X > RackDim.GetRotatedToVert(compEid, doc).GetEndPoint(0).X) { countElem -= 1; }
                    }
                    
                }
                xSorted[countElem - 1] = eid;
            }
            return xSorted;
        }
        public List<List<ElementId>> SortZ(ICollection<ElementId> toSort, Document doc, double dMax)
        {
            List<ElementId> zAbove = new List<ElementId>();
            List<ElementId> zBelow = new List<ElementId>();
            List<ElementId> zRow = new List<ElementId>();
            foreach (ElementId eid in toSort)
            {
                LocationCurve cLine = doc.GetElement(eid).Location as LocationCurve;
                double checkZ = cLine.Curve.Evaluate(0.5, true).Z;
                bool row = true;
                foreach (ElementId compEid in toSort)
                {
                    if (eid.IntegerValue == compEid.IntegerValue) { continue; }
                    LocationCurve compLine = doc.GetElement(compEid).Location as LocationCurve;
                    double diffZ = checkZ - compLine.Curve.Evaluate(0.5, true).Z;
                    if (Math.Abs(diffZ) > dMax / 2)
                    { if (diffZ > 0) { zAbove.Add(eid); row = false; break; }
                        else { zBelow.Add(eid); row = false; break; }
                    }
                }
                if (row) { zRow.Add(eid); }
            }
            List<List<ElementId>> zSorted = new List<List<ElementId>>
            {
            zAbove, zBelow, zRow
            };
            return zSorted;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UniteDims : IExternalCommand
    {
        //Not working? Is this any good?
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
            Line dimLine = null;
            List<XYZ> dimpos = new List<XYZ>();
            List<Dimension> dimlist = new List<Dimension>();
            ReferenceArray updatedref = new ReferenceArray();
            ReferenceArray cleanref = new ReferenceArray();
            RackDim.GetMenuValues(uiapp);
            foreach (ElementId eid in ids)
            {
                Element elem = doc.GetElement(eid);
                if (elem.GetType() == typeof(Dimension))
                {
                    Dimension dim = elem as Dimension;
                    dimlist.Add(dim);
                    if (dimLine == null) { dimLine = dim.Curve as Line; }
                    foreach (Reference refe in dim.References)
                    {
                        if (doc.GetElement(refe.ElementId).GetType() == typeof(Grid))
                        {
                            Grid grid = doc.GetElement(refe.ElementId) as Grid;
                            Reference refgrid = new Reference(grid);
                            updatedref.Append(refgrid);
                        }
                        else
                        {
                            GeometryElement geom = doc.GetElement(refe.ElementId).get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                            Line refLine = RackDim.GetLineOfGeom(geom);
                            updatedref.Append(refLine.Reference);
                        }
                    }
                }
            }
            List<Dimension> dimsort = dimlist.OrderBy(o=>o.Origin.X).ToList();
            foreach (Dimension dim in dimsort)
            {
                foreach (DimensionSegment dimSeg in dim.Segments)
                {
                    dimpos.Add(dimSeg.LeaderEndPosition);
                }
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Unite Dimensions");
                Dimension newDim = doc.Create.NewDimension(doc.ActiveView, dimLine, updatedref);
                int segCount = 0;
                foreach (DimensionSegment segment in newDim.Segments)
                {
                    bool once = true;
                    foreach (Reference upref in updatedref)
                    {
                        if (once && newDim.References.get_Item(segCount).ElementId == upref.ElementId)
                        { cleanref.Append(upref); once = false; }
                    }
                    if (segment.Value == 0) { segCount += 1; }
                    segCount += 1;
                }
                doc.Delete(newDim.Id);
                newDim = doc.Create.NewDimension(doc.ActiveView, dimLine, cleanref);
                int iter = 0;
                foreach (DimensionSegment seg in newDim.Segments)
                {
                    if (iter < newDim.Segments.Size - 1)
                    {
                        seg.LeaderEndPosition = dimpos[iter];
                        iter += 1;
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LinearAnnotation : IExternalCommand
    {
        //needs grid, and only MEP elements, this way its quite limited
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Line dimDirCross = null; Line dimDir = null;
            int selectdim = 0;
            bool update = false;
            ElementArray dimensions = new ElementArray();
            ReferenceArray dimto = new ReferenceArray();
            RackDim.GetMenuValues(uiapp);
            foreach (ElementId eid in ids)
            {
                if (doc.GetElement(eid).GetType() == typeof(Grid))
                {
                    Grid grid = doc.GetElement(eid) as Grid;
                    dimto.Append(new Reference(grid));
                }
                else if (doc.GetElement(eid).GetType() == typeof(Dimension))
                {
                    update = true;
                    dimensions.Append(doc.GetElement(eid));
                }
                else
                {
                    GeometryElement geom = doc.GetElement(eid).get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                    Line refLine = RackDim.GetLineOfGeom(geom);
                    if (selectdim == 0)
                    {
                        dimDir = Line.CreateBound(refLine.GetEndPoint(0), refLine.GetEndPoint(1));
                        dimDirCross = RackDim.GetPerpendicular(dimDir, StoreAnno.mod_place);
                        selectdim = 1;
                    }
                    dimto.Append(refLine.Reference);
                }
            }
            if (update)
            {
                selectdim = 0;
                foreach (Element elem in dimensions)
                {
                    Dimension dim = elem as Dimension;
                    ReferenceArray updatedref = new ReferenceArray();
                    foreach (Reference refe in dim.References)
                    {
                        if (doc.GetElement(refe.ElementId).GetType() == typeof(Grid))
                        {
                            Grid grid = doc.GetElement(refe.ElementId) as Grid;
                            updatedref.Append(new Reference(grid));
                        }
                        else
                        {
                            GeometryElement geom = doc.GetElement(refe.ElementId).get_Geometry(StoreAnno.Dimop(doc.ActiveView));
                            Line refLine = RackDim.GetLineOfGeom(geom);
                            if (selectdim == 0)
                            {
                                dimDir = Line.CreateBound(refLine.GetEndPoint(0), refLine.GetEndPoint(1));
                                dimDirCross = RackDim.GetPerpendicular(dimDir, StoreAnno.mod_place);
                                selectdim = 1;
                            }
                            updatedref.Append(refLine.Reference);
                        }
                    }
                    RackDim.CreateLinearDim(doc, dimDir, dimDirCross, updatedref);
                    using (Transaction deltrans = new Transaction(doc))
                    {
                        deltrans.Start("Update Dimension");
                        doc.Delete(elem.Id);
                        deltrans.Commit();
                    }
                }
            }
            else { RackDim.CreateLinearDim(doc, dimDir, dimDirCross, dimto); }
            return Result.Succeeded;
        }
    }
}