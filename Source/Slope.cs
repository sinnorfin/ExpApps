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
using Autodesk.Revit.DB.Plumbing;
using System.Linq;

namespace Slope
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ChangeSlope : IExternalCommand
    {
        // creates a pip at the same with the 1/8@12 slope
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
            Options options = new Options();
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            options.View = doc.ActiveView;
            
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Sloped");
                
                    foreach (ElementId eid in ids)
                    {     
                        if (doc.GetElement(eid).GetType() == typeof(Pipe))
                        {
                            Line refline = null;
                            var geoElem = doc.GetElement(eid).get_Geometry(options);
                            foreach (var item in geoElem)
                            {
                                Line lineObj = item as Line;
                                if (lineObj != null)
                                {
                                    refline = lineObj;
                                break;
                                }
                            }
                                var pipeTypes = new FilteredElementCollector(doc)
                                                .OfClass(typeof(PipeType))
                                                .OfType<PipeType>()
                                                .ToList();

                        var firstPipeType =
                                pipeTypes.FirstOrDefault();
                            Pipe newpipe = Pipe.Create(doc, doc.GetElement(eid).get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId(),firstPipeType.Id,
                                doc.GetElement(eid).get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId(),
                                
                                refline.GetEndPoint(0),new XYZ (refline.GetEndPoint(1).X,refline.GetEndPoint(1).Y,refline.GetEndPoint(0).Z+refline.Length*0.0104166666666667 ));}
                    }     
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}