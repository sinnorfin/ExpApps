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
                        tag = doc.Create.NewTag(doc.ActiveView, elem, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, refline.Curve.Evaluate(0.5, true));
                 
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
}
