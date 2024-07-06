/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018,2019,2020,2021 by Attila Kalina <attilakalina.arch@gmail.com>
 *                     and Ildikó Trick <ildiko_trick@trimble.com>
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

using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using System.Linq;

namespace ExportAll
{
    [Transaction(TransactionMode.Manual)]
    public class ExportAll : IExternalCommand
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
            ViewSet myViewSet = new ViewSet();

            //create ExportOptions
            DWGExportOptions DwgOptions = new DWGExportOptions();
            DwgOptions.ExportingAreas = false;
            DwgOptions.MergedViews = true;

            NavisworksExportOptions nweOptions = new NavisworksExportOptions();
            nweOptions.ExportScope = NavisworksExportScope.View;

            // create a new ViewSet - add views to it that match the desired criteria

            string match = "Export";

            // Use Title on Sheet for select the export views

            foreach (View vs in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
            {
                if (vs.IsTemplate == false && vs.Category != null && vs.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString().Contains(match))
                {
                    myViewSet.Insert(vs);
                }
            }

            List<ElementId> viewIds = new List<ElementId>();
            string display = "List of Views:";

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string path = dlg.FileName;
                string rootpath = path.Remove(dlg.FileName.LastIndexOf(@"\"));
                path = path.Remove(dlg.FileName.LastIndexOf(@"\")) + "\\NWC";
                TaskDialog td = new TaskDialog("Exporting DWGs");
                td.MainInstruction = myViewSet.Size + " Views will be Exported to: " + path;
                td.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.No;
                td.ExpandedContent = display;
                TaskDialogResult response = td.Show();
                if ((response != TaskDialogResult.Cancel) && (response != TaskDialogResult.No))
                {
                    
                    if (!Directory.Exists(path))
                    { Directory.CreateDirectory(path); }

                    TaskDialog.Show("Exporting to Navis and DWG", myViewSet.Size + " Views will be Exported to: " + rootpath);
                    // export
                    var ToRename = new List<String>();
                    foreach (View View in myViewSet)
                    {
                        nweOptions.ViewId = View.Id;
                        doc.Export(path, View.get_Parameter(BuiltInParameter.VIEW_NAME).AsString() + ".nwc", options: nweOptions);
                        viewIds.Add(View.Id);
                        display = display + Environment.NewLine + View.Title;
                        String Filename = doc.Title.Replace(" ", "") + "-" + View.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString() + " - " + View.Name + ".dwg";
                        ToRename.Add(Filename);
                    }
                    path = rootpath + "\\DWG";
                    if (!Directory.Exists(path))
                    { Directory.CreateDirectory(path); }
                    doc.Export(path, "", viewIds, DwgOptions);

                    for (int i = 0; i < ToRename.Count; i++)
                    {
                        if (File.Exists(path + "\\" + ToRename[i].Substring(ToRename[i].LastIndexOf(" - ") + 3)))
                        { File.Delete(path + "\\" + ToRename[i].Substring(ToRename[i].LastIndexOf(" - ") + 3)); }
                        File.Move(path + "\\" + ToRename[i], path + "\\" + ToRename[i].Substring(ToRename[i].LastIndexOf(" - ") + 3));
                    }
                }
            } 
            return Result.Succeeded;
        }
    }
}
