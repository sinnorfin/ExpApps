/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018,2019,2020,2021 by Attila Kalina <attilakalina.arch@gmail.com>
 *                     and Ildikó Trick <ildiko_trick@trimble.com> (2017)
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
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.IO;

namespace NavisExport
{
    [Transaction(TransactionMode.Manual)]
    public class NavisExport : IExternalCommand
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

            //create NWExportOptions
            NavisworksExportOptions nweOptions = new NavisworksExportOptions();
            nweOptions.ExportScope = NavisworksExportScope.View;
            

            // create a new ViewSet - add views to it that match the desired criteria

            string match = "Export";

            // Use Title on Sheet for select the export views.
            string display = "List of Views:";
            foreach (View vs in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
            {
                if (vs.IsTemplate == false && vs.Category != null && vs.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString().Contains(match))
                {
                    myViewSet.Insert(vs);
                    display = display + Environment.NewLine + vs.Title;
                }
            }

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string path = dlg.FileName;
                path = path.Remove(dlg.FileName.LastIndexOf(@"\")) + "\\NWC";
                TaskDialog td = new TaskDialog("Exporting Navis");
                td.MainInstruction = myViewSet.Size + " Views will be Exported to: " + path;
                td.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.No;
                td.ExpandedContent = display;
                TaskDialogResult response = td.Show();
                if ((response != TaskDialogResult.Cancel) && (response != TaskDialogResult.No))
                {
                    if (!Directory.Exists(path))
                    { Directory.CreateDirectory(path); }

                    foreach (View View in myViewSet)
                    {
                        nweOptions.ViewId = View.Id;
                        doc.Export(path, View.get_Parameter(BuiltInParameter.VIEW_NAME).AsString() + ".nwc", options: nweOptions);
                    }
                }
            }
            return Result.Succeeded;
        }
    }
}
