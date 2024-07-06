#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Linq;
#endregion

namespace DWGExport
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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
                DWGExportOptions nweOptions = new DWGExportOptions();
                nweOptions.ExportingAreas = false;
                nweOptions.MergedViews = true;

                // create a new ViewSet - add views to it that match the desired criteria

                string match = "Export";

                TaskDialog.Show("tool tip", "Use Title on Sheet for select the export views. Match word: " + match);

                foreach (View vs in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
                {
                    if (vs.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString().Contains(match))
                    {
                        myViewSet.Insert(vs);
                    }
                }

                List<ElementId> viewIds = new List<ElementId>();

                //viewIds.Add(myViewSet.Id);

                string path = "";
                // get the current date and time
                DateTime dtnow = DateTime.Now;
                string dt = string.Format("{0:yyyyMMdd}", dtnow);

                if (doc.PathName != "")
                {
                    // use model path + date and time
                    path = Path.GetDirectoryName(doc.PathName) + "\\" + dt;
                }
                else
                {
                    // model has not been saved
                    // use C:\DWG_Export + date and time
                    path = "C:\\DWG_Export\\" + dt;
                }

                // create folder
                Directory.CreateDirectory(path);

                // export
                foreach (View View in myViewSet)
                {
                    viewIds.Add(View.Id);
                    //View.get_Parameter(BuiltInParameter.VIEW_NAME).AsString()
                }
                doc.Export(path, "", viewIds, nweOptions);
                TaskDialog.Show("Export to DWG", "  exported to:\n" + path);

                string filename = doc.Title;
                //TaskDialog.Show("filename: ",filename);
                //TaskDialog.Show("filenpath: ", filepath);
                //string dlg = filepath;
                string[] originalFileName = System.IO.Directory.GetFiles(path, "*.dwg");

                foreach (string e in originalFileName)
                {
                    TaskDialog.Show("dsf", e);
                }

                for (int i = 0; i < originalFileName.Length; i++)
                {
                    string[] splitName = new string[originalFileName.Length];
                    string filepath = System.IO.Path.GetDirectoryName(originalFileName[i]);

                    File.Move(originalFileName[i], filepath + "\\" + originalFileName[i].Substring(originalFileName[i].LastIndexOf(" - ") + 3));
                }

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Transaction Name");
                    tx.Commit();
                }
    
                return Result.Succeeded;

            
        }
    }
    

}
