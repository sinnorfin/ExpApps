﻿/**
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

using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
namespace PrintRevision
{
    public class SelectRev : System.Windows.Forms.Form
    {
        List<Revision> revList = null;
        System.Windows.Forms.ComboBox RevDD = new System.Windows.Forms.ComboBox();
        System.Windows.Forms.Button ok_button = new System.Windows.Forms.Button();
        public bool closed = false;
        public SelectRev(List<Revision> revList)
        {
            this.MaximizeBox = false; this.MinimizeBox = false;
            this.MaximumSize = new System.Drawing.Size(250, 120);
            this.revList = revList;
            this.FormClosed += quit;
           

            ok_button.Location = new System.Drawing.Point(10, 50);
            ok_button.Name = "Ok";
            ok_button.Text = "OK";

            RevDD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            RevDD.Location = new System.Drawing.Point(10, 10);
            RevDD.Name = "Select Revision";
            RevDD.Size = new System.Drawing.Size(180, 20);
            RevDD.TabIndex = 0;
            RevDD.Text = "Available Revisions";
            
            foreach (Revision revis in this.revList)
            { RevDD.Items.Add("Revision " + revis.SequenceNumber); }
            RevDD.SelectedIndex = 0;
            this.Controls.Add(RevDD);
            this.Controls.Add(ok_button);
            ok_button.Click += new EventHandler(ok_Click);
            System.Windows.Forms.Application.Run(this);
        }
        public Revision getRevision()
        {
            Revision rev = this.revList[this.RevDD.SelectedIndex];
            return rev;
        }
        private void ok_Click(object sender, System.EventArgs e)
        {
            this.Close();
            this.closed = false;
        }
        private void quit(object sender, System.EventArgs e)
        {
            this.closed = true ;
        }
    }
    [Transaction(TransactionMode.Manual)]
    public class PrintRevision : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ViewSet FilteredSheets = new ViewSet();

            List <Revision> revList = new List<Revision>();

            foreach ( Revision revis in new FilteredElementCollector(doc).OfClass(typeof(Revision)))
            {            
                revList.Add(revis);
            }

            SelectRev revMenu = new SelectRev(revList);
            Revision targetRev = revMenu.getRevision();
            if (revMenu.closed)
            { return Result.Succeeded; }
            foreach (ViewSheet vs in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>())
            {
                IList<ElementId> Revisions = vs.GetAllRevisionIds();
                foreach (ElementId rev in Revisions)
                {
                    if (rev.ToString() == targetRev.Id.ToString())
                    {
                        FilteredSheets.Insert(vs);
                    }
                }
            }
            PrintManager printManager = doc.PrintManager;
            printManager.PrintRange = PrintRange.Select;
            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;
            viewSheetSetting.CurrentViewSheetSet.Views = FilteredSheets;
            String sheetSetName ="Rev "+targetRev.SequenceNumber;
            using (Transaction t = new Transaction(doc, "Creating PrintSet"))
            {
                t.Start();
                try
                {
                    viewSheetSetting.SaveAs(sheetSetName);
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    foreach (ViewSheetSet viewSheetSetToDelete in (from element
                                                                   in new FilteredElementCollector(doc)
                                                                   .OfClass(typeof(ViewSheetSet)).ToElements()
                                                                   where element.Name.Contains(sheetSetName) && element.IsValidObject
                                                                   select element as ViewSheetSet).ToList().Distinct())
                    {
                        viewSheetSetting.CurrentViewSheetSet = printManager.ViewSheetSetting.InSession;
                        viewSheetSetting.CurrentViewSheetSet = viewSheetSetToDelete as ViewSheetSet;
                        printManager.ViewSheetSetting.Delete();
                    }
                    ViewSheetSetting viewSheetSetting_alt = printManager.ViewSheetSetting;
                    viewSheetSetting_alt.CurrentViewSheetSet.Views = FilteredSheets;
                    viewSheetSetting_alt.SaveAs(sheetSetName);
                }
                t.Commit();
                try
                { if (FilteredSheets.Size != 0)
                    { doc.Print(FilteredSheets); }
                else { TaskDialog.Show("Please select another Revision", "Selected Revision has not been added to any of the sheets."); }
                }
               
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                { TaskDialog.Show("Please select another Revision", "Selected Revision has not been added to any of the sheets."); }
            }
            return Result.Succeeded;
        }
    }
}
