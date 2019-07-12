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
using Autodesk.Revit.ApplicationServices;
using System.Collections.Generic;
namespace Revision_Editor
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Revision_Editor : IExternalCommand
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
            List<ViewSheet> allSheets = new List<ViewSheet>();
            List<Revision> allRevs = new List<Revision>();
            foreach (ViewSheet sheet in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)))
            { allSheets.Add(sheet); }

            foreach (Revision rev in new FilteredElementCollector(doc).OfClass(typeof(Revision)))
            { allRevs.Add(rev); }
            allRevs.Sort((a, b) => a.SequenceNumber.CompareTo(b.SequenceNumber));
            Revision_Editor_Window Editor = new Revision_Editor_Window(allSheets, allRevs, doc);
            return Result.Succeeded;
        }
    }
    public class Revision_Editor_Window : System.Windows.Forms.Form
    {
        System.Windows.Forms.Panel sheetListPanel = new System.Windows.Forms.Panel();
        System.Windows.Forms.Panel revListPanel = new System.Windows.Forms.Panel();
        public Revision_Editor_Window(List<ViewSheet> sheetList, List<Revision> revList, Document doc)
        {
            if (sheetList != null)
            {
                int sh_y = sheetList.Count * 30;
                int re_y = revList.Count * 30;
                int size_y = sh_y + re_y;
                if (sh_y > 400) { sh_y = 400; }
                if (re_y > 400) { re_y = 400; }
                if (size_y > 880) { size_y = 880; }
                this.MaximizeBox = false; this.MinimizeBox = false;
                this.MaximumSize = new System.Drawing.Size(610, size_y);
                this.MinimumSize = new System.Drawing.Size(610, size_y);
                
                sheetListPanel.Location = new System.Drawing.Point(10, 10);
                sheetListPanel.Size = new System.Drawing.Size(580, sh_y);
                sheetListPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                //this.AutoScroll = true;
                int ypos = 0;
                foreach (ViewSheet sheet in sheetList)
                {
                    System.Windows.Forms.Button button = new System.Windows.Forms.Button();
                    button.Text = sheet.SheetNumber + " - " + sheet.Name;
                    button.Size = new System.Drawing.Size(545, 20);
                    button.Location = new System.Drawing.Point(sheetListPanel.Location.X, 10 + ypos);
                    button.BackColor = System.Drawing.SystemColors.ControlLight;
                    button.Click += new System.EventHandler(ok_Click);
                    ypos += 26;
                    sheetListPanel.Controls.Add(button);
                }
                sheetListPanel.AutoScroll = true;
                this.Controls.Add(sheetListPanel);

                revListPanel.Location = new System.Drawing.Point(10, sh_y + 26);
                revListPanel.Size = new System.Drawing.Size(580, re_y);
                revListPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                ypos = 0;
                foreach (Revision rev in revList)
                {
                    System.Windows.Forms.Label label = new System.Windows.Forms.Label();
                    label.Text = rev.RevisionDate + " - " + rev.Description;
                    label.Size = new System.Drawing.Size(545, 20);
                    label.Location = new System.Drawing.Point(revListPanel.Location.X, 10 + ypos);
                    label.BackColor = System.Drawing.SystemColors.ControlLight;
                    ypos += 26;
                    revListPanel.Controls.Add(label);
                }
                revListPanel.AutoScroll = true;
                this.Controls.Add(revListPanel);
            }
            ShowDialog();

            void ok_Click(object sender, System.EventArgs e)
            {
                System.Windows.Forms.Button button = (System.Windows.Forms.Button)sender;
                foreach (System.Windows.Forms.Label label in revListPanel.Controls)
                {
                    label.BackColor = System.Drawing.SystemColors.HighlightText;
                }
                    foreach (ViewSheet sheet in sheetList)
                {
                    if (button.Text.Substring(0).Contains(sheet.SheetNumber))
                    {
                        IList<ElementId> sheetrev = sheet.GetAllRevisionIds();
                        foreach (ElementId eid in sheetrev)
                        {
                            Revision rev = doc.GetElement(eid) as Revision;
                            foreach (System.Windows.Forms.Label label in revListPanel.Controls)
                            {
                                if (label.Text == rev.RevisionDate + " - " + rev.Description)
                                {
                                    label.BackColor = System.Drawing.SystemColors.Highlight;
                                }
                            }
                        }
                    }
                }
            }

        }
    }
}
