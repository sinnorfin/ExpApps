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
using System.Linq;
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
            if (Autodesk.Revit.DB.RevisionSettings.GetRevisionSettings(doc).RevisionNumbering.ToString() == "PerProject")
            { TaskDialog.Show("Revision Manager", " Revision numbering is Per Project, Application designed for Per Sheet option.");
                return Result.Succeeded;
            }
            List<ViewSheet> allSheets = new List<ViewSheet>();
            List<Revision> allRevs = new List<Revision>();
            //Get All Sheets and Revisions
            foreach (ViewSheet sheet in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)))
            { allSheets.Add(sheet); }
            allSheets.Sort((a, b) => a.SheetNumber.CompareTo(b.SheetNumber));
            foreach (Revision rev in new FilteredElementCollector(doc).OfClass(typeof(Revision)))
            { allRevs.Add(rev); }
            //Sort revisions by sequencenumber
            allRevs.Sort((a, b) => a.SequenceNumber.CompareTo(b.SequenceNumber));
            //Create window
            Revision_Editor_Window Editor = new Revision_Editor_Window(allSheets, allRevs, doc);
            return Result.Succeeded;
        }
    }
    public class Revision_Editor_Window : System.Windows.Forms.Form
    {
        System.Windows.Forms.Panel sheetListPanel = new System.Windows.Forms.Panel();
        System.Windows.Forms.Panel revListPanel = new System.Windows.Forms.Panel();
        List<ViewSheet> selectedSheets = new List<ViewSheet>();
        bool selectedRev = false;
        List<System.Windows.Forms.CheckBox> sheetListButtons = new List<System.Windows.Forms.CheckBox>();
        List<System.Windows.Forms.Label> revListButtons = new List<System.Windows.Forms.Label>();
        List<System.Windows.Forms.Button> selButtons = new List<System.Windows.Forms.Button>();
        List<System.Windows.Forms.Button> addrmvButtons = new List<System.Windows.Forms.Button>();

        static IList<ElementId> check_getrev_Ids( List<ViewSheet> sheetList, string text)
        {
            IList<ElementId> sheetrev = new List<ElementId>();
            foreach (ViewSheet sheet in sheetList)
            {
                if (text.Substring(0, sheet.SheetNumber.Length) == sheet.SheetNumber)
                {
                    sheetrev = sheet.GetAllRevisionIds();
                }
            }
            return sheetrev;
        }
        static ViewSheet getSheet_bySheetNum(List<ViewSheet> sheetlist, string text)
        {
            foreach (ViewSheet sheet in sheetlist)
            {
                if (text.Substring(0, sheet.SheetNumber.Length) == sheet.SheetNumber)
                { return sheet; }
            }
            return null;
        }
        static List<ViewSheet> getSheets_byRev(List<ViewSheet> sheetlist, string text,Document doc)
        {
            List<ViewSheet> sheetswithRev = new List<ViewSheet>();
            foreach (ViewSheet sheet in sheetlist)
            {
                ICollection<ElementId> revsonsheet = sheet.GetAllRevisionIds();
                foreach (ElementId eid in revsonsheet)
                {
                    Revision rev = doc.GetElement(eid) as Revision;
                    if (rev.SequenceNumber.ToString() == text.Split(' ')[0])
                    {
                        sheetswithRev.Add(sheet);
                    }
                }
            }
            return sheetswithRev;
        }
        static Revision get_rev(List<Revision> revList, string text)
        {
            foreach (Revision rev in revList)
            {
                if (text.Split(' ')[0] == rev.SequenceNumber.ToString())
                { return rev; }
            }
            return null;
        }
        static void Update_rev_highlight(Revision rev, List<System.Windows.Forms.Label> revListButtons, bool on)
        {
            System.Drawing.Color highlight = System.Drawing.SystemColors.ControlLight;
            if (on) { highlight = System.Drawing.SystemColors.ActiveCaption; }
            foreach (System.Windows.Forms.Label label in revListButtons)
            {
                if (label.Text == rev.SequenceNumber + " : " + rev.RevisionDate + " - " + rev.Description)
                {
                    label.BackColor = highlight;
                }
            }
        }
        static void Update_sheet_highlight(List<ViewSheet> sheets, List<System.Windows.Forms.CheckBox> sheetListButtons)
        {
            System.Drawing.Color highlight = System.Drawing.SystemColors.ActiveCaption;
            foreach (System.Windows.Forms.CheckBox button in sheetListButtons)
            {
                foreach (ViewSheet sheet in sheets)
                {
                   
                    if (button.Text.Contains(sheet.SheetNumber.ToString()))
                    {
                        button.BackColor = highlight;
                    }
                }
            }
        }
        static void resetLabelColor(List<System.Windows.Forms.Label> controls)
        {
            foreach (System.Windows.Forms.Label label in controls)
            {
                label.BackColor = System.Drawing.SystemColors.ControlLight;
            }
        }
        static void resetBoxColor(List<System.Windows.Forms.CheckBox> controls)
        {
            foreach (System.Windows.Forms.CheckBox button in controls)
            {
                button.BackColor = DefaultBackColor;
            }
        }
        static void resetButtonColor(List<System.Windows.Forms.Button> controls, bool hl)
        {
            System.Drawing.Color color = DefaultBackColor;
            if (hl) { color = System.Drawing.SystemColors.Info; }
            foreach (System.Windows.Forms.Button button in controls)
            { if (!hl)
                { button.BackColor = color; }
                else if(button.BackColor == System.Drawing.SystemColors.ActiveCaption)
                { button.BackColor = color; }
            }
        }
        public Revision_Editor_Window(List<ViewSheet> sheetList, List<Revision> revList, Document doc)
        {
            if (sheetList != null)
            {
                //Adjust Size to content
                this.Text = "ExpApps - Revision Editor";
                int sh_y = sheetList.Count * 26 + 15 ;
                int re_y = revList.Count * 26 + 15;
                int size_y = System.Math.Max(sh_y, re_y) +100;

                if (sh_y > 700) { sh_y = 720; }
                if (re_y > 700) { re_y = 720; }
                if (size_y > 800) { size_y = 800; }
                this.MaximizeBox = false; this.MinimizeBox = false;
                this.MaximumSize = new System.Drawing.Size(1285, size_y);
                this.MinimumSize = new System.Drawing.Size(1285, size_y);
                //Tooltip
                System.Windows.Forms.ToolTip tooltips = new System.Windows.Forms.ToolTip();
                sheetListPanel.Location = new System.Drawing.Point(10, 30);
                sheetListPanel.Size = new System.Drawing.Size(570, sh_y);
                sheetListPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                int ypos = 0;
                //Create buttons for each sheet
                foreach (ViewSheet sheet in sheetList)
                {
                    System.Windows.Forms.CheckBox sheetbutton = new System.Windows.Forms.CheckBox
                    {
                        Appearance = System.Windows.Forms.Appearance.Button,
                        Text = sheet.SheetNumber + " - " + sheet.Name,
                        Size = new System.Drawing.Size(545, 20),
                        Location = new System.Drawing.Point(10, 10 + ypos),
                        BackColor = DefaultBackColor
                    };
                    sheetbutton.Click += new System.EventHandler(sel_Click);
                    tooltips.SetToolTip(sheetbutton, "Show Revisions on this Sheet."+ System.Environment.NewLine + "( Shift + click to Edit or Add more Sheets ) ");
                    ypos += 26;
                    sheetListPanel.Controls.Add(sheetbutton);
                    sheetListButtons.Add(sheetbutton);
                }
                //Add sheet buttons
                sheetListPanel.AutoScroll = true;
                this.Controls.Add(sheetListPanel);

                revListPanel.Location = new System.Drawing.Point(590, 30);
                revListPanel.Size = new System.Drawing.Size(670, re_y);
                revListPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                ypos = 0;
                //Create and add Buttons for Revisions
                foreach (Revision rev in revList)
                {
                    //Title
                    System.Windows.Forms.Label label = new System.Windows.Forms.Label
                    {
                        Text = rev.SequenceNumber + " : " + rev.RevisionDate + " - " + rev.Description,
                        Size = new System.Drawing.Size(495, 20),
                        Location = new System.Drawing.Point(55, 10 + ypos),
                        BackColor = System.Drawing.SystemColors.ControlLight
                    };
                    //'Add' Button
                    System.Windows.Forms.Button add = new System.Windows.Forms.Button
                    {
                        Text = "+",
                        Size = new System.Drawing.Size(40, 20),
                        Location = new System.Drawing.Point(605, 10 + ypos)
                    };
                    add.Click += new System.EventHandler((sender, e) => add_Click(sender, e, label.Text));
                    add.Enabled = false;
                    tooltips.SetToolTip(add, "Add on selected Sheets.");
                    //'Rmv' Button
                    System.Windows.Forms.Button rmv = new System.Windows.Forms.Button
                    {
                        Text = "-",
                        Size = new System.Drawing.Size(40, 20),
                        Location = new System.Drawing.Point(560, 10 + ypos)
                    };
                    rmv.Click += new System.EventHandler((sender, e) => rmv_Click(sender, e, label.Text));
                    rmv.Enabled = false;
                    tooltips.SetToolTip(rmv, "Remove from selected Sheets." +System.Environment.NewLine + "( Revision Cloud prevents removal )");
                    //'Sel' Button
                    System.Windows.Forms.Button sel = new System.Windows.Forms.Button
                    {
                        Text = "o",
                        Size = new System.Drawing.Size(40, 20),
                        Location = new System.Drawing.Point(10, 10 + ypos)
                    };
                    sel.Click += new System.EventHandler((sender, e) => selsheets_Click(sender, e, label.Text));
                    tooltips.SetToolTip(sel, "Select all Sheets using Revision." + System.Environment.NewLine + 
                                            "( Shift + click to select Sheets that use all the selected Revisions ) ");
                    ypos += 26;
                    revListPanel.Controls.Add(label);
                    revListButtons.Add(label);
                    revListPanel.Controls.Add(add);
                    revListPanel.Controls.Add(rmv);
                    revListPanel.Controls.Add(sel);
                    selButtons.Add(sel);
                    addrmvButtons.Add(add); addrmvButtons.Add(rmv);
                }
                revListPanel.AutoScroll = true;
                this.Controls.Add(revListPanel);
                //Labels
                System.Windows.Forms.Label SheetTitle = new System.Windows.Forms.Label
                {
                    Text = "Sheets in document",
                    Size = new System.Drawing.Size(200, 20),
                    Location = new System.Drawing.Point(sheetListPanel.Location.X, 5),
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };
                this.Controls.Add(SheetTitle);

                System.Windows.Forms.Label RevTitle = new System.Windows.Forms.Label
                {
                    Text = "Revisions in document",
                    Size = new System.Drawing.Size(200, 20),
                    Location = new System.Drawing.Point(revListPanel.Location.X, 5),
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };
                this.Controls.Add(RevTitle);
            }
            ShowDialog();
            void switch_Buttons(List<System.Windows.Forms.Button> buttons,bool enable)
            {
                foreach ( System.Windows.Forms.Button button in buttons)
                { button.Enabled = enable; }
            }
            void add_Click(object sender, System.EventArgs e, string rev)
            {
                if (selectedSheets.Count > 0)
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        Revision revision = get_rev(revList, rev);
                        tx.Start("Add Revs:" + revision.Description);
                        foreach (ViewSheet sheet in selectedSheets)
                        {
                            ICollection<ElementId> currentrevs = sheet.GetAllRevisionIds();
                            currentrevs.Add(revision.Id);
                            sheet.SetAdditionalRevisionIds(currentrevs);
                        }
                        tx.Commit();
                        Update_rev_highlight(get_rev(revList, rev), revListButtons,true);
                    }
                }
            }
            void rmv_Click(object sender, System.EventArgs e, string rev)
            {
                if (selectedSheets.Count > 0)
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        Revision revision = get_rev(revList, rev);
                        tx.Start("Rm. Revs:" + revision.Description);
                        foreach (ViewSheet sheet in selectedSheets)
                        {
                            ICollection<ElementId> currentrevs = sheet.GetAllRevisionIds();
                            currentrevs.Remove(revision.Id);
                            sheet.SetAdditionalRevisionIds(currentrevs);
                        }
                        tx.Commit();
                        Update_rev_highlight(get_rev(revList, rev), revListButtons, false);
                    }
                }
            }
            void selsheets_Click(object sender, System.EventArgs e, string rev)
            {
                System.Windows.Forms.Button togglebutton = (System.Windows.Forms.Button)sender;
                if (System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Shift && selectedRev)
                {
                    List<ViewSheet> addtoSelected = getSheets_byRev(sheetList, rev, doc);
                    selectedSheets = selectedSheets.Intersect(addtoSelected).ToList();
                    togglebutton.BackColor = System.Drawing.SystemColors.ActiveCaption;
                }
                else
                {
                    resetButtonColor(selButtons,false);
                    selectedSheets = getSheets_byRev(sheetList, rev, doc);
                    selectedRev = true;
                    togglebutton.BackColor = System.Drawing.SystemColors.ActiveCaption;
                }
                resetLabelColor(revListButtons);
                Update_rev_highlight(get_rev(revList, rev), revListButtons, true);
                resetBoxColor(sheetListButtons);
                Update_sheet_highlight(selectedSheets, sheetListButtons);

                if (selectedSheets.Count > 0) { switch_Buttons(addrmvButtons, true); }
                else { switch_Buttons(addrmvButtons, false); }
            }
            void sel_Click(object sender, System.EventArgs e)
            {
                System.Windows.Forms.CheckBox togglebutton = (System.Windows.Forms.CheckBox)sender;
                if (System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Shift)
                {
                    if (togglebutton.Checked)
                    {
                        selectedSheets.Add(getSheet_bySheetNum(sheetList, togglebutton.Text));
                        togglebutton.BackColor = System.Drawing.SystemColors.ActiveCaption;
                    }
                    else
                    {
                        selectedSheets.Remove(getSheet_bySheetNum(sheetList, togglebutton.Text));
                        togglebutton.BackColor = DefaultBackColor;
                    }
                    if (selectedSheets.Count > 0) { switch_Buttons(addrmvButtons, true); }
                    else { switch_Buttons(addrmvButtons, false); }
                }
                else
                {
                    switch_Buttons(addrmvButtons, false);
                    togglebutton.Checked = false;
                    selectedSheets = new List<ViewSheet>();
                    resetBoxColor(sheetListButtons);
                }
                resetLabelColor(revListButtons);
                resetButtonColor(selButtons,true);
                selectedRev = false;
                foreach (ElementId eid in check_getrev_Ids(sheetList, togglebutton.Text))
                {
                    Revision rev = doc.GetElement(eid) as Revision;
                    foreach (System.Windows.Forms.Label label in revListButtons)
                    {
                        if (label.Text ==rev.SequenceNumber + " : " + rev.RevisionDate + " - " + rev.Description)
                        {
                            label.BackColor = System.Drawing.SystemColors.ActiveCaption;
                        }
                    }
                }
            }
        }
    }
}
