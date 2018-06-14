﻿/**
 * Experimental Apps - Add-in For AutoDesk Revit
 *
 *  Copyright 2017,2018 by Attila Kalina <attilakalina.arch@gmail.com>
 *
 * This file is part of Experimental Apps.
 * Exp Apps has been developed from June 2017 until end of March 2018 under the encouragement and for the use of hungarian BackOffice of Trimble VDC Services.
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
using System;
using _ExpApps;

namespace QuickViews
{
    public class AssignViews : System.Windows.Forms.Form
    {
        List<View> viewList = null;
        List<System.Windows.Forms.ComboBox> dropdowns = AssignViews.createdropdown();

        System.Windows.Forms.Button ok_button = new System.Windows.Forms.Button();
        public static List<System.Windows.Forms.Button> createmenu()
        {
            List<System.Windows.Forms.Button> menu = new List<System.Windows.Forms.Button>();
            for (int i = 1; i <= 6; i++)
            { menu.Add(new System.Windows.Forms.Button()); }
            return menu;
        }
        public static List<System.Windows.Forms.ComboBox> createdropdown()
        {
            List<System.Windows.Forms.ComboBox> menu = new List<System.Windows.Forms.ComboBox>();
            for (int i = 1; i <= 6; i++)
            { menu.Add(new System.Windows.Forms.ComboBox()); }
            return menu;
        }
        public AssignViews(List<View> viewList)
        {
            if (viewList != null)
            {
                this.MaximizeBox = false; this.MinimizeBox = false;
                this.MaximumSize = new System.Drawing.Size(610, 260);
                this.MinimumSize = new System.Drawing.Size(610, 260);
                this.viewList = viewList;

                ok_button.Location = new System.Drawing.Point(10, 180);
                ok_button.Text = "Save Quick Keys";
                int ypos = 10;
                int ind = 0;
                foreach (System.Windows.Forms.ComboBox combo in dropdowns)
                {
                    combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
                    combo.Location = new System.Drawing.Point(10, ypos);
                    combo.Size = new System.Drawing.Size(280, 20);
                    combo.TabIndex = 0;
                    combo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
                    combo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;

                    foreach (View view in this.viewList)
                    { combo.Items.Add(view.Title);}
                    this.Controls.Add(combo);
                    System.Windows.Forms.TextBox tb = new System.Windows.Forms.TextBox();
                    
                    tb.Location = new System.Drawing.Point(300, ypos);
                    tb.Size = new System.Drawing.Size(280, 20);
                    if (StoreExp.quickViews[ind] != null) {
                        try
                        {
                            tb.Text = "("+ (ind+1).ToString() + ") " + StoreExp.quickViews[ind].ViewName;
                            tb.ReadOnly = true;
                            this.Controls.Add(tb);
                        }
                        catch { TaskDialog.Show("View Lost","View lost due to renaming or deletion");
                            tb.Text ="("+ (ind+1).ToString()+") "+ ":Lost:";
                            tb.ReadOnly = true;
                            this.Controls.Add(tb);
                        }
                    }
                    ypos += 26;
                    ind += 1;
                }
                this.Controls.Add(ok_button);
                ok_button.Click += new EventHandler(ok_Click);
                this.ShowDialog();
            }
        }
        public void setView(int qv_ind, int c_ind)
        {
           if ((dropdowns[c_ind].SelectedIndex >= 0) && (this.viewList[dropdowns[c_ind].SelectedIndex].Category.Name == "Views"))
            {
                StoreExp.quickViews[qv_ind] = this.viewList[dropdowns[c_ind].SelectedIndex]; 
            }
        }
        private void ok_Click(object sender, System.EventArgs e)
        {
            for (int i = 0; i <= 5; i++)
            { setView(i, i); }
            this.Close();
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickViews : IExternalCommand
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
            List<View> allViews = new List<View>();
            foreach (View view in new FilteredElementCollector(doc).OfClass(typeof(View)))
            {if ((view.IsTemplate == false) && (view.Category != null))
                { allViews.Add(view);}
            }
            AssignViews assign = new AssignViews(allViews); 
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView1 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[0] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[0]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[0].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[0];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView2 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[1] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[1]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[1].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[1];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView3 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[2] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[2]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[2].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[2];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView4 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[3] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[3]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[3].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[3];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView5 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[4] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[4]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[4].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[4];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickView6 : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            if (StoreExp.quickViews[5] != null)
            {
                try { uidoc.ActiveView = StoreExp.quickViews[5]; }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    uiapp.OpenAndActivateDocument(StoreExp.quickViews[5].Document.PathName);
                    uidoc = uiapp.ActiveUIDocument;
                    uidoc.ActiveView = StoreExp.quickViews[5];
                }
            }
            else { TaskDialog.Show("Error", "Quick key has not been assigned"); }
            return Result.Succeeded;
        }
    }
}