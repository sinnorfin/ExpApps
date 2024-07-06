/**
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
using System;
using System.Windows.Media.Imaging;

namespace Toggle
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Toggle : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            RibbonPanel inputpanel = null ;
            PushButton toggle_Insulation = null;
            foreach (RibbonPanel panel in uiapp.GetRibbonPanels("Exp. Add-Ins"))
            {
                if (panel.Name == "Re-Elevate")
                { inputpanel = panel; }
            }
            foreach (RibbonItem item in inputpanel.GetItems())
            {
                if (item.Name == "Toggle_Insulation")
                { toggle_Insulation = (PushButton)item; }
            }
            string s = toggle_Insulation.ItemText;
            toggle_Insulation.ItemText = s.Equals("Align to INS") ? "Align to MEP" : "Align to INS";
            //find and switch button and data 
            string IconsPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Icons\\");
            string im_ins = IconsPath + "Button_Ins.png"; string im_nin = IconsPath + "Button_NoIns.png";
            Uri uriImage = new Uri(im_nin);
            BitmapImage Image = new BitmapImage(uriImage);
            if (toggle_Insulation.ItemText == "Align to MEP")
            {
                uriImage = new Uri(im_nin);
                Image = new BitmapImage(uriImage);
                toggle_Insulation.Image = Image;
            }
            else
            {
                uriImage = new Uri(im_ins);
                Image = new BitmapImage(uriImage);
                toggle_Insulation.Image = Image;
            }
            return Result.Succeeded;
        }
    }
}