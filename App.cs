#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Windows.Markup;

#endregion

namespace RAA_Sheet_Maker
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            // 1. Create ribbon tab
            try
            {
                app.CreateRibbonTab("JPW Utilities");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists.");
            }

            // 2. Create ribbon panel 
            RibbonPanel panel = Utils.CreateRibbonPanel(app, "JPW Utilities", "Utils");

            // 3. Create button data instances
            ButtonDataClass myButtonData = new ButtonDataClass("btnRAA_Sheet_Maker", "Sheet Maker", Command.GetMethod(), Properties.Resources.Red_32, Properties.Resources.Red_16, "Create new Sheets and specify Title Blocks.");

            // 4. Create buttons
            PushButton myButton = panel.AddItem(myButtonData.Data) as PushButton;
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }


    }
}
