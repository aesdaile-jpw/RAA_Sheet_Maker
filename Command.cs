#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using ExcelDataReader;
using System.Linq;


#endregion

namespace RAA_Sheet_Maker
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

            // put any code needed for the form here
            List<Element> allTitleblocks = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            allTitleblocks = collector.ToElements().Cast<Element>().ToList();



            // open form
            MyForm currentForm = new MyForm(allTitleblocks)
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            currentForm.ShowDialog();

            // get form data and do something

            List<SheetDataClass> sheetsToCreate = new List<SheetDataClass>();

            if (currentForm.DialogResult == true)
            {
               sheetsToCreate = currentForm.GetData();
            }

            if (sheetsToCreate.Count == 0)
            {
                TaskDialog.Show("Info", "No sheets to create.");
                return Result.Cancelled;
            }
            else
            {
                using (Transaction t = new Transaction(doc, "Create Sheets"))
                {
                    t.Start();
                    foreach (SheetDataClass s in sheetsToCreate)
                    {
                        // get titleblock
                        FamilySymbol titleblock = allTitleblocks
                            .Where(x => x.Name == s.TitleBlock)
                            .FirstOrDefault() as FamilySymbol;
                        if (titleblock == null)
                        {
                            TaskDialog.Show("Error", $"Titleblock {s.TitleBlock} not found. Sheet {s.SheetNumber} not created.");
                            continue;
                        }
                        if (!titleblock.IsActive)
                        {
                            titleblock.Activate();
                            doc.Regenerate();
                        }
                        // create sheet
                        if (s.PlaceHolder == true)
                        {
                            ViewSheet placeholderSheet = ViewSheet.CreatePlaceholder(doc);
                            placeholderSheet.SheetNumber = s.SheetNumber;
                            placeholderSheet.Name = s.SheetName;
                        }
                        else
                        {
                            ViewSheet sheet = ViewSheet.Create(doc, titleblock.Id);
                            if (sheet == null)
                            {
                                TaskDialog.Show("Error", $"Failed to create sheet {s.SheetNumber}.");
                                continue;
                            }
                            sheet.SheetNumber = s.SheetNumber;
                            sheet.Name = s.SheetName;
                        }

                        // set sheet number and name

                        //// set issue parameters
                        //if (s.Issue)
                        //{
                        //    Parameter issueDateParam = sheet.LookupParameter("Issue Date");
                        //    if (issueDateParam != null && issueDateParam.IsModifiable)
                        //    {
                        //        issueDateParam.Set(DateTime.Now.ToShortDateString());
                        //    }
                        //    Parameter issuedParam = sheet.LookupParameter("Issued");
                        //    if (issuedParam != null && issuedParam.IsModifiable)
                        //    {
                        //        issuedParam.Set(1); // yes/no parameter: 1 = yes, 0 = no
                        //    }
                        //    Parameter issuedByParam = sheet.LookupParameter("Issued By");
                        //    if (issuedByParam != null && issuedByParam.IsModifiable)
                        //    {
                        //        issuedByParam.Set(Environment.UserName);
                        //    }
                        //}
                    }
                    t.Commit();
                }
            }

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }


    }
}
