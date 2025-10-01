#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;


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

            List<View> unplacedViews = new List<View>();
            try
            {
                // Get all views in the document
                FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(View));

                // Filter out views that are not valid for placement on sheets
                IEnumerable<View> validViews = viewCollector.Cast<View>().ToList()
                    .Where(v => v.CanBePrinted && !v.IsTemplate);

                // Get all views that are already placed on sheets
                HashSet<ElementId> placedViewIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .SelectMany(sheet => sheet.GetAllPlacedViews())
                    .ToHashSet();

                // Find views that are not placed on sheets
                unplacedViews = validViews.Cast<View>()
                    .Where(v => !placedViewIds.Contains(v.Id))
                    .ToList();

                // Output the names of unplaced views
                // TaskDialog.Show("Unplaced Views", string.Join(Environment.NewLine, unplacedViews.Select(v => v.Name)));
            }
            catch (Exception ex)
            {
                // Handle exceptions and display an error message
                message = $"An error occurred: {ex.Message}";
                return Result.Failed;
            }

            //Excel Reader Licence
            ExcelPackage.License.SetNonCommercialOrganization("JPW");

            // open form
            MyForm currentForm = new MyForm(allTitleblocks, unplacedViews)
            {
                Width = 950,
                Height = 475,
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

                        if (titleblock == null && s.PlaceHolder == false)
                        {
                            TaskDialog.Show("Error", $"Title Block not selected. Sheet {s.SheetNumber} not created.");
                            continue;
                        }

                        if (IsSheetNumberInUse(doc, s.SheetNumber))
                        {
                            TaskDialog.Show("Error", $"Sheet number {s.SheetNumber} is already in use. Sheet not created.");
                            continue;
                        }

                        if (s.PlaceHolder && s.ViewToPlace != "")
                        {
                            TaskDialog.Show("Error", $"Views cannot be created on Placeholder Sheets. Sheet {s.SheetNumber} not created.");
                            continue;
                        }

                        if (titleblock != null && s.PlaceHolder == false)
                        {
                            if (!titleblock.IsActive)
                            {
                                titleblock.Activate();
                                doc.Regenerate();
                            }
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

                            // place view on sheet
                            if (s.ViewToPlace != "")
                            {
                                // place view on sheet
                                View viewToPlace = unplacedViews
                                    .Where(x => x.Name == s.ViewToPlace)
                                    .FirstOrDefault() as View;
                                if (viewToPlace != null)
                                {
                                    try
                                    {
                                        XYZ insertPoint = new XYZ(1, 1, 0);
                                        Viewport.Create(doc, sheet.Id, viewToPlace.Id, insertPoint);
                                    }
                                    catch (Exception)
                                    {
                                        TaskDialog.Show("Error", $"View already placed on a Sheet. Sheet {s.SheetNumber} created without view.");
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                    t.Commit();
                    t.Dispose();
                }
            }
            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

        private bool IsSheetNumberInUse(Document doc, string sheetNumber)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet));

            foreach (ViewSheet sheet in collector)
            {
                if (sheet.SheetNumber.Equals(sheetNumber, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
