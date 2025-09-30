using Autodesk.Revit.DB;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace RAA_Sheet_Maker
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyForm : Window
    {
        ObservableCollection<SheetDataClass> SheetList { get; set; }
        ObservableCollection<string> TitleBlockList { get; set; }
        ObservableCollection<string> ViewList { get; set; }

        String txtFilePath = "";
        String saveFilePath = "";
        public MyForm(List<Element> _titleblocks, List<View> _validViews)
        {
            InitializeComponent();
            SheetList = new ObservableCollection<SheetDataClass>();
            TitleBlockList = new ObservableCollection<string>();
            foreach (Element tb in _titleblocks)
            {
                TitleBlockList.Add(tb.Name.ToString());
            }
            ViewList = new ObservableCollection<string>();
            foreach (View v in _validViews)
            {
                ViewList.Add(v.Name.ToString());
            }
            dataGrid.ItemsSource = SheetList;
            TitleblockCombo.ItemsSource = TitleBlockList;
            ViewPlaceCombo.ItemsSource = ViewList;
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
        {
            SheetList.Add(new SheetDataClass("", "", false, "", ""));
        }

        private void BtnRemoveRow_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (SheetDataClass curRow in SheetList)
                {
                    if (dataGrid.SelectedItem == curRow)
                        SheetList.Remove(curRow);
                }
            }
            catch (Exception)
            { }
        }

        public List<SheetDataClass> GetData()
        {
            List<SheetDataClass> returnData = new List<SheetDataClass>();
            foreach (SheetDataClass row in SheetList)
            {
                if (row.SheetNumber == null) row.SheetNumber = "";
                if (row.SheetName == null) row.SheetName = "";
                if (row.TitleBlock == null) row.TitleBlock = "";
                if (row.ViewToPlace == null) row.ViewToPlace = "";
                // row.PlaceHolder is a bool, no need to check for null
                returnData.Add(row);
            }
            return returnData;
        }

        private void ExportXLS_Click(object sender, RoutedEventArgs e)
        {
            // Export to XLS functionality to be implemented
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.RestoreDirectory = true;
            saveFile.Filter = "Excel files (*.xlsx) |*.xlsx";
            saveFile.Title = "Save as Excel file";
            if (saveFile.ShowDialog() == true)
            {
                saveFilePath = saveFile.FileName;
            }
            else
            {
                saveFilePath = "";
            }
            using (var package = new ExcelPackage(saveFilePath))
            {   var worksheet = package.Workbook.Worksheets.Add("Sheets");
                // Add headers
                worksheet.Cells[1, 1].Value = "Sheet Number";
                worksheet.Cells[1, 2].Value = "Sheet Name";
                worksheet.Cells[1, 3].Value = "Place Holder";
                worksheet.Cells[1, 4].Value = "Title Block";
                worksheet.Cells[1, 5].Value = "View to Place";
                // Add data
                for (int i = 0; i < SheetList.Count; i++)
                {
                    var sheet = SheetList[i];
                    worksheet.Cells[i + 2, 1].Value = sheet.SheetNumber;
                    worksheet.Cells[i + 2, 2].Value = sheet.SheetName;
                    worksheet.Cells[i + 2, 3].Value = sheet.PlaceHolder;
                    worksheet.Cells[i + 2, 4].Value = sheet.TitleBlock;
                    worksheet.Cells[i + 2, 5].Value = sheet.ViewToPlace;
                }
                // Save the package
                package.Save();
            }
        }

        private void ImportXLS_Click(object sender, RoutedEventArgs e)
        {
            // Import from XLS functionality to be implemented
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Multiselect = false;
            //openFile.InitialDirectory = @"C:\"; // or use "C:\\"
            //openFile.Filter = "csv files (*.csv) |*.csv";
            //openFile.Title = "Select a CSV file";
            openFile.RestoreDirectory = true;
            openFile.Filter = "Excel files (*.xlsx;*.xls) |*.xlsx;*.xls";
            openFile.Title = "Select an Excel file";

            if (openFile.ShowDialog() == true)
            {
                txtFilePath = openFile.FileName;
            }
            else
            {
                txtFilePath = "";
            }
            using (var package = new ExcelPackage(new System.IO.FileInfo(txtFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                // Clear existing data
                SheetList.Clear();
                // Read data starting from row 2 to skip headers
                for (int row = 2; row <= rowCount; row++)
                {
                    string sheetNumber = worksheet.Cells[row, 1].Text;
                    string sheetName = worksheet.Cells[row, 2].Text;
                    bool placeHolder = false;
                    bool.TryParse(worksheet.Cells[row, 3].Text, out placeHolder);
                    string titleBlock = worksheet.Cells[row, 4].Text;
                    string viewToPlace = worksheet.Cells[row, 5].Text;
                    SheetList.Add(new SheetDataClass(sheetNumber, sheetName, placeHolder, titleBlock, viewToPlace));
                }
            }
        }
    }

    public class SheetDataClass
    {
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public bool PlaceHolder { get; set; }
        public string TitleBlock { get; set; }
        public string ViewToPlace { get; set; }

        public SheetDataClass(string sheetNumber, string sheetName, bool placeholder, string titleBlock, string viewtoPlace)
        {
            SheetNumber = sheetNumber;
            SheetName = sheetName;
            PlaceHolder = placeholder;
            TitleBlock = titleBlock;
            ViewToPlace = viewtoPlace;
        }
    }
}
