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
using Autodesk.Revit.DB;


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
        }

        private void ImportXLS_Click(object sender, RoutedEventArgs e)
        {
            // Import from XLS functionality to be implemented
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
