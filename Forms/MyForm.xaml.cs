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
        public MyForm(List<Element> _titleblocks)
        {
            InitializeComponent();
            SheetList = new ObservableCollection<SheetDataClass>();
            List<string> tbNames = new List<string>();
            foreach (Element tb in _titleblocks)
            {
                tbNames.Add(tb.Name);
            }
            TitleBlockList = new ObservableCollection<string>() { tbNames.ToString() };
            dataGrid.ItemsSource = SheetList;
            TitleblockCombo.ItemsSource = TitleBlockList;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnAddRow_Click(object sender, RoutedEventArgs e)
        {
            SheetList.Add(new SheetDataClass("", "", false, ""));
        }

        private void btnRemoveRow_Click(object sender, RoutedEventArgs e)
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
    }

    public class SheetDataClass
    {
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public bool PlaceHolder { get; set; }
        public string TitleBlock { get; set; }

        public SheetDataClass(string sheetNumber, string sheetName, bool placeholder, string titleBlock)
        {
            SheetNumber = sheetNumber;
            SheetName = sheetName;
            PlaceHolder = placeholder;
            TitleBlock = titleBlock;
        }
    }
}
