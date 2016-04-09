using Microsoft.Win32;
using System;
using System.Collections.Generic;
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

namespace ExcelImporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<ExcelFile> loadedExcel = null;

        public MainWindow()
        {
            InitializeComponent();
            btnStore.IsEnabled = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            OpenFileDialog dialog = new OpenFileDialog();

            string filename = null;

            // Set filter for file extension and default file extension
            dialog.DefaultExt = ".xlsx";
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dialog.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                filename = dialog.FileName;
                txtBoxFileName.Text = filename;
            }

            try
            {
                ReadExcel read = new ReadExcel(filename);
                loadedExcel = read.ExcelReader();

                btnStore.IsEnabled = true;
            }
            catch (Exception)
            {
                
                throw;
            }
        }

        private void btnStore_Click(object sender, RoutedEventArgs e)
        {
            btnStore.IsEnabled = false;

            try
            {
                StoreExcelToDB storeExcel = new StoreExcelToDB();
                listView1.ItemsSource = storeExcel.Store(loadedExcel);
            }
            catch (Exception)
            {

                throw;
            }
        }


    }
}
