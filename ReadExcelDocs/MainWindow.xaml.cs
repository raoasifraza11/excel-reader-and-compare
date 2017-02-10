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
using ReadExcelDocs;
namespace ReadExcelDocs
{


    using Excel = Microsoft.Office.Interop.Excel;
    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        Excel.Application excelApp;
        Excel.Workbook excelWorkBook;
        Excel.Worksheet excelWorksheet;
        Excel.Range range;

        public MainWindow()
        {
            InitializeComponent();
            excelApp = new Excel.Application();
        }

        //private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog openFileDialog = new OpenFileDialog();
        //    //if (openFileDialog.ShowDialog() == true)
        //      //  txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
        //}

        private void btnBrowes_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog filebrowes = new OpenFileDialog();

            filebrowes.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (filebrowes.ShowDialog() == true)
            {
                string filepath = filebrowes.FileName;

                string text = "";

                readExcelBook();
                //MessageBox.Show(filename);
            }
        }

        private void readExcelBook()
        {
            excelWorkBook = excelApp.Workbooks.Open(@"C:\Users\Asif Raza\Desktop\excelbook.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelWorksheet = (Excel.Worksheet)excelWorkBook.Sheets[1];
            range = excelWorksheet.UsedRange;
            //MessageBox.Show("@" + path);

        }
    }
}
