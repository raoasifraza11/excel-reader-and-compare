using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadExcelDocs
{
    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Create a list of accounts.
        Account data = new Account();
        

        public MainWindow()
        {
            InitializeComponent();
        }

        

        private void btnBrowes_Click(object sender, RoutedEventArgs e)
        {
            data.ID = 3456;
            data.Balance = 546.24;

            var excelApp = new Excel.Application();
            excelApp.Visible = true;


            excelApp.Workbooks.Add();

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "4435";
            workSheet.Cells[1, "B"] = "453";
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();


        }

      
    }

    // Account Class
    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }
}
