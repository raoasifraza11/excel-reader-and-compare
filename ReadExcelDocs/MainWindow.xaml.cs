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
        private string row, col;
        private List<string> data1;
        private List<string> data2;

        public MainWindow()
        {
            InitializeComponent();
            row = "A1";
            col = "A4";
            data1 = new List<string>();
            data2 = new List<string>();
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

        private void excelbtn_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application { Visible = true };

            excelApp.Workbooks.Open(@"E:\first.xlsx");

            Excel.Worksheet workSheet = (Excel.Worksheet) excelApp.ActiveSheet;

            Excel.Range range = workSheet.Range[row,col];

            int rc = range.Rows.Count;
            int cc = range.Columns.Count;

            int rCnt, cCnt;
            string str;
            output.Text = string.Empty;
            for (rCnt = 1; rCnt <= rc; rCnt++)
            {
                for (cCnt = 1; cCnt <= cc; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    output.Text += str + "\n";
                    data1.Add(str);
                }
            }
        }

        private void comparebtn_Click(object sender, RoutedEventArgs e)
        {
            List<string> notfound = new List<string>();
            foreach(var item1 in data1)
            {
                foreach(var item2 in data2)
                {
                    if(item1 != item2)
                    {
                        notfound.Add(item1);
                        break;
                    }
                }
            }

            foreach(var item in notfound)
            {
                compareOutput.Text = item;
            }
        }

        private void second_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application { Visible = true };

            excelApp.Workbooks.Open(@"E:\second.xlsx");

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            Excel.Range range = workSheet.Range[row, col];

            int rc = range.Rows.Count;
            int cc = range.Columns.Count;

            int rCnt, cCnt;
            string str;
            output_2.Text = string.Empty;
            int i;
            for (rCnt = 1, i = 0; rCnt <= rc; rCnt++)
            {
                for (cCnt = 1; cCnt <= cc; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    output_2.Text += str + "\n";
                    data2.Add(str);
                }
            }
        }
    }

    // Account Class
    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }
}
