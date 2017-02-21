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
            row = "D2";
            col = "D50";
            data1 = new List<string>();
            data2 = new List<string>();



        }

        

        private void btnBrowes_Click(object sender, RoutedEventArgs e)
        {
           data.ID = 3456;
            data.Balance = 546.24;

            var excelApp = new Excel.Application();
            excelApp.Visible = false;

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

            excelApp.Workbooks.Open(@"E:\ee101synergy.xls");

            Excel.Worksheet workSheet = (Excel.Worksheet) excelApp.ActiveSheet;

            Excel.Range range = workSheet.Range[row,col];

            int rc = range.Rows.Count;
            int cc = range.Columns.Count;
            
            int rCnt, cCnt;

            //

            //string a = "ABCD-0448";
            //string required = a.Substring(5);

            string str;
            //output.Text = string.Empty;
            for (rCnt = 1; rCnt <= rc; rCnt++)
            {
                for (cCnt = 1; cCnt <= cc; cCnt++)
                {
                    str = (string)((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToString().Substring(15);
                    //output.Text += str + "\n";
                    data1.Add(str);
                }
            }
            listView.ItemsSource = data1;
            // excelApp.Workbooks.Close();
            //excelApp.Application.Quit();
        }

        private void comparebtn_Click(object sender, RoutedEventArgs e)
        {
            List<string> notfound = new List<string>();
            foreach(var item1 in data1)
            {
                if (!(data2.Contains(item1)))
                {
                    notfound.Add(item1);
                }
            }

            listView2.ItemsSource = notfound;

        }

        private void second_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application { Visible = false };

            excelApp.Workbooks.Open(@"E:\EE101spring2017_Attendances_20170131-0006.xlsx");

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            row = "B5";
            col = "B53";
            Excel.Range range = workSheet.Range[row, col];

            int rc = range.Rows.Count;
            int cc = range.Columns.Count;

            int rCnt, cCnt;
            string str;
            //output_2.Text = string.Empty;
            // int i; no need more
            for (rCnt = 1; /* i = 0 */ rCnt <= rc; rCnt++)
            {
                for (cCnt = 1; cCnt <= cc; cCnt++)
                {
                    str = (string)((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToString();
                    //output_2.Text += str + "\n";
                    data2.Add(str);
                }
            }

            listView1.ItemsSource = data2;
            // don't need to close workbook
            //excelApp.Workbooks.Close();

            excelApp.Application.Quit();
        }
    }

    // Account Class
    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }
}
