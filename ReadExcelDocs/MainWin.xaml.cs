using System.Windows;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;

namespace ReadExcelDocs
{
    /// <summary>
    /// Interaction logic for MainWin.xaml
    /// </summary>
    public partial class MainWin : Window
    {
        //private Excel.Application excelApp;
        private List<string> data1;
        private List<string> data2;

        public MainWin()
        {
            InitializeComponent();
            data1 = new List<string>();
            data2 = new List<string>();
            listView.Visibility = Visibility.Hidden;
            //excelApp = new Excel.Application { Visible = false };
        }

        private void syngbtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var excelApp = new Excel.Application { Visible = true };

                excelApp.Workbooks.Open(@"E:\ee101synergy.xls");

                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                Excel.Range range = workSheet.Range[synRow.Text.ToString(), synCol.Text.ToString()];
                int rc = range.Rows.Count;
                int cc = range.Columns.Count;

                int rCnt, cCnt;
                
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
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Reading file", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void moddlebtn_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application { Visible = false };

            excelApp.Workbooks.Open(@"E:\EE101spring2017_Attendances_20170131-0006.xlsx");

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            Excel.Range range = workSheet.Range[moodRow.Text.ToString(), moodCol.Text.ToString()];

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

        }

        private void comparebtn_Click(object sender, RoutedEventArgs e)
        {
            List<string> notfound = new List<string>();
            foreach (var item1 in data1)
            {
                if (!(data2.Contains(item1)))
                {
                    notfound.Add(item1);
                }
            }
            listView.ItemsSource = notfound;
        }


        /*
         * open excel applicaiton to given path;
         */ 
        private void run_excel_application()
        {
            try
            {
                
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private string file_browser()
        {
            OpenFileDialog getfile = new OpenFileDialog { Filter = "Excel files(*.xlsx) | *.xlsx" };
            

            if(getfile.ShowDialog() == true)
            {
                return getfile.FileName;   
            }else
            {
                return null;
            }
        }

        private void showresultbtn_Click(object sender, RoutedEventArgs e)
        {
            listView.Visibility = Visibility.Visible;
        }
    }
}
