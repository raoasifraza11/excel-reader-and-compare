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

        public MainWin()
        {
            InitializeComponent();
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
                        //data1.Add(str);
                        MessageBox.Show(str);
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

        }

        private void comparebtn_Click(object sender, RoutedEventArgs e)
        {

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


    }
}
