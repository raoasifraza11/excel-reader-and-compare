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
        private string row, col;
        public MainWin()
        {
            InitializeComponent();
            row = "D2";
            col = "D50";
        }

        private void syngbtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                run_excel_application();
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
                var excelApp = new Excel.Application { Visible = false };

                excelApp.Workbooks.Open(@"E:\ee101synergy.xls");

                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                Excel.Range range = workSheet.Range[row, col];
            }catch(Exception ex)
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
