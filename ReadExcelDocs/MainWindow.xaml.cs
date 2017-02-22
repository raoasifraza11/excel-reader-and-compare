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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelDocs
{
    /// <summary>
    /// Interaction logic for TempProgram.xaml
    /// </summary>
    public partial class TempProgram : Window
    {
        // Required variables
        //private string synRow, synCol, modRow, modCol;
        private List<string> synData;
        private List<string> moodleData;
        List<string> compareOneToTwo;
        List<string> compareTwoToOne;

        public TempProgram()
        {
            InitializeComponent();
            // Intilize local variables or Entities
            //synRow = synCol = modRow = modCol = null;
            synData = new List<string>();
            moodleData = new List<string>();
            compareOneToTwo = new List<string>();
            compareTwoToOne = new List<string>();
        }

        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            synTittle.Text = Core.getFilePath().ToString();


            OpenExcelApplication(synTittle.Text, synData);

        }

        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            moodleTitle.Text = Core.getFilePath().ToString();
            OpenExcelApplication(moodleTitle.Text, moodleData);
        }

        private void viewResult_Click(object sender, RoutedEventArgs e)
        {
            Result r = new Result();
            r.Show();
            r.result1Listview.ItemsSource = synData;
            r.moodleListView.ItemsSource = moodleData;

            r.compareResult1to2.ItemsSource = compareOneToTwo;
            r.compareResult2to1.ItemsSource = compareTwoToOne;
        }


        private void OpenExcelApplication(string path, List<string> data)
        {
            try
            {
                var excelApp = new Excel.Application { Visible = false };
                excelApp.Workbooks.Open(path);
                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                string row = synRangeRow.Text.ToString();
                string col = synRangeCol.Text.ToString();
                //string[] SelectedRange = { synRangeRow.Text.ToString(), synRangeCol.Text.ToString() };

                Excel.Range range = workSheet.Range[row, col];
                for (int i = 1; i <= range.Rows.Count; i++)
                {
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        data.Add((string)((range.Cells[i, j]
                        as Excel.Range).Value2).ToString());
                    }
                }

                excelApp.Workbooks.Close();
                excelApp.Application.Quit();
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK);
            }
            
        }

        private void compare_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in synData)
            {
                if (!(moodleData.Contains(item)))
                {
                    compareOneToTwo.Add(item);
                }
            }

            foreach (var item in moodleData)
            {
                if (!(synData.Contains(item)))
                {
                    compareTwoToOne.Add(item);
                }
            }

        }
    }
}
