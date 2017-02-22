using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
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
            // Intilize the Variables
            synData = new List<string>();
            moodleData = new List<string>();
            compareOneToTwo = new List<string>();
            compareTwoToOne = new List<string>();

            // Login username
            username.Content = Environment.UserName;
        }


        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            synTittle.Text = Core.getFilePath().ToString();

        }

        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            moodleTitle.Text = Core.getFilePath().ToString();
          
        }

        private void viewResult_Click(object sender, RoutedEventArgs e)
        {
            Result r = new Result();
            r.Show();
            r.result1Listview.ItemsSource = synData;
            r.synCount.Content = synData.Count.ToString();

            r.moodleListView.ItemsSource = moodleData;
            r.moodleCount.Content = moodleData.Count.ToString();

            r.compareResult1to2.ItemsSource = compareOneToTwo;
            r.oneToTwocount.Content = compareOneToTwo.Count.ToString();

            r.compareResult2to1.ItemsSource = compareTwoToOne;
            r.twoToOnecount.Content = compareTwoToOne.Count.ToString();
        }


        private void OpenExcelApplication(string path, List<string> data, string filename)
        {
            try
            {
                var excelApp = new Excel.Application { Visible = false };
                excelApp.Workbooks.Open(path);
                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                //string row = synRangeRow.Text.ToString();
                //string col = synRangeCol.Text.ToString();
                //string[] SelectedRange = { synRangeRow.Text.ToString(), synRangeCol.Text.ToString() };

                if(filename == "synTittle")
                {
                    string row = synRangeRow.Text.ToString();
                    string col = synRangeCol.Text.ToString();

                    Excel.Range range = workSheet.Range[row, col];
                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        for (int j = 1; j <= range.Columns.Count; j++)
                        {
                            //data.Add((string)((range.Cells[i, j]
                            //as Excel.Range).Value2).ToString().Substring(((range.Cells[i, j]
                            //as Excel.Range).Value2).ToString().Length - 4, 4));

                            string value = ((string)(((range.Cells[i, j]
                            as Excel.Range).Value2).ToString().Substring(((range.Cells[i, j]
                            as Excel.Range).Value2).ToString().Length - 4, 4)));

                            value = value.TrimStart('0');
                            data.Add(value);
                        }
                    }
                }
                else
                {

                    string row = moodleRangeRow.Text.ToString();
                    string col = moodleRangeCol.Text.ToString();

                    Excel.Range range = workSheet.Range[row, col];
                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        for (int j = 1; j <= range.Columns.Count; j++)
                        {
                            data.Add((string)((range.Cells[i, j]
                            as Excel.Range).Value2).ToString());
                        }
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
            //reading data from snergy
            OpenExcelApplication(synTittle.Text, synData, synTittle.Name);
            //reading data from moodles
            OpenExcelApplication(moodleTitle.Text, moodleData, moodleTitle.Name);



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
        //Reseting lists used   & textboxex
        private void resetBtn_Click(object sender, RoutedEventArgs e)
        {
            synData.Clear();
            moodleData.Clear();
            compareOneToTwo.Clear(); 
            compareTwoToOne.Clear();
            synTittle.Text = "FileName";
            moodleTitle.Text="FileName";


        }
    }
}
