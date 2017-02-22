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
        private List<string> synData;
        private List<string> moodleData;
        List<string> compareOneToTwo;
        List<string> compareTwoToOne;

        /// <summary>
        /// Constructor
        /// Intialize the Object
        /// </summary>
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

        /// <summary>
        /// Open FileDialog for file selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            synTittle.Text = Core.getFilePath().ToString();

        }

        /// <summary>
        /// Open FileDialog for file selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            moodleTitle.Text = Core.getFilePath().ToString();

        }

        /// <summary>
        /// View result activity
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void viewResult_Click(object sender, RoutedEventArgs e)
        {
            // launch result activity
            Result r = new Result();
            r.Show();

            // populate the data
            r.result1Listview.ItemsSource = synData;
            r.synCount.Content = synData.Count.ToString();

            r.moodleListView.ItemsSource = moodleData;
            r.moodleCount.Content = moodleData.Count.ToString();

            // update source
            r.compareResult1to2.ItemsSource = compareOneToTwo;
            r.oneToTwocount.Content = compareOneToTwo.Count.ToString();

            r.compareResult2to1.ItemsSource = compareTwoToOne;
            r.twoToOnecount.Content = compareTwoToOne.Count.ToString();
        }

        /// <summary>
        /// Launch excel applicaiton and read data form given
        /// range and update it into given lists.
        /// </summary>
        /// <param name="path">Path of the file</param>
        /// <param name="data">Temprary stored values in List</param>
        /// <param name="filename">File Name for select login based on file name </param>
        private void OpenExcelApplication(string path, List<string> data, string filename)
        {
            try
            {
                // Instantiate excel object
                var excelApp = new Excel.Application { Visible = false };

                // Open workbook
                excelApp.Workbooks.Open(path);

                // select active worksheet
                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                // select login based on filename
                if (filename == "synTittle")
                {
                    // select rows and cols form fields
                    string row = synRangeRow.Text.ToString();
                    string col = synRangeCol.Text.ToString();

                    Excel.Range range = workSheet.Range[row, col];
                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        for (int j = 1; j <= range.Columns.Count; j++)
                        {
                            // read value form cell and make substring then select last 4 digits
                            // as per requirement
                            string value = ((string)(((range.Cells[i, j]
                            as Excel.Range).Value2).ToString().Substring(((range.Cells[i, j]
                            as Excel.Range).Value2).ToString().Length - 4, 4)));

                            // Remove 0 form start
                            value = value.TrimStart('0');

                            // generate list
                            data.Add(value);
                        }
                    }
                }
                else
                {
                    // if not synergy generate file
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
            }
            // Handle exception if any
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK);
            }

        }

        /// <summary>
        /// Compare two files dependent on OpenExcelApplication method
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void compare_Click(object sender, RoutedEventArgs e)
        {
            //reading data from snergy
            OpenExcelApplication(synTittle.Text, synData, synTittle.Name);
            //reading data from moodles
            OpenExcelApplication(moodleTitle.Text, moodleData, moodleTitle.Name);


            // Compare lists form 1 To 2
            foreach (var item in synData)
            {
                if (!(moodleData.Contains(item)))
                {
                    compareOneToTwo.Add(item);
                }
            }

            // Comparer lists form 2 to 1
            foreach (var item in moodleData)
            {
                if (!(synData.Contains(item)))
                {
                    compareTwoToOne.Add(item);
                }
            }

        }

        /// <summary>
        /// Reset the all Lits and fiels
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void resetBtn_Click(object sender, RoutedEventArgs e)
        {
            synData.Clear();
            moodleData.Clear();
            compareOneToTwo.Clear();
            compareTwoToOne.Clear();
            synTittle.Text = moodleTitle.Text = "FileName";
            
        }
    }
}
