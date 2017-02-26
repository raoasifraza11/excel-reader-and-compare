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

        // Result windows
        Result ResultWindow;
        string moodlesrc;
        string syncsrc;

        /// <summary>
        /// Constructor
        /// Intialize the Object
        /// </summary>
        public TempProgram()
        {
            InitializeComponent();
            // Intilize the Variables
            synData = moodleData = compareOneToTwo = compareTwoToOne = new List<string>();
            
            //preventing click
            viewResult.IsEnabled = resetBtn.IsEnabled = comparebtn.IsEnabled = false;

            // Login username
            username.Content = Environment.UserName;
            //prevent edditing in input boxex
            synTittle.IsEnabled = moodleTitle.IsEnabled = false;

            //preventing muddles button click
            moodlebtn.IsEnabled = false;
            
        }

        /// <summary>
        /// Open FileDialog for file selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            string filename = Core.getFileName();
            string ext = filename.Substring(filename.Length -3, 3);

            if(ext == "xls")
            {
                syncsrc = Core.filepath;
                synTittle.Text = filename;
            }else
            {
                MessageBox.Show("You choose the worng file.");
            }
            moodlebtn.IsEnabled = true;
            synergybtn.IsEnabled = false;

        }

        /// <summary>
        /// Open FileDialog for file selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            string filename = Core.getFileName();
            string ext = filename.Substring(filename.Length - 4, 4);

            if (ext == "xlsx")
            {
                moodlesrc = Core.filepath;
                moodleTitle.Text = filename;
            }
            else
            {
                MessageBox.Show("You choose the worng file.");
            }
            comparebtn.IsEnabled = true;
            moodlebtn.IsEnabled = false;
        }

        /// <summary>
        /// View result activity
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void viewResult_Click(object sender, RoutedEventArgs e)
        {
            // launch result activity
            ResultWindow = new Result();
            ResultWindow.Show();

           
            // populate the data
            ResultWindow.result1Listview.ItemsSource = synData;
            ResultWindow.synCount.Content = synData.Count.ToString();

            ResultWindow.moodleListView.ItemsSource = moodleData;
            ResultWindow.moodleCount.Content = moodleData.Count.ToString();

            // update source
            ResultWindow.compareResult1to2.ItemsSource = compareOneToTwo;
            ResultWindow.oneToTwocount.Content = compareOneToTwo.Count.ToString();

            ResultWindow.compareResult2to1.ItemsSource = compareTwoToOne;
            ResultWindow.twoToOnecount.Content = compareTwoToOne.Count.ToString();
            //prevent click 
            viewResult.IsEnabled = false;
            //allowing click
            resetBtn.IsEnabled = true;

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
            OpenExcelApplication(syncsrc, synData, synTittle.Name);
            //reading data from moodles
            OpenExcelApplication(moodlesrc, moodleData, moodleTitle.Name);
            //sorting lists
            synData.Sort();
            moodleData.Sort();
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

            // prevent the repeated clicks
            comparebtn.IsEnabled = false;
            viewResult.IsEnabled = true;

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

            // 
            if(ResultWindow != null)
            {
                ResultWindow.Close();
            }

            // rest original state
            comparebtn.IsEnabled = true;
            //hiddin rest button
            resetBtn.IsEnabled = false;
            synergybtn.IsEnabled = true;
            moodlebtn.IsEnabled = true;
        }
    }
}
