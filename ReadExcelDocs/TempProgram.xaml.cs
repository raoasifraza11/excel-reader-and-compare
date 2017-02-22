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

        public TempProgram()
        {
            InitializeComponent();
            // Intilize local variables or Entities
            //synRow = synCol = modRow = modCol = null;
            synData = new List<string>();
        }

        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            synTittle.Text = Core.getFile();

            var excelApp = new Excel.Application { Visible = true };
            excelApp.Workbooks.Open(@"E:\EE101spring2017_Attendances_20170131-0006.xlsx");
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            string row = synRangeRow.Text.ToString();
            string col = synRangeCol.Text.ToString();
            //string[] SelectedRange = { synRangeRow.Text.ToString(), synRangeCol.Text.ToString() };

            Excel.Range range = workSheet.Range[row, col];
            for (int i = 1; i <= range.Rows.Count; i++)
            {
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    synData.Add((string)((range.Cells[range.Row, range.Column]
                    as Excel.Range).Value2).ToString());
                }
            }

            excelApp.Workbooks.Close();
            excelApp.Application.Quit();
        }

        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            moodleTitle.Text = Core.getFile().ToString();
        }
    }
}
