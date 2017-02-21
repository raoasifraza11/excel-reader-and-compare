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

namespace ReadExcelDocs
{
    /// <summary>
    /// Interaction logic for TempProgram.xaml
    /// </summary>
    public partial class TempProgram : Window
    {
        public TempProgram()
        {
            InitializeComponent();
        }

        private void synergybtn_Click(object sender, RoutedEventArgs e)
        {
            synTittle.Text = Core.getFile();
        }

        private void moodlebtn_Click(object sender, RoutedEventArgs e)
        {
            moodleTitle.Text = Core.getFile().ToString();
        }
    }
}
