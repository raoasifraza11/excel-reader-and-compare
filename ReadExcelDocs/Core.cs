using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelDocs
{
    public static class Core
    {
        public static string getFile()
        {
            OpenFileDialog getfile = new OpenFileDialog { Filter = "Excel files(*.xlsx) | *.xlsx" };


            if (getfile.ShowDialog() == true)
            {
                return getfile.FileName;
            }
            else
            {
                return null;
            }
        }
    }
}
