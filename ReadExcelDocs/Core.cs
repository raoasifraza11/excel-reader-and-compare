using Microsoft.Win32;
using System.IO;

namespace ReadExcelDocs
{
    public static class Core
    {

        public static OpenFileDialog getfile = new OpenFileDialog { Filter = "Excel files(*.xlsx;*.xls) | *.xlsx;*.xls" };
        public static string filepath;


        /// <summary>
        /// Static method for open fileBrowseDialog with excel filters
        /// </summary>
        /// <returns>string</returns>
        public static string getFilePath()
        {

            if (getfile.ShowDialog() == true)
            {
                filepath = getfile.FileName;
                return filepath;
            }
            else
            {
                return "Please select again ->";
            }
        }

        public static string getFileName()
        {
            getFilePath();
            return getfile.SafeFileName;
        }


        public static string getFileExtention()
        {
            return Path.GetExtension(getfile.FileName);
        }

    }
}
