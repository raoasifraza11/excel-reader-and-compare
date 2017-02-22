using Microsoft.Win32;

namespace ReadExcelDocs
{
    public static class Core
    {
        public static string getFilePath()
        {
            OpenFileDialog getfile = new OpenFileDialog { Filter = "Excel files(*.xlsx;*.xls) | *.xlsx;*.xls" };


            if (getfile.ShowDialog() == true)
            {
                return getfile.FileName;
            }
            else
            {
                return "Please select again ->";
            }
        }

    }
}
