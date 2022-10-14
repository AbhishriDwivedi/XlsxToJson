
using System;
using System.Data.Common;
using System.Data.OleDb;
using System.Text;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace nms
{
    class read_from_excel_file
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter the path for xlsx file to be converted:");
            string inputPath = Console.ReadLine();
            if (!File.Exists(inputPath)) { Console.WriteLine("Exiting program due to wrong input"); return; }
            string strFileName = inputPath; //C:\Users\ja0136\Desktop\Attendance.xlsx
            if (!File.Exists(strFileName))
            {
                Console.WriteLine("Invalid Location");
                return;
            }
            Console.WriteLine("Please enter the location where file is to be stored");
            string destination = Console.ReadLine();
            Console.WriteLine("Enter a name for the created file");
            string name = Console.ReadLine();
            string destFileName = Path.Combine(destination, name);
            Application excel = new Application();
            Workbook workBook = excel.Application.Workbooks.Open(strFileName);
            Worksheet worksheet = (Worksheet)workBook.Sheets[1];
            int columns = 1;
            while (worksheet.Cells[1, columns].Value != null)
            {
                columns++;
            }
            columns -= 1;
            int rows = 1;
            while (worksheet.Cells[rows, 1].Value != null)
            {
                rows++;
            }
            rows -= 1;
           JArray array= new JArray();
            JObject o = new JObject();
            for (int i = 2; i <= rows; i++)
            {
                o = new JObject();
                for (int j = 1; j <= columns; j++)
                {   
                    o.Add((worksheet.Cells[1, j].Value).ToString(), worksheet.Cells[i, j].Value);
                }
                array.Add(o);
            }

            string json = o.ToString();
            Console.WriteLine(array);
            try
            {
                var details = JsonConvert.SerializeObject(array);
                File.WriteAllText(destFileName, details.ToString());
                Console.WriteLine("File is created in the requested location");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

       
    }
}
