
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace nms
{
    class read_from_excel_file
    {
        public class Root
        {
            public string Name { get; set; }
            public double Age { get; set; }
            public string Gender { get; set; }
            public string Profession { get; set; }
            public string Attendance { get; set; }
        }
        static void Main(string[] args)
        {
            
            Console.WriteLine("Please enter the path for xlsx file to be converted:");
            string inputPath = Console.ReadLine();
            if (!File.Exists(inputPath)) { Console.WriteLine("Exiting program due to wrong input"); return; }
            string strFileName = inputPath; //D:\People.xlsx
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
            Workbooks workBooks = excel.Workbooks;
            Workbook workBook = workBooks.Open(strFileName);
            Sheets sheets = workBook.Sheets;
            Worksheet worksheet = sheets[1];

            var UR = worksheet.UsedRange;

            var usedRange = UR.Value;
            int rows = usedRange.Length/5;
            var obj = new Root();
            List<Root> list = new List<Root>();
            for (int i = 2; i <= rows; i++)
            {
                obj = new Root()
                {
                    Name = usedRange[i, 1],
                    Age = usedRange[i, 2],
                    Gender = usedRange[i, 3],
                    Profession = usedRange[i, 4],
                    Attendance = usedRange[i, 5]
                };
                list.Add(obj);
            }

            try
            {
                string details = JsonConvert.SerializeObject(list);
                Console.Write(details);
                File.WriteAllText(destFileName, details.ToString());
                Console.WriteLine("File is created in the requested location");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            Marshal.FinalReleaseComObject(UR);
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(sheets);
            workBook.Close();
            Marshal.FinalReleaseComObject(workBook);
            workBooks.Close();
            Marshal.FinalReleaseComObject(workBooks);
            excel.Quit();
            Marshal.FinalReleaseComObject(excel);
        }
    }
    }


