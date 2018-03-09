using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;


namespace ExcelTool.Dev
{
    class Program
    {
        static void Main(string[] args)
        {
            bool exit = false;
            Console.WriteLine("***********************************");
            Console.WriteLine("* Excel Commandline Modifier Tool *");
            Console.WriteLine("***********************************");
            Console.Write("Enter the path to the Excel File of interest: ");
            String filePath = Console.ReadLine();
            FileInfo file = new FileInfo(filePath);

            while (exit == false) {
                Console.WriteLine();
                Console.WriteLine("     ****************");
                Console.WriteLine("     * MENU OPTIONS *");
                Console.WriteLine("     ****************");
                Console.WriteLine(" Enter 1 to replace Column values");
                Console.WriteLine(" Enter 2 to remove a Column");
                String choice = Console.ReadLine();

                if(choice == "1")
                {
                    ReadWorkbook(file);
                }
                else if(choice == "2")
                {
                    Console.Write("Enter the column you wish to remove: ");
                    String column = Console.ReadLine();
                    int _col = GetColumn(column.ToUpper());
                    RemoveColumn(file, _col);
                }
                else
                {
                    Console.WriteLine("INVALID INPUT!");
                    Console.WriteLine("Would you like to continue? Y/N");
                    String again = Console.ReadLine();
                    if(again == "N" || again == "n")
                    {
                        exit = true;
                        Console.WriteLine("Later! ");
                    }
                }
            }
        }

        public static void ReadWorkbook(FileInfo _file)
        {
            using (ExcelPackage package = new ExcelPackage(_file)) { 
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                Console.WriteLine("Enter the Column ID for the column you wish to edit: ");
                String column = Console.ReadLine();
                int col = GetColumn(column.ToUpper());
                Console.WriteLine("Enter a new value: ");
                String newValue = Console.ReadLine();
                for(int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    worksheet.Cells[row, col].Value = newValue; 
                }
                Console.WriteLine();
                package.Save();
            }
        }

        public static void RemoveColumn(FileInfo _file, int column)
        {
            using (ExcelPackage package = new ExcelPackage(_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                worksheet.DeleteColumn(column);
                package.Save();
            }
        }

        public static int GetColumn(string colID)
        {
            IDictionary<string, int> columnDictionary = new Dictionary<string, int>()
            {
                {"A",1},{"B",2},{"C",3},{"D",4},{"E",5},{"F",6},{"G", 7},{"H",8},{"I",9},{"J",10},
                {"K",11}, {"L",12},{"M",13},{"N",14},{"O", 15 },{"P", 16 },{"Q", 17 },{"R", 18 },{"S", 19 },
                {"T", 20},{"U", 21},{"V", 22},{"W", 23},{"X", 24},{"Y", 25},{"Z", 26},{"AA", 27},{"AB", 28},{"AC", 29},{"AD", 30}
            };
            return columnDictionary[colID];
        }


    }
}
