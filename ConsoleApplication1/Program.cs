using System.Globalization;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo newFile = new FileInfo(@"Book1_Data.xlsx");
            FileInfo template = new FileInfo(@"Book1.xlsx");

            using (ExcelPackage xlPackage = new ExcelPackage(newFile, template))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Sheet1"];
                ExcelCell cell;
                const int startRow = 3;
                int row = startRow;

                GetUsers().ToList().ForEach(user =>
                                            {
                                                if (row >= startRow)
                                                {
                                                    worksheet.InsertRow(row);
                                                }

                                                worksheet.Cell(row, 1).Value = row.ToString(CultureInfo.InvariantCulture);
                                                worksheet.Cell(row, 2).Value = user.Name;
                                                worksheet.Cell(row, 3).Value = user.Email;
                                                worksheet.Cell(row, 4).Value = user.Value.ToString(CultureInfo.InvariantCulture);

                                                // insert the email address as a hyperlink for the name
                                                string hyperlink = "mailto:" + user.Email;
                                                worksheet.Cell(row, 3).Hyperlink = new Uri(hyperlink, UriKind.Absolute);

                                                row++;
                                            });

                xlPackage.Save();
            }
        }

        static IList<UserDto> GetUsers()
        {
            return new UserDto[]
                   {
                       new UserDto()
                       {
                           Email = "huanhvhd@gmail.com",
                           Name = "HOANG VAN HUAN",
                           Value = 1.1
                       }, 
                       new UserDto()
                       {
                           Email = "hhoangvan@pentalog.fr",
                           Name = "Huan HOANG VAN",
                           Value = 2.3
                       }, 
                       new UserDto()
                       {
                           Email = "hi@pentalog.fr",
                           Name = "Hi VN",
                           Value = 2.1
                       }, 
                   };
        }
    }
}
