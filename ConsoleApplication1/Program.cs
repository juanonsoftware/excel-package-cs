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
            var newFile = new FileInfo(@"Book1_Data.xlsx");
            var template = new FileInfo(@"Book1.xlsx");

            using (var xlPackage = new ExcelPackage(newFile, template))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Sheet1"];
                ExcelCell cell;
                const int startRow = 3;
                int row = startRow;

                GetUsers().ToList().ForEach(user =>
                                            {
                                                if (row >= startRow + 3)
                                                {
                                                    worksheet.InsertRow(row);
                                                }

                                                worksheet.Cell(row, 1).Value = row.ToString(CultureInfo.InvariantCulture);
                                                worksheet.Cell(row, 2).Value = user.Name;
                                                worksheet.Cell(row, 3).Value = user.Email;
                                                worksheet.Cell(row, 4).Value = user.Value.ToString(CultureInfo.InvariantCulture);

                                                worksheet.Cell(row, 4).Style = (user.Value <= 0)
                                                    ? "Bad"
                                                    : (user.Value >= 1) ? "Good" : string.Empty;

                                                // insert the email address as a hyperlink for the name
                                                string hyperlink = "mailto:" + user.Email;
                                                worksheet.Cell(row, 3).Hyperlink = new Uri(hyperlink, UriKind.Absolute);

                                                row++;
                                            });

                for (int iCol = 1; iCol < 4; iCol++)
                {
                    cell = worksheet.Cell(startRow, iCol);
                    for (int iRow = startRow; iRow <= row; iRow++)
                    {
                        worksheet.Cell(iRow, iCol).StyleID = cell.StyleID;
                    }
                }

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
                           Value = 0.9
                       }, 
                       new UserDto()
                       {
                           Email = "hhoangvan@pentalog.fr",
                           Name = "Huan HOANG VAN",
                           Value = 0.3
                       }, 
                       new UserDto()
                       {
                           Email = "hi@pentalog.fr",
                           Name = "Hi VN",
                           Value = 2.1
                       }, 
                       new UserDto()
                       {
                           Email = "lta@pentalog.fr",
                           Name = "Lta VN",
                           Value = 2.11
                       }, 
                        new UserDto()
                       {
                           Email = "lta@pentalog.fr",
                           Name = "Lta VN",
                           Value = 0
                       }, 
                   };
        }
    }
}
