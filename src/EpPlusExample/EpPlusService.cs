using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace EpPlusExample
{
    public static class EpPlusService
    {
        public static Stream Generate(Stream inputStream)
        {
            var columnHeaders = new string[]
            {
                "Id",
                "Name",
                "Nickname",
                "Age"
            };
            var users = User.GenerateMockData();
            var datas = users.Select(x => new
            {
                x.Id,
                x.Name,
                x.NickName,
                x.Age
            });
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("รายงานยืมคืน"); // Worksheet name
                worksheet.Column(2).Width = 40; // set column width

                using (var cells = worksheet.Cells[1, 1, 1, 4]) // Title content bold
                {
                    cells.Style.Font.Bold = true;
                }

                //First add the headers
                for (var i = 0; i < columnHeaders.Count(); i++)
                {
                    worksheet.Cells[1, i + 1].Value = columnHeaders[i];
                }

                var j = 2; // Start content row
                foreach (var data in datas)
                {
                    worksheet.Cells[("A" + j)].Value = data.Id;
                    worksheet.Cells[("B" + j)].Value = data.Name;
                    worksheet.Cells[("C" + j)].Value = data.NickName;
                    worksheet.Cells[("D" + j)].Value = data.Age;
                    j++;
                }
                // worksheet.Cells.AutoFitColumns();
                package.SaveAs(inputStream);
                return inputStream;
            }
        }
    }

    public sealed class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string NickName { get; set; }
        public int Age { get; set; }

        public static IEnumerable<User> GenerateMockData()
        {
            var rnd = new Random();

            var user = new List<User>();
            foreach (var index in Enumerable.Range(0, 20))
            {
                user.Add(new User
                {
                    Id = index + 1,
                    Name = Guid.NewGuid().ToString(),
                    NickName = Guid.NewGuid().ToString(),
                    Age = rnd.Next(1, 99)
                });
            }
            return user;
        }
    }
}