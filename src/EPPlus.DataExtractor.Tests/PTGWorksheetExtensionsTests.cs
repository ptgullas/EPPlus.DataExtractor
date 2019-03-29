using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace EPPlus.DataExtractor.Tests {
    public class PTGWorksheetExtensionsTests {

        public class Member {
            // REQUIRED:
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string ZipCode { get; set; }
            // OPTIONAL BUT VERY HELPFUL:
            public string Address { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Phone { get; set; }
            public string Email { get; set; }
            public DateTime DateOfBirth { get; set; }
            public string MiddleName { get; set; }
            public string NameSuffix { get; set; }

        }

        private Stream GetSpreadsheetFileInfo() =>
            GetType().Assembly.GetManifestResourceStream(GetType(), "spreadsheets.PTGTestWorkbook.xlsx");


        [Fact]
        public void ExtractSimpleData() {
            // var fileInfo = GetSpreadsheetFileInfo();
            string filePath = @"C:\Users\Prime Time Pauly G\Source\Repos\EPPlus.DataExtractor\src\EPPlus.DataExtractor.Tests\spreadsheets\PTGTestWorkbook.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            int expectedCount = 4;

            List<Member> members = new List<Member>();

            using (var package = new ExcelPackage(fileInfo)) {
                var sheet = package.Workbook.Worksheets[0];
                members = sheet
                    .Extract<Member>()
                    .WithProperty(p => p.LastName, "A")
                    .WithProperty(p => p.FirstName, "B")
                    .WithProperty(p => p.ZipCode, "F")
                    .GetData(2, sheet.Dimension.Rows)
                    .ToList();

            }

            Assert.Equal(4, members.Count);
            Assert.Contains(members,
                m => m.FirstName == "Aegon" && m.LastName == "Targaryen" && m.ZipCode == "10003");
            Assert.Contains(members,
                m => m.FirstName == "Arianne" && m.LastName == "Martell" && m.ZipCode == "10025");
        }

        [Fact]
        public void ExtractSimpleData_NullColumn() {
            // var fileInfo = GetSpreadsheetFileInfo();
            string filePath = @"C:\Users\Prime Time Pauly G\Source\Repos\EPPlus.DataExtractor\src\EPPlus.DataExtractor.Tests\spreadsheets\PTGTestWorkbook.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            int expectedCount = 4;

            List<Member> members = new List<Member>();

            using (var package = new ExcelPackage(fileInfo)) {
                var sheet = package.Workbook.Worksheets[0];
                members = sheet
                    .Extract<Member>()
                    .WithProperty(p => p.LastName, "A")
                    .WithProperty(p => p.FirstName, "B")
                    .WithProperty(p => p.ZipCode, "F")
                    .WithOptionalProperty(p => p.Address, null)
                    .GetData(2, sheet.Dimension.Rows)
                    .ToList();

            }

            Assert.Equal(expectedCount, members.Count);
            Assert.Contains(members,
                m => m.FirstName == "Aegon" && m.LastName == "Targaryen" && m.ZipCode == "10003");
            Assert.Contains(members,
                m => m.FirstName == "Arianne" && m.LastName == "Martell" && m.ZipCode == "10025");
            Assert.Null(members[0].Address);
        }

    }
}
