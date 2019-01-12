using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelOpenXmlLibrary
{
    /// <summary>
    /// Taken from http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-3-add-stylesheet-to-the-spreadsheet/
    /// with minor changes by Karen Payne
    /// </summary>
    public class Report
    {

        //public void InsertWorksheet(string docName)
        //{
        //    string excelFileName = "";
        //    string WorksheetName = "";
        //    using (var workbook = SpreadsheetDocument.Open(excelFileName, true))
        //    {
        //        // Add a blank WorksheetPart.
        //        var newWorksheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
        //        var sheetData = new SheetData();
        //        newWorksheetPart.Worksheet = new Worksheet(sheetData);

        //        var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        //        var relationshipId = workbook.WorkbookPart.GetIdOfPart(newWorksheetPart);

        //        // Get a unique ID for the new worksheet.
        //        uint sheetId = 1;
        //        if (sheets.Elements<Sheet>().Any())
        //        {
        //            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        //        }

        //        // Give the new worksheet a name.
        //        var sheetName = WorksheetName;

        //        // Append the new worksheet and associate it with the workbook.
        //        var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        //        sheets.Append(sheet);

        //        Row headerRow = new Row();

        //        // Construct column names 
        //        var columns = new List<string>();
        //        foreach (System.Data.DataColumn column in table.Columns)
        //        {
        //            columns.Add(column.ColumnName);

        //            var cell = new Cell
        //            {
        //                DataType = CellValues.String,
        //                CellValue = new CellValue(column.ColumnName)
        //            };
        //            headerRow.AppendChild(cell);
        //        }

        //        // Add the row values to the excel sheet 
        //        sheetData.AppendChild(headerRow);

        //        foreach (System.Data.DataRow dsrow in table.Rows)
        //        {
        //            Row newRow = new Row();
        //            foreach (var col in columns)
        //            {
        //                var cell = new Cell
        //                {
        //                    DataType = CellValues.String,
        //                    CellValue = new CellValue(dsrow[col].ToString())
        //                };
        //                newRow.AppendChild(cell);
        //            }
        //            sheetData.AppendChild(newRow);
        //        }
        //    }
        //}
        public void CreateExcelDoc(string fileName)
        {
            
            using (var document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Adding style
                var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                // Setting up columns
                var columns = new Columns(
                        new Column // Id column
                        {
                            Min = 1,
                            Max = 1,
                            Width = 4,
                            CustomWidth = true
                        },
                        new Column // Name and Birthday columns
                        {
                            Min = 2,
                            Max = 3,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column // Salary column
                        {
                            Min = 4,
                            Max = 4,
                            Width = 8,
                            CustomWidth = true
                        });

                worksheetPart.Worksheet.AppendChild(columns);

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Employees" };

                // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                var employees = new List<Employee>()
                {
                    new Employee() {Id = 1, Name = "Karen Payne", BirthDateTimeDOB = new System.DateTime(1960,1,14), Salary = 200000M},
                    new Employee() {Id = 2, Name = "Jim Jones", BirthDateTimeDOB = new System.DateTime(1956,12,1), Salary = 90000M},
                    new Employee() {Id = 3, Name = "Mary Adams", BirthDateTimeDOB = new System.DateTime(1989,11,14), Salary = 250000M}
                };

                var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                var row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String, 2),
                    ConstructCell("Name", CellValues.String, 2),
                    ConstructCell("Birth Date", CellValues.String, 2),
                    ConstructCell("Salary", CellValues.String, 2));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var employee in employees)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                        ConstructCell(employee.Name, CellValues.String, 1),
                        ConstructCell(employee.BirthDateTimeDOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                        ConstructCell(employee.Salary.ToString(), CellValues.Number, 1));

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            var fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            var fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            var borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            var cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }

    public class Employee
    {
        public DateTime BirthDateTimeDOB;
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Salary { get; set; }
    }
}
