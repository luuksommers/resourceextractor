using System;
using System.IO;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ResourceExtractor
{
    public class NpoiHelper
    {
        public static HSSFWorkbook CreateWorkBook()
        {
            var workbook = new HSSFWorkbook();

            var documentSummaryInformation = PropertySetFactory.CreateDocumentSummaryInformation();
            documentSummaryInformation.Company = "Same Problem More Code";
            workbook.DocumentSummaryInformation = documentSummaryInformation;

            var summaryInformation = PropertySetFactory.CreateSummaryInformation();
            summaryInformation.Subject = "ResourceExtractor";
            workbook.SummaryInformation = summaryInformation;

            return workbook;
        }

        public static HSSFWorkbook ReadWorkBook(string inputPath)
        {
            using (var fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new HSSFWorkbook(fileStream);
                return workbook;
            }
        }

        public static void WriteWorkBook(HSSFWorkbook workbook, string outputPath)
        {
            using (var fileStream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                workbook.Write(fileStream);
            }
        }

        public static Sheet GetSheet(HSSFWorkbook workbook, string sheetName)
        {
            if (workbook != null)
            {
                var sheet = workbook.GetSheet(sheetName);
                if (sheet != null)
                {
                    return sheet;
                }

                return workbook.CreateSheet(sheetName);
            }

            return null;
        }

        public static void SetCell(Row row, int cellNum, string cellValue)
        {
            var cell = row.GetCell(cellNum);
            if (cell == null)
                cell = row.CreateCell(cellNum);

            cell.SetCellValue(cellValue);
        }

        public static Row GetRow(Sheet sheet, int? row)
        {
            if (sheet != null)
            {
                if (row.HasValue)
                {
                    var existingRow = sheet.GetRow(row.Value);
                    if (existingRow != null)
                    {
                        return existingRow;
                    }

                    return sheet.CreateRow(row.Value);
                }

                return sheet.CreateRow(sheet.LastRowNum + 1);
            }

            return null;
        }
    }
}
