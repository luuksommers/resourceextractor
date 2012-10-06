using System;
using System.IO;
using System.Resources;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ResourceExtractor
{
    public class ResourceImport
    {
        private readonly string _importXls;
        private readonly string _rootCulture;
        private readonly string _outputPath;

        private const int CultureRowNum = 0;
        private const int NameColumnNum = 0;

        public ResourceImport(string importXls, string rootCulture, string outputPath)
        {
            _importXls = importXls;
            _rootCulture = rootCulture;
            _outputPath = outputPath;
        }

        public void Import()
        {
            if (File.Exists(_importXls))
            {
                var fileName = Path.GetFileName(_importXls);
                Console.WriteLine("Opening xls file {0}...", _importXls);

                var workbook = ReadWorkBook(fileName);
                for (int sheetNr = 0; sheetNr < workbook.NumberOfSheets; sheetNr++)
                {
                    WorkbookToResx(workbook, workbook.GetSheetAt(sheetNr).SheetName);
                }
            }
            else
            {
                Console.WriteLine("Import file {0} does not exist.", _importXls);
            }
        }

        private HSSFWorkbook ReadWorkBook(string inputPath)
        {
            using (var fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new HSSFWorkbook(fileStream);
                return workbook;
            }
        }

        private void WorkbookToResx(HSSFWorkbook workbook, string sheetName)
        {
            var sheet = GetSheet(workbook, sheetName);
            var cultureRow = GetRow(sheet, CultureRowNum);
            for (int cultureColumnIndex = 1; cultureColumnIndex < cultureRow.LastCellNum; cultureColumnIndex++)
            {
                string culture = cultureRow.GetCell(cultureColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                var resourceFile = Path.Combine(_outputPath, sheetName + "." + culture + ".resx");
                if(culture == _rootCulture)
                {
                    resourceFile = Path.Combine(_outputPath, sheetName + ".resx");
                }
                using (var resxWriter = new ResXResourceWriter(resourceFile))
                {
                    for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = GetRow(sheet, rowIndex);
                        resxWriter.AddResource(
                            row.GetCell(NameColumnNum, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue, 
                            row.GetCell(cultureColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue);
                    }
                }
            }
        }

        private Sheet GetSheet(HSSFWorkbook workbook, string sheetName)
        {
            if (workbook != null)
            {
                var sheet = workbook.GetSheet(sheetName);
                if (sheet != null)
                {
                    return sheet;
                }
            }

            return null;
        }

        private Row GetRow(Sheet sheet, int? row)
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
