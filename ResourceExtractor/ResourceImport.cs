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

                var workbook = NpoiHelper.ReadWorkBook(fileName);
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



        private void WorkbookToResx(HSSFWorkbook workbook, string sheetName)
        {
            var sheet = NpoiHelper.GetSheet(workbook, sheetName);
            var cultureRow = NpoiHelper.GetRow(sheet, CultureRowNum);
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
                        var row = NpoiHelper.GetRow(sheet, rowIndex);
                        var valueRow = row.GetCell(cultureColumnIndex);
                        if (valueRow != null)
                        {
                            resxWriter.AddResource(
                                row.GetCell(NameColumnNum, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue,
                                valueRow.StringCellValue);
                        }
                    }
                }
            }
        }
    }
}
