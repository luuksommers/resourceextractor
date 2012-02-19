using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Resources;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.SS.UserModel;

namespace ResourceExtractor
{
    public class ResourceExport
    {
        private readonly string _searchDirectory;
        private readonly string _rootCulture;
        private readonly string _outputPath;
        private readonly bool _recursive;

        private const int CultureRowNum = 0;
        private const int NameColumnNum = 0;
        private int _currentColumn = 1;

        public ResourceExport(string searchDirectory, string rootCulture, string outputPath, bool recursive)
        {
            _searchDirectory = searchDirectory;
            _rootCulture = rootCulture;
            _outputPath = outputPath;
            _recursive = recursive;
        }

        public void Export()
        {
            if (Directory.Exists(_searchDirectory))
            {
                var filePaths = Directory.GetFiles(_searchDirectory, "*.resx", _recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
                var workbook = CreateWorkBook();

                foreach (var filePath in filePaths)
                {
                    var fileName = Path.GetFileName(filePath);
                    Console.WriteLine("Processing file {0}...", fileName);
                    
                    var resourceName = GetResourceName(fileName);
                    var culture = GetCulture(fileName);

                    NamesToWorkbook(workbook, filePath, resourceName, culture);
                }

                WriteWorkBook(workbook, _outputPath);
            }
            else
            {
                Console.WriteLine("Directory {0} does not exist.", _searchDirectory);
            }
        }

        private void NamesToWorkbook(HSSFWorkbook workbook, string resourcePath, string sheetName, string culture)
        {
            var sheet = GetSheet(workbook, sheetName);
            using (var resxReader = new ResXResourceReader(resourcePath))
            {
                var cultureRow = GetRow(sheet, CultureRowNum);
                SetCell(cultureRow, _currentColumn, culture);

                foreach (DictionaryEntry reader in resxReader)
                {
                    var exists = false;
                    for (var i = 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = GetRow(sheet, i);
                        if (row.GetCell(NameColumnNum).StringCellValue == reader.Key.ToString())
                        {
                            SetCell(row, _currentColumn, reader.Value.ToString());
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        var nameRow = GetRow(sheet, sheet.LastRowNum + 1);
                        SetCell(nameRow, NameColumnNum, reader.Key.ToString());
                        SetCell(nameRow, _currentColumn, reader.Value.ToString());
                    }
                }
                _currentColumn++;
            }
        }

        private HSSFWorkbook CreateWorkBook()
        {
            var workbook = new HSSFWorkbook();

            var documentSummaryInformation = PropertySetFactory.CreateDocumentSummaryInformation();
            documentSummaryInformation.Company = "Same Problem More Code";
            workbook.DocumentSummaryInformation = documentSummaryInformation;

            var summaryInformation = PropertySetFactory.CreateSummaryInformation();
            summaryInformation.Subject = "Automated resource export";
            workbook.SummaryInformation = summaryInformation;

            return workbook;
        }

        private void WriteWorkBook(HSSFWorkbook workbook, string outputPath)
        {
            using (var fileStream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                workbook.Write(fileStream);
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

                return workbook.CreateSheet(sheetName);
            }

            return null;
        }

        private void SetCell(Row row, int cellNum, string cellValue)
        {
            var cell = row.GetCell(cellNum);
            if (cell == null)
                cell = row.CreateCell(cellNum);

            cell.SetCellValue(cellValue);
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

        private string GetCulture(string resourceName)
        {
            var culture = resourceName.Remove(resourceName.LastIndexOf("."));
            culture = culture.Substring(culture.LastIndexOf(".") + 1);

            if (IsCulture(culture))
                return culture;

            return _rootCulture;
        }

        private bool IsCulture(string culture)
        {
            try
            {
                CultureInfo cultureInfo = new CultureInfo(culture);
                if (string.IsNullOrEmpty(cultureInfo.Name))
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private string GetResourceName(string resourceName)
        {
            resourceName = resourceName.Substring(0, resourceName.IndexOf("."));
            return resourceName;
        }
    }
}
