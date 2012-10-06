using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Resources;
using NPOI.HSSF.UserModel;

namespace ResourceExtractor
{
    public class ResourceExport
    {
        private readonly string _searchDirectory;
        private readonly string _rootCulture;
        private readonly string _outputPath;

        private const int CultureRowNum = 0;
        private const int NameColumnNum = 0;
        private int _currentColumn = 1;

        public ResourceExport(string searchDirectory, string rootCulture, string outputPath)
        {
            _searchDirectory = searchDirectory;
            _rootCulture = rootCulture;
            _outputPath = outputPath;
        }

        public void Export()
        {
            if (Directory.Exists(_searchDirectory))
            {
                var filePaths = Directory.GetFiles(_searchDirectory, "*.resx", SearchOption.TopDirectoryOnly);
                var workbook = NpoiHelper.CreateWorkBook();

                foreach (var filePath in filePaths)
                {
                    var fileName = Path.GetFileName(filePath);
                    Console.WriteLine("Processing file {0}...", fileName);
                    
                    var resourceName = GetResourceName(fileName);
                    var culture = GetCulture(fileName);

                    NamesToWorkbook(workbook, filePath, resourceName, culture);
                }

                NpoiHelper.WriteWorkBook(workbook, _outputPath);
            }
            else
            {
                Console.WriteLine("Directory {0} does not exist.", _searchDirectory);
            }
        }

        private void NamesToWorkbook(HSSFWorkbook workbook, string resourcePath, string sheetName, string culture)
        {
            var sheet = NpoiHelper.GetSheet(workbook, sheetName);
            using (var resxReader = new ResXResourceReader(resourcePath))
            {
                var cultureRow = NpoiHelper.GetRow(sheet, CultureRowNum);
                NpoiHelper.SetCell(cultureRow, _currentColumn, culture);

                foreach (DictionaryEntry reader in resxReader)
                {
                    var exists = false;
                    for (var i = 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = NpoiHelper.GetRow(sheet, i);
                        if (row.GetCell(NameColumnNum).StringCellValue == reader.Key.ToString())
                        {
                            NpoiHelper.SetCell(row, _currentColumn, reader.Value.ToString());
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        var nameRow = NpoiHelper.GetRow(sheet, sheet.LastRowNum + 1);
                        NpoiHelper.SetCell(nameRow, NameColumnNum, reader.Key.ToString());
                        NpoiHelper.SetCell(nameRow, _currentColumn, reader.Value.ToString());
                    }
                }
                _currentColumn++;
            }
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
                var cultureInfo = new CultureInfo(culture);
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
