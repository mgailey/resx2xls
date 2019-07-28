using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Resx2Xls
{
	public static class XlsFileExtensions
	{		
        public static IReadOnlyCollection<Translation> ReadXlsTranslations(this FileInfo sourceFile)
        {
            if (sourceFile == null) throw new Exception("File not found");
            var culture = sourceFile.GetCulture();
            if (culture == null) throw new Exception("Culture not found");

            var translations = new List<Translation>();

            using (var package = new ExcelPackage(sourceFile))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var rowCount = worksheet.Dimension.End.Row;
                for (var row = 2; row <= rowCount; row++)
                {
                    translations.Add(new Translation
                    {
                        Key = worksheet.Cells[row, 1].Value.ToString().Trim(),
                        Source = worksheet.Cells[row, 2].Value.ToString().Trim(),
                        Native = worksheet.Cells[row, 3].Value.ToString().Trim(),
                        Usage = worksheet.Cells[row, 4].Value?.ToString().Trim()
                    });
                }
            }
            return translations;
        }

        public static void WriteXlsTranslations(this FileInfo destFile, IReadOnlyCollection<Translation> translations)
        {
            using (var excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add($"Translations ({destFile.GetCulture().Name})");
                var headerRow = new List<string[]>()
                {
                    new[] { "Key", "Default text", "Translated text", "Usage" }
                };
                var headerRange = "A1:" + char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[headerRange].Style.Font.Bold = true;
                
                var cellData = translations
                    .OrderBy(tr => tr.Key)
                    .Select(tr => new []{tr.Key, tr.Source, tr.Native, tr.Usage})
                    .ToList();
                worksheet.Cells[2, 1].LoadFromArrays(cellData);
                
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                excel.SaveAs(destFile);
            }
        }
	}
}