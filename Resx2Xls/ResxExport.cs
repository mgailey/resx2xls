using System;
using System.Collections;
using System.ComponentModel.Design;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Resx2Xls
{
    public class ResxExport
    {
        private static readonly string KeyColumn = "Key";
        private static readonly string SourceTextColumn = "Default text";
        private static readonly string TranslatedTextColumn = "Translated text";
        private static readonly string UsageColumn = "Usage";
        private readonly string[] excludedKeyList;
        private readonly DirectoryInfo resxDirectory;
        private readonly DirectoryInfo xlsDirectory;
        private DataTable Data { get; set; }

        public ResxExport(DirectoryInfo resxDirectory, DirectoryInfo xlsDirectory, string[] excludedKeyList = null) 
        {
            this.resxDirectory = resxDirectory;
            this.xlsDirectory = xlsDirectory;
            this.excludedKeyList = excludedKeyList ?? new string[] { };
            Data = new DataTable();
            Data.Columns.Add(KeyColumn);
            Data.Columns.Add(SourceTextColumn);
            Data.Columns.Add(TranslatedTextColumn);
            Data.Columns.Add(UsageColumn);

            var neutralFile = resxDirectory.GetFiles("*.resx").FirstOrDefault(cult => cult.GetCulture() == null);
            ReadNeutralResx(neutralFile);
        }

        public void Export(params string[] cultures)
        {
            if (!xlsDirectory.Exists) return;
            if (!resxDirectory.Exists) return;

            Array.ForEach(cultures, culture =>
                                        {
                                            var cultureFile = resxDirectory.GetFiles(String.Format("*.{0}.resx", culture))
                                                .FirstOrDefault(cult => cult.GetCulture() != null && cult.GetCulture().Name == culture);
                                            Export(cultureFile);
                                        });
        }

        public void Export(FileInfo cultureFile)
        {
            if (cultureFile == null) throw new Exception("File not found");
            var culture = cultureFile.GetCulture();
            if (culture == null) throw new Exception("Culture not found");

            AppendCulture(cultureFile, culture.Name);
            var destFile = new FileInfo(Path.Combine(xlsDirectory.FullName, Path.ChangeExtension(cultureFile.Name, ".xlsx")));
            DataTableToXls(destFile, culture.Name);
        }

        private void ReadNeutralResx(FileSystemInfo neutralFile)
        {
            using (var reader = new ResXResourceReader(neutralFile.FullName))
            {
                reader.UseResXDataNodes = true;
                foreach (DictionaryEntry de in reader)
                {
                    var dataNode = de.Value as ResXDataNode;
                    if (dataNode == null) continue;

                    var key = (string)de.Key;
                    var exclude = excludedKeyList.Any(key.EndsWith);
                    if (exclude) continue;

                    var value = (string)dataNode.GetValue((ITypeResolutionService)null);

                    var r = Data.NewRow();

                    r[KeyColumn] = key;

                    value = value.Replace("\r", "\\r");
                    value = value.Replace("\n", "\\n");

                    r[SourceTextColumn] = value;
                    r[UsageColumn] = dataNode.Comment;

                    Data.Rows.Add(r);
                }
            }
        }

        private void AppendCulture(FileSystemInfo cultureFile, string culture)
        {
            using (var reader = new ResXResourceReader(cultureFile.FullName))
            {
                foreach (DictionaryEntry de in reader)
                {
                    if (!(de.Value is string)) continue;

                    var key = (string)de.Key;

                    var exclude = excludedKeyList.Any(key.EndsWith);
                    if (exclude) continue;

                    var value = de.Value.ToString();

                    var strWhere = String.Format("Key='{0}'", key);
                    var rows = Data.Select(strWhere);
                    if (rows.Length == 0) throw new Exception("Row not Found");
                    var row = rows[0];
                    if (!Data.Columns.Contains(culture))
                        Data.Columns.Add(culture);

                    // update row
                    row.BeginEdit();

                    value = value.Replace("\r", "\\r");
                    value = value.Replace("\n", "\\n");
                    row[culture] = value;

                    row.EndEdit();
                }
            }
        }

        private void DataTableToXls(FileSystemInfo destFile, string culture)
        {
            var app = new Application();
            try
            {
                var wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                try
                {
                    var sheets = wb.Worksheets;
                    var sheet = (Worksheet)sheets.Item[1];
                    sheet.Name = string.Format("Translations ({0})", culture);

                    var headerRow = sheet.Rows[1, Type.Missing];
                    headerRow.Font.Bold = true;
                    var bottomEdge = headerRow.Borders[XlBordersIndex.xlEdgeBottom];
                    bottomEdge.LineStyle = XlLineStyle.xlContinuous;
                    bottomEdge.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                    sheet.Cells[1, 1] = KeyColumn;
                    sheet.Cells[1, 2] = SourceTextColumn;
                    sheet.Cells[1, 3] = TranslatedTextColumn;
                    sheet.Cells[1, 4] = UsageColumn;

                    var dw = Data.DefaultView;
                    dw.Sort = KeyColumn;

                    var row = 2;

                    foreach (var r in from DataRowView drw in dw select drw.Row)
                    {
                        sheet.Cells[row, 1] = r[KeyColumn];
                        sheet.Cells[row, 2] = r[SourceTextColumn];
                        sheet.Cells[row, 3] = r[culture];
                        sheet.Cells[row, 4] = r[UsageColumn];

                        row++;
                    }

                    sheet.Cells.Range["A1", "Z1"].EntireColumn.AutoFit();

                    // Save the Workbook and quit Excel.
                    wb.SaveAs(destFile.FullName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange,
                              Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                finally
                {
                    wb.Close(false, Missing.Value, Missing.Value);
                }
            }
            finally
            {
                app.Quit();   
            }
        }
    }
}