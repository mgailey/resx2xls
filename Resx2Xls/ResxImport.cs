using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using Microsoft.Office.Interop.Excel;

namespace Resx2Xls
{
    public class ResxImport
    {
        private readonly DirectoryInfo resxDirectory;
        private readonly DirectoryInfo xlsDirectory;

        public ResxImport(DirectoryInfo resxDirectory, DirectoryInfo xlsDirectory)
        {
            this.resxDirectory = resxDirectory;
            this.xlsDirectory = xlsDirectory;
        }

        public void Import(params string[] cultures)
        {
            if (!resxDirectory.Exists) return;
            if (!xlsDirectory.Exists) return;

            Array.ForEach(cultures, culture =>
            {
                var xlsFile = xlsDirectory.GetFiles(String.Format("*.{0}.xlsx", culture))
                    .FirstOrDefault(cult => cult.GetCulture() != null && cult.GetCulture().Name == culture);

                Import(xlsFile);
            });
        }

        public void Import(FileInfo fileInfo)
        {
            if (fileInfo == null) throw new Exception("File not found");
            var culture = fileInfo.GetCulture();
            if (culture == null) throw new Exception("Culture not found");

            var app = new Application();
            try
            {
                var wb = app.Workbooks.Open(fileInfo.FullName, 0, false, 5, string.Empty, string.Empty, false, XlPlatform.xlWindows, string.Empty,
                                            true, false, 0, true, false, false);
                try
                {
                    var sheets = wb.Worksheets;
                    var sheet = (Worksheet) sheets.Item[1];

                    var fileDest = Path.Combine(resxDirectory.FullName, Path.ChangeExtension(fileInfo.Name, ".resx"));
                    using (var rw = new ResXResourceWriter(fileDest))
                    {
                        var row = 2;
                        string key;
                        do
                        {
                            key = ((Range) sheet.Cells[row, 1]).Text as string;
                            if (string.IsNullOrEmpty(key)) continue;

                            var text = ((Range) sheet.Cells[row, 3]).Text as string;
                            if (!string.IsNullOrEmpty(text)) rw.AddResource(new ResXDataNode(key, text));
                            row++;
                        } while (!String.IsNullOrEmpty(key));
                    }
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