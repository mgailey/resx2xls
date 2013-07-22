using System.Globalization;
using System.IO;

namespace Resx2Xls
{
    public static class FileInfoExtensions
    {
        public static CultureInfo GetCulture(this FileInfo fi)
        {
            //Remove the extension and return the string	
            var cult = new FileInfo(Path.GetFileNameWithoutExtension(fi.Name) ?? string.Empty).Extension.TrimStart('.');

            if (string.IsNullOrEmpty(cult)) return null;

            try
            {
                return new CultureInfo(cult);
            }
            catch
            {
                return null;
            }
        }
    }
}