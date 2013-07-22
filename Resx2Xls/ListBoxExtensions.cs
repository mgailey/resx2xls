using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Resx2Xls
{
    public static class ListBoxExtensions
    {
        public static void AddCulturesFrom(this ListBox listBox, string dirPath, string filePattern = "*.resx")
        {
            listBox.Items.Clear();
            listBox.SelectedItems.Clear();
            if (string.IsNullOrEmpty(dirPath)) return;

            var files = Directory.GetFiles(dirPath, filePattern);
            foreach (var f in files.Select(f => new FileInfo(f)))
            {
                var culture = f.GetCulture();
                if (culture == null) continue;
                if (listBox.Items.IndexOf(culture) == -1)
                    listBox.Items.Add(culture);
            }
        }
    }
}