using System.Data;

namespace Resx2Xls
{
    public class Resx
    {
        public static readonly string KeyColumn = "Key";
        public static readonly string SourceTextColumn = "Source Text";
        public static readonly string UsageColumn = "Usage";

        public Resx()
        {
            Data = new DataTable();
            Data.Columns.Add(KeyColumn);
            Data.Columns.Add(SourceTextColumn);
            Data.Columns.Add(UsageColumn);
        }

        public DataTable Data { get; private set; }
    }
}