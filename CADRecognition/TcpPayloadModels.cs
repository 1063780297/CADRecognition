using System.Collections.Generic;

namespace CADRecognition
{
    public sealed class TcpWorksheetRow
    {
        public int Index { get; set; }
        public string Definition { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public string Length { get; set; } = string.Empty;
        public string StartAddress { get; set; } = string.Empty;
        public string DataSource { get; set; } = string.Empty;
        public string Desc1 { get; set; } = string.Empty;
        public string Desc2 { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public bool IsEditable { get; set; }
        public string FieldKey { get; set; } = string.Empty;
        public bool IsCustomField { get; set; }
    }

    public sealed class TcpCustomContentStore
    {
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();
    }

    public sealed class TcpConnectionHistoryStore
    {
        public List<string> Hosts { get; set; } = new List<string>();
        public List<string> Ports { get; set; } = new List<string>();
        public string LastHost { get; set; } = string.Empty;
        public string LastPort { get; set; } = string.Empty;
    }
}
