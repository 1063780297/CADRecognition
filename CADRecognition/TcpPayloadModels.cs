using System.Collections.Generic;

namespace CADRecognition
{
    public sealed class TcpCustomContentStore
    {
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();
        public string Encoding { get; set; } = "UTF-8";
        public bool SwapBytes { get; set; } = false;
    }

    public sealed class TcpConnectionHistoryStore
    {
        public List<string> Hosts { get; set; } = new List<string>();
        public List<string> Ports { get; set; } = new List<string>();
        public string LastHost { get; set; } = string.Empty;
        public string LastPort { get; set; } = string.Empty;
        public string LastStation { get; set; } = "1";
        public string LastHoldingRegisterAddress { get; set; } = "0";
    }
}
