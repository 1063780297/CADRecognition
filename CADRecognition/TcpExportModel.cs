using System.Collections.Generic;

namespace CADRecognition
{
    /// <summary>
    /// 与《功能需求.xlsx》对应的 TCP 导出数据结构。
    /// </summary>
    public sealed class TcpExportModel
    {
        public string ProgramName { get; set; } = string.Empty;
        public string ProgramNo { get; set; }
        public int LeftRightDoor { get; set; }
        public int Material { get; set; }
        public int Type { get; set; }
        public int FormingLength { get; set; }
        public int FormingWidth { get; set; }
        public int FormingThickness { get; set; }
        public int Stage1PunchCount { get; set; }
        public int Stage2PunchCount { get; set; }
        public int Spare2 { get; set; }
        public double PlateLength { get; set; }
        public double PlateWidth { get; set; }
        public double PlateThickness { get; set; }
        public double Spare3 { get; set; }
        public int Spare4 { get; set; }
        public List<TcpCoordinateRow> Stage1DiagramCoordinates { get; set; } = [];
        public List<string> Stage1PositionMoldIds { get; set; } = [];
        public List<string> Stage1PunchMoldIds { get; set; } = [];
        public List<TcpCoordinateRow> Stage2DiagramCoordinates { get; set; } = [];
        public List<string> Stage2PositionMoldIds { get; set; } = [];
        public List<string> Stage2PunchMoldIds { get; set; } = [];
        public string CustomContent { get; set; } = string.Empty;
    }

    public sealed class TcpCoordinateRow
    {
        public double X { get; set; }
        public double Y { get; set; }
    }
}
