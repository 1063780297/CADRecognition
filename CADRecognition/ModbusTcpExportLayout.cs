namespace CADRecognition
{
    /// <summary>
    /// Modbus 保持寄存器区布局（长度单位为 16 位字），字段顺序和类型与表格一致。
    /// REAL/DINT 均占 2 个 16 位寄存器。
    /// </summary>
    internal static class ModbusTcpExportLayout
    {
        /// <summary>1 程序名，Unicode 字符串，20 字（支持中文，最多 20 个 16 位字符）。</summary>
        public const int ProgramNameStart = 0;

        public const int ProgramNameWordLength = 20;

        /// <summary>2-11 程序号、左右门…备用2，共 10 个 INT。</summary>
        public const int IntBlockStart = 20;

        public const int IntBlockWordLength = 10;

        /// <summary>12-15 板长、板宽、板厚、备用3，各 REAL 2 字。</summary>
        public const int RealBlockStart = 30;

        public const int RealBlockWordLength = 8;

        /// <summary>16 备用4，DINT 2 字。</summary>
        public const int Spare4Start = 38;

        public const int Spare4WordLength = 2;

        /// <summary>17 台1图纸坐标组 (X,Y)，(REAL*2)*100，共 400 字。</summary>
        public const int Stage1CoordStart = 40;

        public const int Stage1CoordWordLength = 400;

        /// <summary>19 台1位置模具编号组，INT*100，共 100 字。</summary>
        public const int Stage1PositionMoldStart = 440;

        public const int Stage1PositionMoldWordLength = 100;

        /// <summary>20 台1冲孔模具编号组，DINT*100，共 200 字。</summary>
        public const int Stage1PunchMoldStart = 540;

        public const int Stage1PunchMoldWordLength = 200;

        /// <summary>21 台2图纸坐标组 (X,Y)，(REAL*2)*100，共 400 字。</summary>
        public const int Stage2CoordStart = 740;

        public const int Stage2CoordWordLength = 400;

        /// <summary>23 台2位置模具编号组，INT*100，共 100 字。</summary>
        public const int Stage2PositionMoldStart = 1140;

        public const int Stage2PositionMoldWordLength = 100;

        /// <summary>24 台2冲孔模具编号组，DINT*100，共 200 字。</summary>
        public const int Stage2PunchMoldStart = 1240;

        public const int Stage2PunchMoldWordLength = 200;

        /// <summary>台1/台2 坐标与模具编号槽位数。</summary>
        public const int CoordinateSlotCount = 100;

        public const int MoldSlotCount = 100;

        /// <summary>表占用总字数（1240 + 200）。</summary>
        public const int TotalWordLength = 1440;
    }
}
