namespace CADRecognition
{
    /// <summary>
    /// Modbus 保持寄存器区布局（长度单位为 16 位字），与《功能需求》表一致；板厚/备用3 起始地址按连续字偏移修正。
    /// </summary>
    internal static class ModbusTcpExportLayout
    {
        /// <summary>程序名，UTF-16，20 字。</summary>
        public const int ProgramNameStart = 0;

        public const int ProgramNameWordLength = 20;

        /// <summary>程序号及后续 INT 块起始。</summary>
        public const int IntBlockStart = 20;

        /// <summary>程序号、左右门…备用2，共 10 个 INT。</summary>
        public const int IntBlockWordLength = 10;

        /// <summary>板长、板宽、板厚、备用3，各 REAL 2 字。</summary>
        public const int RealBlockStart = 30;

        public const int RealBlockWordLength = 8;

        /// <summary>备用4，DINT 2 字。</summary>
        public const int Spare4Start = 38;

        public const int Spare4WordLength = 2;

        public const int Stage1CoordStart = 40;

        public const int Stage1CoordWordLength = 320;

        public const int Stage1PositionMoldStart = 360;

        public const int Stage1PositionMoldWordLength = 80;

        public const int Stage1PunchMoldStart = 440;

        public const int Stage1PunchMoldWordLength = 80;

        public const int Stage2CoordStart = 520;

        public const int Stage2CoordWordLength = 320;

        public const int Stage2PositionMoldStart = 840;

        public const int Stage2PositionMoldWordLength = 80;

        public const int Stage2PunchMoldStart = 920;

        public const int Stage2PunchMoldWordLength = 80;

        /// <summary>台1/台2 坐标与模具编号槽位数。</summary>
        public const int CoordinateSlotCount = 40;

        public const int MoldSlotCount = 40;

        /// <summary>表占用总字数（920 + 80）。</summary>
        public const int TotalWordLength = 1000;
    }
}
