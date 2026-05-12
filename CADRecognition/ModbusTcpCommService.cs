using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HslCommunication.ModBus;

namespace CADRecognition
{
    /// <summary>
    /// 按固定字地址将 <see cref="TcpExportModel"/> 写入 Modbus TCP 保持寄存器（与功能需求表长度一致）。
    /// 坐标区 80 字：40 点 ×（X、Y 各占 1 字），与表实际导出数据一致。
    /// </summary>
    internal sealed class ModbusTcpCommService
    {
        private static readonly Regex TrailingDigits = new(@"\d+$", RegexOptions.Compiled);

        public async Task SendExportModelAsync(string host, int port, byte station, string registerBaseText, TcpExportModel model)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                throw new ArgumentException("Modbus 主机不能为空。", nameof(host));
            }

            if (port <= 0 || port > 65535)
            {
                throw new ArgumentOutOfRangeException(nameof(port), "Modbus 端口必须在 1-65535 之间。");
            }

            if (model is null)
            {
                throw new ArgumentNullException(nameof(model));
            }

            var (addressPrefix, baseOffset) = ParseRegisterBase(registerBaseText);
            string Addr(int relativeWord) => string.IsNullOrEmpty(addressPrefix)
                ? (baseOffset + relativeWord).ToString(CultureInfo.InvariantCulture)
                : addressPrefix + (baseOffset + relativeWord).ToString(CultureInfo.InvariantCulture);

            using var client = new ModbusTcpNet(host, port, station)
            {
                AddressStartWithZero = true
            };

            await WriteProgramNameAsync(client, Addr, model.ProgramName).ConfigureAwait(false);

            var programNoInt = ParseProgramNo(model.ProgramNo);
            var intBlock = new short[]
            {
                ToInt16(programNoInt),
                ToInt16(model.LeftRightDoor),
                ToInt16(model.Material),
                ToInt16(model.Type),
                ToInt16(model.FormingLength),
                ToInt16(model.FormingWidth),
                ToInt16(model.FormingThickness),
                ToInt16(model.Stage1PunchCount),
                ToInt16(model.Stage2PunchCount),
                ToInt16(model.Spare2)
            };

            if (intBlock.Length != ModbusTcpExportLayout.IntBlockWordLength)
            {
                throw new InvalidOperationException("INT 块长度与布局定义不一致。");
            }

            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.IntBlockStart), intBlock).ConfigureAwait(false));

            var plateWidthForTable = model.PlateWidth2 > 0 ? model.PlateWidth2 : model.PlateWidth;
            var reals = new[]
            {
                (float)model.PlateLength,
                (float)plateWidthForTable,
                (float)model.PlateThickness,
                (float)model.Spare3
            };

            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.RealBlockStart), reals).ConfigureAwait(false));

            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Spare4Start), new[] { model.Spare4 }).ConfigureAwait(false));

            var stage1Coords = BuildCoordinateWords(model.Stage1DiagramCoordinates);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage1CoordStart), stage1Coords).ConfigureAwait(false));

            var stage1Pos = BuildMoldDints(model.Stage1PositionMoldIds);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage1PositionMoldStart), stage1Pos).ConfigureAwait(false));

            var stage1Punch = BuildMoldDints(model.Stage1PunchMoldIds);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage1PunchMoldStart), stage1Punch).ConfigureAwait(false));

            var stage2Coords = BuildCoordinateWords(model.Stage2DiagramCoordinates);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage2CoordStart), stage2Coords).ConfigureAwait(false));

            var stage2Pos = BuildMoldDints(model.Stage2PositionMoldIds);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage2PositionMoldStart), stage2Pos).ConfigureAwait(false));

            var stage2Punch = BuildMoldDints(model.Stage2PunchMoldIds);
            ThrowIfFailed(await client.WriteAsync(Addr(ModbusTcpExportLayout.Stage2PunchMoldStart), stage2Punch).ConfigureAwait(false));
        }

        private static async Task WriteProgramNameAsync(ModbusTcpNet client, Func<int, string> addr, string programName)
        {
            var text = programName ?? string.Empty;
            var encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);
            var bytes = new byte[ModbusTcpExportLayout.ProgramNameWordLength * 2];
            var encoded = encoding.GetBytes(text);
            Array.Copy(encoded, bytes, Math.Min(encoded.Length, bytes.Length));

            ThrowIfFailed(await client.WriteAsync(addr(ModbusTcpExportLayout.ProgramNameStart), bytes).ConfigureAwait(false));
        }


        private static int[] BuildMoldDints(System.Collections.Generic.IReadOnlyList<string> moldIds)
        {
            var buf = new int[ModbusTcpExportLayout.MoldSlotCount];
            var n = Math.Min(moldIds.Count, ModbusTcpExportLayout.MoldSlotCount);
            for (var i = 0; i < n; i++)
            {
                buf[i] = MoldStringToDInt(moldIds[i]);
            }

            return buf;
        }

        private static int MoldStringToDInt(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return 0;
            }

            var m = TrailingDigits.Match(s.Trim());
            if (m.Success && int.TryParse(m.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
            {
                return v;
            }

            return 0;
        }

        private static int ParseProgramNo(string programNo)
        {
            if (string.IsNullOrWhiteSpace(programNo))
            {
                return 0;
            }

            var m = TrailingDigits.Match(programNo.Trim());
            if (m.Success && int.TryParse(m.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
            {
                return v;
            }

            return 0;
        }

        private static short ToInt16(int v)
        {
            if (v < short.MinValue)
            {
                return short.MinValue;
            }

            if (v > short.MaxValue)
            {
                return short.MaxValue;
            }

            return (short)v;
        }

        private static short[] BuildCoordinateWords(System.Collections.Generic.IReadOnlyList<TcpCoordinateRow> rows)
        {
            var buf = new short[ModbusTcpExportLayout.CoordinateSlotCount * 2];
            var n = Math.Min(rows.Count, ModbusTcpExportLayout.CoordinateSlotCount);
            for (var i = 0; i < n; i++)
            {
                buf[i * 2] = ToInt16((int)Math.Round(rows[i].X, MidpointRounding.AwayFromZero));
                buf[i * 2 + 1] = ToInt16((int)Math.Round(rows[i].Y, MidpointRounding.AwayFromZero));
            }

            return buf;
        }

        private static void ThrowIfFailed(HslCommunication.OperateResult result)
        {
            if (result is null || !result.IsSuccess)
            {
                var msg = result?.Message;
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(msg) ? "Modbus 写入失败。" : msg);
            }
        }

        /// <summary>
        /// 解析保持寄存器基址：支持纯数字或 Hsl 前缀如 <c>s=1;</c> 后接数字。
        /// </summary>
        private static (string Prefix, int Base) ParseRegisterBase(string text)
        {
            var t = text?.Trim() ?? string.Empty;
            if (t.Length == 0)
            {
                return (string.Empty, 0);
            }

            var semi = t.LastIndexOf(';');
            if (semi >= 0)
            {
                var prefix = t.Substring(0, semi + 1);
                var num = t.Substring(semi + 1).Trim();
                return int.TryParse(num, NumberStyles.Integer, CultureInfo.InvariantCulture, out var b)
                    ? (prefix, b)
                    : (prefix, 0);
            }

            return int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var baseOffset)
                ? (string.Empty, baseOffset)
                : (string.Empty, 0);
        }
    }
}
