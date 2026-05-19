using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HslCommunication.ModBus;

namespace CADRecognition
{
    /// <summary>
    /// 按固定字地址将 <see cref="TcpExportModel"/> 写入 Modbus TCP 保持寄存器。
    /// 首字段使用 UTF-8 字符串写入，支持中文；其余字段的顺序、数组长度和 INT/REAL/DINT 宽度与表格定义一致。
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

            await WriteInt16BlockAsync(client, Addr, ModbusTcpExportLayout.IntBlockStart, intBlock).ConfigureAwait(false);

            var plateWidthForTable = model.PlateWidth2 > 0 ? model.PlateWidth2 : model.PlateWidth;
            var reals = new[]
            {
                (float)model.PlateLength,
                (float)plateWidthForTable,
                (float)model.PlateThickness,
                (float)model.Spare3
            };

            await WriteFloatBlockAsync(client, Addr, ModbusTcpExportLayout.RealBlockStart, reals).ConfigureAwait(false);

            await WriteInt16BlockAsync(client, Addr, ModbusTcpExportLayout.Spare4Start, new[] { ToInt16(model.Spare4) }).ConfigureAwait(false);

            await WriteFloatBlockAsync(client, Addr, ModbusTcpExportLayout.Stage1CoordStart, BuildCoordinateReals(model.Stage1DiagramCoordinates)).ConfigureAwait(false);

            await WriteInt16BlockAsync(client, Addr, ModbusTcpExportLayout.Stage1PositionMoldStart, BuildMoldInts(model.Stage1PositionMoldIds)).ConfigureAwait(false);

            await WriteDIntBlockAsync(client, Addr, ModbusTcpExportLayout.Stage1PunchMoldStart, BuildMoldDints(model.Stage1PunchMoldIds)).ConfigureAwait(false);

            await WriteFloatBlockAsync(client, Addr, ModbusTcpExportLayout.Stage2CoordStart, BuildCoordinateReals(model.Stage2DiagramCoordinates)).ConfigureAwait(false);

            await WriteInt16BlockAsync(client, Addr, ModbusTcpExportLayout.Stage2PositionMoldStart, BuildMoldInts(model.Stage2PositionMoldIds)).ConfigureAwait(false);

            await WriteDIntBlockAsync(client, Addr, ModbusTcpExportLayout.Stage2PunchMoldStart, BuildMoldDints(model.Stage2PunchMoldIds)).ConfigureAwait(false);

            ValidateLayout();
        }

        private static async Task WriteProgramNameAsync(ModbusTcpNet client, Func<int, string> addr, string programName)
        {
            var bytes = EncodeUtf8Fixed(programName ?? string.Empty, ModbusTcpExportLayout.ProgramNameWordLength);
            ThrowIfFailed(await client.WriteAsync(addr(ModbusTcpExportLayout.ProgramNameStart), bytes).ConfigureAwait(false));
        }

        private static byte[] EncodeUtf8Fixed(string text, int maxWordLength)
        {
            var maxBytes = maxWordLength * 2;
            var buffer = new byte[maxBytes];
            if (string.IsNullOrEmpty(text))
            {
                return buffer;
            }

            var encoded = Encoding.UTF8.GetBytes(text);
            var byteCount = Math.Min(encoded.Length, maxBytes);
            if (byteCount > 0)
            {
                Array.Copy(encoded, buffer, byteCount);
            }

            return buffer;
        }

        private static void ValidateLayout()
        {
            if (ModbusTcpExportLayout.ProgramNameWordLength != 20 ||
                ModbusTcpExportLayout.IntBlockWordLength != 10 ||
                ModbusTcpExportLayout.RealBlockWordLength != 8 ||
                ModbusTcpExportLayout.Spare4WordLength != 2 ||
                ModbusTcpExportLayout.Stage1CoordWordLength != ModbusTcpExportLayout.CoordinateSlotCount * 4 ||
                ModbusTcpExportLayout.Stage1PositionMoldWordLength != ModbusTcpExportLayout.MoldSlotCount ||
                ModbusTcpExportLayout.Stage1PunchMoldWordLength != ModbusTcpExportLayout.MoldSlotCount * 2 ||
                ModbusTcpExportLayout.Stage2CoordWordLength != ModbusTcpExportLayout.CoordinateSlotCount * 4 ||
                ModbusTcpExportLayout.Stage2PositionMoldWordLength != ModbusTcpExportLayout.MoldSlotCount ||
                ModbusTcpExportLayout.Stage2PunchMoldWordLength != ModbusTcpExportLayout.MoldSlotCount * 2)
            {
                throw new InvalidOperationException("Modbus 发送数据结构与表格定义不一致。");
            }
        }

        private static short[] BuildMoldInts(System.Collections.Generic.IReadOnlyList<string> moldIds)
        {
            var buf = new short[ModbusTcpExportLayout.MoldSlotCount];
            var n = Math.Min(moldIds.Count, ModbusTcpExportLayout.MoldSlotCount);
            for (var i = 0; i < n; i++)
            {
                buf[i] = ToInt16(MoldStringToInt(moldIds[i]));
            }

            return buf;
        }

        private static int[] BuildMoldDints(System.Collections.Generic.IReadOnlyList<string> moldIds)
        {
            var buf = new int[ModbusTcpExportLayout.MoldSlotCount];
            var n = Math.Min(moldIds.Count, ModbusTcpExportLayout.MoldSlotCount);
            for (var i = 0; i < n; i++)
            {
                buf[i] = MoldStringToInt(moldIds[i]);
            }

            return buf;
        }

        private static async Task WriteInt16BlockAsync(ModbusTcpNet client, Func<int, string> addr, int startWord, short[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                ThrowIfFailed(await client.WriteAsync(addr(startWord + i), new[] { (int)values[i] }).ConfigureAwait(false));
            }
        }

        private static async Task WriteDIntBlockAsync(ModbusTcpNet client, Func<int, string> addr, int startWord, int[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                ThrowIfFailed(await client.WriteAsync(addr(startWord + i * 2), values[i]).ConfigureAwait(false));
            }
        }

        private static async Task WriteFloatBlockAsync(ModbusTcpNet client, Func<int, string> addr, int startWord, float[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                ThrowIfFailed(await client.WriteAsync(addr(startWord + i * 2), values[i]).ConfigureAwait(false));
            }
        }

        private static int MoldStringToInt(string s)
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

        private static float[] BuildCoordinateReals(System.Collections.Generic.IReadOnlyList<TcpCoordinateRow> rows)
        {
            var buf = new float[ModbusTcpExportLayout.CoordinateSlotCount * 2];
            var n = Math.Min(rows.Count, ModbusTcpExportLayout.CoordinateSlotCount);
            for (var i = 0; i < n; i++)
            {
                buf[i * 2] = (float)rows[i].X;
                buf[i * 2 + 1] = (float)rows[i].Y;
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
