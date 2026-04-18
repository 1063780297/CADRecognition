using System;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CADRecognition
{
    internal sealed class TcpCommService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        public async Task SendJsonAsync(string host, int port, object payload)
        {
            if (string.IsNullOrWhiteSpace(host))
            {
                throw new ArgumentException("TCP 主机不能为空。", nameof(host));
            }

            if (port <= 0 || port > 65535)
            {
                throw new ArgumentOutOfRangeException(nameof(port), "TCP 端口必须在 1-65535 之间。");
            }

            var json = payload is string text ? text : JsonSerializer.Serialize(payload, JsonOptions);

            using var client = new TcpClient();
            await client.ConnectAsync(host, port).ConfigureAwait(false);
            using var stream = client.GetStream();
            var bytes = Encoding.UTF8.GetBytes(json);
            await stream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await stream.FlushAsync().ConfigureAwait(false);
        }
    }
}
