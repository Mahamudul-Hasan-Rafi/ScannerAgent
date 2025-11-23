using Newtonsoft.Json;
using ScannerAgent.Controllers;
using ScannerAgent.Model;
using ScannerAgent.Services;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ScannerAgent
{
    public class ScannerHttpListener
    {
        private readonly string _url;
        private readonly HttpListener _listener;
        private readonly ScannerController _controller;

        public ScannerHttpListener(string url, ScannerController controller)
        {
            _url = url;
            _listener = new HttpListener();
            _listener.Prefixes.Add(_url);
            _controller = controller;
        }

        public void Start()
        {
            _listener.Start();
            Console.WriteLine($"Listening on {_url}");
            Task.Run(() => ListenLoop());
        }

        private async Task ListenLoop()
        {
            while (true)
            {
                var context = await _listener.GetContextAsync();
                _ = Task.Run(() => ProcessRequest(context));
            }
        }

        private async Task ProcessRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var resp = context.Response;

            try
            {
                // Handle CORS
                if (req.HttpMethod == "OPTIONS")
                {
                    resp.Headers.Add("Access-Control-Allow-Origin", "*");
                    resp.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                    resp.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                    resp.StatusCode = 200;
                    resp.Close();
                    return;
                }

                // Ping
                if (req.Url.AbsolutePath == "/ping" && req.HttpMethod == "GET")
                {
                    WriteText(resp, "ok");
                    return;
                }

                // Scan
                if (req.Url.AbsolutePath == "/scan" && req.HttpMethod == "POST")
                {
                    string requestBody;
                    using (var reader = new StreamReader(req.InputStream))
                        requestBody = reader.ReadToEnd();

                    var scanRequest = JsonConvert.DeserializeObject<ScanRequest>(requestBody);
                    var result = await _controller.ScanAsync(scanRequest);

                    // Set HTTP status from service result and send a clean JSON payload
                    resp.StatusCode = result.statusCode;
                    var payload = new ScanResponse
                    {
                        images = result.images,
                        status = result.status,
                        statusCode = result.statusCode
                    };

                    WriteJsonResponse(resp, payload);
                    return;
                }

                // Scan Stream
                if (req.Url.AbsolutePath == "/scan/stream" && req.HttpMethod == "POST")
                {

                    string requestBody;
                    using (var reader = new StreamReader(req.InputStream))
                        requestBody = reader.ReadToEnd();

                    var scanRequest = JsonConvert.DeserializeObject<ScanRequest>(requestBody);

                    await _controller.StreamScannedDocumentAsync(context, scanRequest);
                    return;
                }

                // Devices
                if (req.Url.AbsolutePath == "/devices" && req.HttpMethod == "GET")
                {
                    var devices = _controller.Devices();
                    WriteJsonResponse(resp, devices);
                    return;
                }

                // Not found
                resp.StatusCode = 404;
                WriteJsonResponse(resp, new { error = "Endpoint not found" });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                resp.StatusCode = 500;
                WriteJsonResponse(resp, new
                {
                    error = "Internal server error",
                    message = ex.Message
                });
            }
        }

        private void WriteText(HttpListenerResponse resp, string text)
        {
            try
            {
                resp.ContentType = "text/plain; charset=utf-8";
                resp.Headers.Add("Access-Control-Allow-Origin", "*");

                var buffer = Encoding.UTF8.GetBytes(text ?? string.Empty);
                resp.ContentLength64 = buffer.LongLength;

                using (var outStream = resp.OutputStream)
                {
                    outStream.Write(buffer, 0, buffer.Length);
                    outStream.Flush();
                }
            }
            catch
            {
                // swallow to avoid throwing while trying to write error responses
            }
            finally
            {
                try { resp.OutputStream.Close(); } catch { }
            }
        }

        private void WriteJsonResponse(HttpListenerResponse resp, object obj)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                var buffer = Encoding.UTF8.GetBytes(json ?? string.Empty);

                resp.ContentType = "application/json; charset=utf-8";
                resp.Headers.Add("Access-Control-Allow-Origin", "*");

                // We know the payload length, so set Content-Length for predictable behavior.
                // For very large payloads you can instead use resp.SendChunked = true and stream chunks.
                resp.ContentLength64 = buffer.LongLength;

                using (var outStream = resp.OutputStream)
                {
                    outStream.Write(buffer, 0, buffer.Length);
                    outStream.Flush();
                }
            }
            catch (Exception e)
            {
                // Log and ignore - avoid throwing while trying to write error responses
                Console.WriteLine("WriteJsonResponse error: " + e.Message);
            }
            finally
            {
                try { resp.OutputStream.Close(); } catch { }
            }
        }
    }

}
