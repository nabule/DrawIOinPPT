using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace DrawioPpt.PowerPointAddIn.UI
{
    public class UrlDiagramEditorForm : Form
    {
        private readonly string _editorUrl;
        private readonly string _diagramName;
        private readonly string _initialXml;
        private readonly JavaScriptSerializer _serializer;
        private readonly WebView2 _webView;
        private bool _loadSent;
        private bool _closeAfterExport;
        private string _pendingXml;

        public UrlDiagramEditorForm(string editorUrl, string diagramName, string initialXml)
        {
            _editorUrl = editorUrl ?? string.Empty;
            _diagramName = string.IsNullOrWhiteSpace(diagramName) ? "Draw.io Diagram" : diagramName;
            _initialXml = initialXml ?? string.Empty;
            _serializer = new JavaScriptSerializer();
            _serializer.MaxJsonLength = int.MaxValue;

            this.Text = _diagramName;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.MinimizeBox = true;
            this.MaximizeBox = true;
            this.Width = 1400;
            this.Height = 900;

            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            this.Shown += OnShown;
        }

        public event EventHandler<UrlDiagramSavedEventArgs> DiagramSaved;

        private async void OnShown(object sender, EventArgs e)
        {
            await InitializeWebViewAsync();
        }

        private async Task InitializeWebViewAsync()
        {
            try
            {
                await _webView.EnsureCoreWebView2Async();
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _webView.NavigateToString(BuildHostPageHtml());
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "无法初始化 URL 编辑器。" + Environment.NewLine + ex.Message,
                    "DrawioPpt",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string rawMessage;

            try
            {
                rawMessage = e.TryGetWebMessageAsString();
            }
            catch
            {
                return;
            }

            Dictionary<string, object> message = DeserializeDictionary(rawMessage);
            if (message == null)
            {
                return;
            }

            string messageType = GetString(message, "type");
            if (!string.Equals(messageType, "drawio", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            Dictionary<string, object> payload = GetDictionary(message, "payload");
            if (payload == null)
            {
                return;
            }

            string eventName = GetString(payload, "event");
            if (string.IsNullOrWhiteSpace(eventName))
            {
                return;
            }

            if (string.Equals(eventName, "init", StringComparison.OrdinalIgnoreCase))
            {
                if (!_loadSent)
                {
                    _loadSent = true;
                    SendAction(new Dictionary<string, object>
                    {
                        { "action", "load" },
                        { "xml", _initialXml },
                        { "autosave", 1 },
                        { "saveAndExit", 1 },
                        { "title", _diagramName }
                    });
                }

                return;
            }

            if (string.Equals(eventName, "save", StringComparison.OrdinalIgnoreCase))
            {
                _pendingXml = GetString(payload, "xml");
                _closeAfterExport = GetBoolean(payload, "exit");
                SendAction(new Dictionary<string, object>
                {
                    { "action", "export" },
                    { "format", "xmlsvg" },
                    { "embedImages", true }
                });
                return;
            }

            if (string.Equals(eventName, "export", StringComparison.OrdinalIgnoreCase))
            {
                string svgMarkup = DecodeSvgMarkup(GetString(payload, "data"));
                if (!string.IsNullOrWhiteSpace(_pendingXml) && !string.IsNullOrWhiteSpace(svgMarkup))
                {
                    EventHandler<UrlDiagramSavedEventArgs> handler = this.DiagramSaved;
                    if (handler != null)
                    {
                        handler(this, new UrlDiagramSavedEventArgs(_pendingXml, svgMarkup, _closeAfterExport));
                    }
                }

                if (_closeAfterExport)
                {
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }

                _closeAfterExport = false;
                return;
            }

            if (string.Equals(eventName, "exit", StringComparison.OrdinalIgnoreCase))
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return;
            }

            if (string.Equals(eventName, "openLink", StringComparison.OrdinalIgnoreCase))
            {
                string href = GetString(payload, "href");
                if (!string.IsNullOrWhiteSpace(href))
                {
                    try
                    {
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = href;
                        startInfo.UseShellExecute = true;
                        Process.Start(startInfo);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void SendAction(object action)
        {
            if (_webView.CoreWebView2 == null)
            {
                return;
            }

            string json = _serializer.Serialize(action);
            string script = "window.dispatchHostMessage(" + json + ");";
            Task task = _webView.CoreWebView2.ExecuteScriptAsync(script);
        }

        private string BuildHostPageHtml()
        {
            string editorUrl = EnsureEmbedUrl(_editorUrl);
            string escapedUrl = SecurityElement.Escape(editorUrl);

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("<!DOCTYPE html>");
            builder.AppendLine("<html>");
            builder.AppendLine("<head>");
            builder.AppendLine("  <meta charset=\"utf-8\" />");
            builder.AppendLine("  <style>");
            builder.AppendLine("    html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; background: #edf2f7; }");
            builder.AppendLine("    iframe { width: 100%; height: 100%; border: 0; display: block; }");
            builder.AppendLine("  </style>");
            builder.AppendLine("</head>");
            builder.AppendLine("<body>");
            builder.AppendLine("  <iframe id=\"editor\" src=\"" + escapedUrl + "\"></iframe>");
            builder.AppendLine("  <script>");
            builder.AppendLine("    (function () {");
            builder.AppendLine("      var editorFrame = document.getElementById('editor');");
            builder.AppendLine("      function postToHost(message) {");
            builder.AppendLine("        window.chrome.webview.postMessage(JSON.stringify(message));");
            builder.AppendLine("      }");
            builder.AppendLine("      window.dispatchHostMessage = function (message) {");
            builder.AppendLine("        if (editorFrame && editorFrame.contentWindow) {");
            builder.AppendLine("          editorFrame.contentWindow.postMessage(message, '*');");
            builder.AppendLine("        }");
            builder.AppendLine("      };");
            builder.AppendLine("      window.addEventListener('message', function (event) {");
            builder.AppendLine("        if (!editorFrame || event.source !== editorFrame.contentWindow) {");
            builder.AppendLine("          return;");
            builder.AppendLine("        }");
            builder.AppendLine("        postToHost({ type: 'drawio', payload: event.data });");
            builder.AppendLine("      });");
            builder.AppendLine("      postToHost({ type: 'hostReady' });");
            builder.AppendLine("    })();");
            builder.AppendLine("  </script>");
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
            return builder.ToString();
        }

        private static string EnsureEmbedUrl(string editorUrl)
        {
            string value = string.IsNullOrWhiteSpace(editorUrl)
                ? "https://embed.diagrams.net/?embed=1&proto=json&spin=1&saveAndExit=1"
                : editorUrl.Trim();

            if (value.IndexOf("embed=", StringComparison.OrdinalIgnoreCase) < 0)
            {
                value = AppendQueryParameter(value, "embed=1");
            }

            if (value.IndexOf("proto=json", StringComparison.OrdinalIgnoreCase) < 0)
            {
                value = AppendQueryParameter(value, "proto=json");
            }

            if (value.IndexOf("spin=", StringComparison.OrdinalIgnoreCase) < 0)
            {
                value = AppendQueryParameter(value, "spin=1");
            }

            if (value.IndexOf("saveAndExit=", StringComparison.OrdinalIgnoreCase) < 0 &&
                value.IndexOf("noSaveBtn=1", StringComparison.OrdinalIgnoreCase) < 0)
            {
                value = AppendQueryParameter(value, "saveAndExit=1");
            }

            return value;
        }

        private static string AppendQueryParameter(string url, string queryPart)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return queryPart;
            }

            return url.IndexOf('?') >= 0 ? url + "&" + queryPart : url + "?" + queryPart;
        }

        private Dictionary<string, object> DeserializeDictionary(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            try
            {
                return _serializer.DeserializeObject(json) as Dictionary<string, object>;
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<string, object> GetDictionary(Dictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key) || !parent.ContainsKey(key))
            {
                return null;
            }

            return parent[key] as Dictionary<string, object>;
        }

        private static string GetString(Dictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key) || !parent.ContainsKey(key) || parent[key] == null)
            {
                return string.Empty;
            }

            return parent[key].ToString();
        }

        private static bool GetBoolean(Dictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key) || !parent.ContainsKey(key) || parent[key] == null)
            {
                return false;
            }

            object value = parent[key];
            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            if (bool.TryParse(value.ToString(), out parsed))
            {
                return parsed;
            }

            int number;
            if (int.TryParse(value.ToString(), out number))
            {
                return number != 0;
            }

            return false;
        }

        private static string DecodeSvgMarkup(string dataUri)
        {
            if (string.IsNullOrWhiteSpace(dataUri))
            {
                return string.Empty;
            }

            int separatorIndex = dataUri.IndexOf(',');
            if (separatorIndex < 0)
            {
                return string.Empty;
            }

            string metadata = dataUri.Substring(0, separatorIndex);
            string payload = dataUri.Substring(separatorIndex + 1);

            try
            {
                if (metadata.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    byte[] bytes = Convert.FromBase64String(payload);
                    return Encoding.UTF8.GetString(bytes);
                }

                return Uri.UnescapeDataString(payload);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
