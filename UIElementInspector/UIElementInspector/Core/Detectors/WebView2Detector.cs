using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Detects elements in modern Chromium-based browsers using WebView2 and Chrome DevTools Protocol
    /// </summary>
    public class WebView2Detector : IElementDetector
    {
        #region Native Methods for Window Detection

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private const uint WM_GETOBJECT = 0x003D;

        #endregion

        public string Name => "WebView2/CDP";

        private WebView2 _webView;
        private bool _isInitialized;

        public WebView2Detector()
        {
            InitializeWebView();
        }

        private async void InitializeWebView()
        {
            try
            {
                _webView = new WebView2();
                await _webView.EnsureCoreWebView2Async();
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 initialization error: {ex.Message}");
                _isInitialized = false;
            }
        }

        public bool CanDetect(System.Windows.Point screenPoint)
        {
            try
            {
                // Check if the point is over a Chrome/Edge window
                var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                var hwnd = WindowFromPoint(point);

                if (hwnd == IntPtr.Zero)
                    return false;

                var className = new System.Text.StringBuilder(256);
                GetClassName(hwnd, className, className.Capacity);
                var classNameStr = className.ToString();

                // Check for Chrome, Edge, or embedded WebView2 windows
                return classNameStr.Contains("Chrome") ||
                       classNameStr.Contains("Edge") ||
                       classNameStr.Contains("WebView");
            }
            catch
            {
                return false;
            }
        }

        public async Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile)
        {
            if (!_isInitialized || _webView?.CoreWebView2 == null)
                return null;

            try
            {
                var stopwatch = Stopwatch.StartNew();
                var info = new ElementInfo
                {
                    DetectionMethod = Name,
                    CollectionProfile = profile.ToString(),
                    CaptureTime = DateTime.Now
                };

                // Get element at point using CDP
                var script = $@"
                    (function() {{
                        const element = document.elementFromPoint({screenPoint.X}, {screenPoint.Y});
                        if (!element) return null;

                        const rect = element.getBoundingClientRect();
                        const computed = window.getComputedStyle(element);
                        const attributes = {{}};

                        for (let attr of element.attributes) {{
                            attributes[attr.name] = attr.value;
                        }}

                        const tableInfo = getTableInfo(element);
                        const treePath = getTreePath(element);
                        const elementPath = getElementPath(element);

                        return {{
                            tagName: element.tagName,
                            id: element.id,
                            className: element.className,
                            innerHTML: element.innerHTML.substring(0, 1000),
                            outerHTML: element.outerHTML.substring(0, 1000),
                            innerText: element.innerText ? element.innerText.substring(0, 500) : '',
                            href: element.href,
                            src: element.src,
                            alt: element.alt,
                            title: element.title,
                            type: element.type,
                            name: element.name,
                            value: element.value,
                            role: element.getAttribute('role'),
                            ariaLabel: element.getAttribute('aria-label'),
                            ariaDescribedby: element.getAttribute('aria-describedby'),
                            ariaLabelledby: element.getAttribute('aria-labelledby'),
                            attributes: attributes,
                            rect: {{
                                x: rect.x,
                                y: rect.y,
                                width: rect.width,
                                height: rect.height,
                                top: rect.top,
                                left: rect.left,
                                right: rect.right,
                                bottom: rect.bottom
                            }},
                            xpath: getXPath(element),
                            cssPath: getCssPath(element),
                            treePath: treePath,
                            elementPath: elementPath,
                            tableInfo: tableInfo,
                            styles: {{
                                display: computed.display,
                                position: computed.position,
                                visibility: computed.visibility,
                                opacity: computed.opacity,
                                color: computed.color,
                                backgroundColor: computed.backgroundColor,
                                fontSize: computed.fontSize,
                                fontWeight: computed.fontWeight,
                                width: computed.width,
                                height: computed.height,
                                padding: computed.padding,
                                margin: computed.margin,
                                border: computed.border,
                                zIndex: computed.zIndex
                            }}
                        }};

                        function getXPath(element) {{
                            if (element.id) return '//*[@id=""' + element.id + '""';
                            if (!element.parentNode) return '';

                            let siblings = element.parentNode.childNodes;
                            let count = 0;
                            let siblingIndex = 0;

                            for (let sibling of siblings) {{
                                if (sibling === element) {{
                                    siblingIndex = count;
                                    break;
                                }}
                                if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {{
                                    count++;
                                }}
                            }}

                            return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() +
                                   '[' + (siblingIndex + 1) + ']';
                        }}

                        function getCssPath(element) {{
                            if (!element) return '';
                            if (element.id) return '#' + element.id;

                            let path = [];
                            while (element.parentElement) {{
                                let selector = element.tagName.toLowerCase();
                                if (element.id) {{
                                    selector += '#' + element.id;
                                    path.unshift(selector);
                                    break;
                                }}
                                if (element.className) {{
                                    selector += '.' + element.className.split(' ').join('.');
                                }}
                                path.unshift(selector);
                                element = element.parentElement;
                            }}
                            return path.join(' > ');
                        }}

                        function getTableInfo(element) {{
                            let current = element;
                            let rowIndex = -1;
                            let columnIndex = -1;
                            let tableSelector = '';

                            // Check if element is inside a table
                            while (current) {{
                                if (current.tagName === 'TD' || current.tagName === 'TH') {{
                                    columnIndex = Array.from(current.parentElement.children).indexOf(current);
                                    current = current.parentElement;
                                }}
                                if (current.tagName === 'TR') {{
                                    const tbody = current.closest('tbody') || current.closest('table');
                                    if (tbody) {{
                                        const rows = tbody.querySelectorAll('tr');
                                        rowIndex = Array.from(rows).indexOf(current);
                                    }}
                                    break;
                                }}
                                current = current.parentElement;
                            }}

                            // Build Playwright table selector
                            if (rowIndex >= 0) {{
                                const table = element.closest('table');
                                if (table) {{
                                    let selector = 'table';
                                    if (table.id) selector = 'table#' + table.id;
                                    else if (table.className) selector = 'table.' + table.className.split(' ')[0];

                                    tableSelector = selector + ` >> tr:nth-child(${{rowIndex + 1}})`;
                                    if (columnIndex >= 0) {{
                                        tableSelector += ` >> td:nth-child(${{columnIndex + 1}})`;
                                    }}
                                }}
                            }}

                            return {{
                                rowIndex: rowIndex,
                                columnIndex: columnIndex,
                                tableSelector: tableSelector
                            }};
                        }}

                        function getTreePath(element) {{
                            if (!element) return '';
                            let path = [];
                            let current = element;

                            while (current && current !== document.body) {{
                                let part = current.tagName.toLowerCase();
                                if (current.id) part += '#' + current.id;
                                else if (current.className) part += '.' + current.className.split(' ')[0];
                                path.unshift(part);
                                current = current.parentElement;
                            }}

                            return path.join(' / ');
                        }}

                        function getElementPath(element) {{
                            if (!element) return '';
                            let path = [];
                            let current = element;

                            while (current && current !== document.body) {{
                                const siblings = current.parentElement ? Array.from(current.parentElement.children) : [];
                                const index = siblings.indexOf(current);
                                let part = current.tagName.toLowerCase() + '[' + index + ']';
                                if (current.id) part += '{{id=' + current.id + '}}';
                                path.unshift(part);
                                current = current.parentElement;
                            }}

                            return path.join(' / ');
                        }}
                    }})();
                ";

                var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);

                if (result != null && result != "null")
                {
                    var elementData = JsonSerializer.Deserialize<Dictionary<string, object>>(result);
                    PopulateElementInfo(info, elementData, profile);
                }

                stopwatch.Stop();
                info.CollectionDuration = stopwatch.Elapsed;

                return info;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 GetElementAtPoint error: {ex.Message}");
                return null;
            }
        }

        public async Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile)
        {
            if (!_isInitialized || _webView?.CoreWebView2 == null)
                return new List<ElementInfo>();

            var elements = new List<ElementInfo>();

            try
            {
                // Get all elements using CDP
                var script = @"
                    (function() {
                        const elements = document.querySelectorAll('*');
                        const result = [];

                        for (let element of elements) {
                            if (element.offsetWidth > 0 && element.offsetHeight > 0) {
                                const rect = element.getBoundingClientRect();
                                result.push({
                                    tagName: element.tagName,
                                    id: element.id,
                                    className: element.className,
                                    innerText: element.innerText ? element.innerText.substring(0, 100) : '',
                                    rect: {
                                        x: rect.x,
                                        y: rect.y,
                                        width: rect.width,
                                        height: rect.height
                                    },
                                    visible: true
                                });
                            }
                        }

                        return result;
                    })();
                ";

                var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);

                if (result != null && result != "null")
                {
                    var elementsData = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(result);

                    foreach (var elementData in elementsData)
                    {
                        var info = new ElementInfo
                        {
                            DetectionMethod = Name,
                            CollectionProfile = profile.ToString(),
                            CaptureTime = DateTime.Now
                        };

                        PopulateElementInfo(info, elementData, profile);
                        elements.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 GetAllElements error: {ex.Message}");
            }

            return elements;
        }

        public async Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile)
        {
            if (!_isInitialized || _webView?.CoreWebView2 == null)
                return new List<ElementInfo>();

            var elements = new List<ElementInfo>();

            try
            {
                var script = $@"
                    (function() {{
                        const elements = document.elementsFromPoint({region.X + region.Width/2}, {region.Y + region.Height/2});
                        const result = [];

                        for (let element of elements) {{
                            const rect = element.getBoundingClientRect();
                            if (rect.x >= {region.X} && rect.y >= {region.Y} &&
                                rect.x + rect.width <= {region.X + region.Width} &&
                                rect.y + rect.height <= {region.Y + region.Height}) {{

                                result.push({{
                                    tagName: element.tagName,
                                    id: element.id,
                                    className: element.className,
                                    innerText: element.innerText ? element.innerText.substring(0, 100) : '',
                                    rect: {{
                                        x: rect.x,
                                        y: rect.y,
                                        width: rect.width,
                                        height: rect.height
                                    }}
                                }});
                            }}
                        }}

                        return result;
                    }})();
                ";

                var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);

                if (result != null && result != "null")
                {
                    var elementsData = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(result);

                    foreach (var elementData in elementsData)
                    {
                        var info = new ElementInfo
                        {
                            DetectionMethod = Name,
                            CollectionProfile = profile.ToString(),
                            CaptureTime = DateTime.Now
                        };

                        PopulateElementInfo(info, elementData, profile);
                        elements.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 GetElementsInRegion error: {ex.Message}");
            }

            return elements;
        }

        public Task<ElementInfo> GetElementTree(ElementInfo rootElement, CollectionProfile profile)
        {
            // TODO: Implement element tree building
            return Task.FromResult(rootElement);
        }

        public async Task<ElementInfo> RefreshElement(ElementInfo element, CollectionProfile profile)
        {
            if (element == null || string.IsNullOrEmpty(element.CssSelector))
                return element;

            try
            {
                var script = $@"
                    (function() {{
                        const element = document.querySelector('{element.CssSelector}');
                        if (!element) return null;

                        const rect = element.getBoundingClientRect();
                        return {{
                            tagName: element.tagName,
                            id: element.id,
                            className: element.className,
                            innerText: element.innerText ? element.innerText.substring(0, 500) : '',
                            visible: element.offsetWidth > 0 && element.offsetHeight > 0,
                            rect: {{
                                x: rect.x,
                                y: rect.y,
                                width: rect.width,
                                height: rect.height
                            }}
                        }};
                    }})();
                ";

                var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);

                if (result != null && result != "null")
                {
                    var elementData = JsonSerializer.Deserialize<Dictionary<string, object>>(result);
                    PopulateElementInfo(element, elementData, profile);
                    element.CaptureTime = DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                element.CollectionErrors.Add($"Refresh error: {ex.Message}");
            }

            return element;
        }

        private void PopulateElementInfo(ElementInfo info, Dictionary<string, object> data, CollectionProfile profile)
        {
            try
            {
                // Basic properties
                if (data.ContainsKey("tagName"))
                    info.TagName = data["tagName"]?.ToString();
                if (data.ContainsKey("id"))
                    info.HtmlId = data["id"]?.ToString();
                if (data.ContainsKey("className"))
                    info.HtmlClassName = data["className"]?.ToString();
                if (data.ContainsKey("innerText"))
                    info.InnerText = data["innerText"]?.ToString();
                if (data.ContainsKey("innerHTML"))
                    info.InnerHTML = data["innerHTML"]?.ToString();
                if (data.ContainsKey("outerHTML"))
                    info.OuterHTML = data["outerHTML"]?.ToString();

                // Links and media
                if (data.ContainsKey("href"))
                    info.Href = data["href"]?.ToString();
                if (data.ContainsKey("src"))
                    info.Src = data["src"]?.ToString();
                if (data.ContainsKey("alt"))
                    info.Alt = data["alt"]?.ToString();
                if (data.ContainsKey("title"))
                    info.Title = data["title"]?.ToString();

                // Form elements
                if (data.ContainsKey("type"))
                    info.InputType = data["type"]?.ToString();
                if (data.ContainsKey("name"))
                    info.Name = data["name"]?.ToString();
                if (data.ContainsKey("value"))
                    info.Value = data["value"]?.ToString();

                // ARIA attributes
                if (data.ContainsKey("role"))
                    info.Role = data["role"]?.ToString();
                if (data.ContainsKey("ariaLabel"))
                    info.AriaLabel = data["ariaLabel"]?.ToString();
                if (data.ContainsKey("ariaDescribedby"))
                    info.AriaDescribedBy = data["ariaDescribedby"]?.ToString();
                if (data.ContainsKey("ariaLabelledby"))
                    info.AriaLabelledBy = data["ariaLabelledby"]?.ToString();

                // Selectors
                if (data.ContainsKey("xpath"))
                    info.XPath = data["xpath"]?.ToString();
                if (data.ContainsKey("cssPath"))
                    info.CssSelector = data["cssPath"]?.ToString();
                if (data.ContainsKey("treePath"))
                    info.TreePath = data["treePath"]?.ToString();
                if (data.ContainsKey("elementPath"))
                    info.ElementPath = data["elementPath"]?.ToString();

                // Position and size
                if (data.ContainsKey("rect") && data["rect"] is JsonElement rectElement)
                {
                    var rect = rectElement.Deserialize<Dictionary<string, double>>();
                    if (rect != null)
                    {
                        info.X = rect.GetValueOrDefault("x");
                        info.Y = rect.GetValueOrDefault("y");
                        info.Width = rect.GetValueOrDefault("width");
                        info.Height = rect.GetValueOrDefault("height");
                        info.BoundingRectangle = new Rect(info.X, info.Y, info.Width, info.Height);
                    }
                }

                // Visibility
                if (data.ContainsKey("visible"))
                    info.IsVisible = Convert.ToBoolean(data["visible"]);

                // Table information
                if (data.ContainsKey("tableInfo") && data["tableInfo"] is JsonElement tableElement)
                {
                    var tableInfo = tableElement.Deserialize<Dictionary<string, object>>();
                    if (tableInfo != null)
                    {
                        if (tableInfo.ContainsKey("rowIndex"))
                        {
                            var rowIndexValue = tableInfo["rowIndex"];
                            if (rowIndexValue is JsonElement rowElement && rowElement.ValueKind == JsonValueKind.Number)
                            {
                                info.RowIndex = rowElement.GetInt32();
                            }
                        }
                        if (tableInfo.ContainsKey("columnIndex"))
                        {
                            var colIndexValue = tableInfo["columnIndex"];
                            if (colIndexValue is JsonElement colElement && colElement.ValueKind == JsonValueKind.Number)
                            {
                                info.ColumnIndex = colElement.GetInt32();
                            }
                        }
                        if (tableInfo.ContainsKey("tableSelector"))
                        {
                            info.PlaywrightTableSelector = tableInfo["tableSelector"]?.ToString();
                        }
                    }
                }

                // Attributes
                if (profile == CollectionProfile.Full && data.ContainsKey("attributes"))
                {
                    if (data["attributes"] is JsonElement attrElement)
                    {
                        var attributes = attrElement.Deserialize<Dictionary<string, string>>();
                        if (attributes != null)
                            info.HtmlAttributes = attributes;
                    }
                }

                // Styles
                if (profile == CollectionProfile.Full && data.ContainsKey("styles"))
                {
                    if (data["styles"] is JsonElement styleElement)
                    {
                        var styles = styleElement.Deserialize<Dictionary<string, string>>();
                        if (styles != null)
                            info.ComputedStyles = styles;
                    }
                }

                // Set element type based on tag name
                info.ElementType = info.TagName ?? "Unknown";
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Population error: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _webView?.Dispose();
        }
    }
}