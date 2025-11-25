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

        // CDP Deep Features
        private List<string> _consoleMessages = new List<string>();
        private List<NetworkRequest> _networkRequests = new List<NetworkRequest>();
        private PerformanceMetrics _performanceMetrics = new PerformanceMetrics();

        // Helper classes for CDP data
        private class NetworkRequest
        {
            public string Url { get; set; }
            public string Method { get; set; }
            public string ResourceType { get; set; }
            public int? StatusCode { get; set; }
            public DateTime Timestamp { get; set; }
        }

        private class PerformanceMetrics
        {
            public double DomContentLoaded { get; set; }
            public double LoadEventEnd { get; set; }
            public long JsHeapSizeUsed { get; set; }
            public long JsHeapSizeLimit { get; set; }
        }

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

                // Setup CDP event listeners
                SetupCDPEventListeners();

                _isInitialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebView2 initialization error: {ex.Message}");
                _isInitialized = false;
            }
        }

        private void SetupCDPEventListeners()
        {
            if (_webView?.CoreWebView2 == null) return;

            try
            {
                // Listen to console messages
                _webView.CoreWebView2.WebMessageReceived += (sender, args) =>
                {
                    try
                    {
                        var message = args.TryGetWebMessageAsString();
                        if (!string.IsNullOrEmpty(message))
                        {
                            _consoleMessages.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
                            Debug.WriteLine($"Console: {message}");
                        }
                    }
                    catch { }
                };

                // Inject console interceptor script
                _webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
                {
                    try
                    {
                        await _webView.CoreWebView2.ExecuteScriptAsync(@"
                            (function() {
                                const originalLog = console.log;
                                const originalError = console.error;
                                const originalWarn = console.warn;

                                console.log = function(...args) {
                                    window.chrome.webview.postMessage('LOG: ' + args.join(' '));
                                    originalLog.apply(console, args);
                                };

                                console.error = function(...args) {
                                    window.chrome.webview.postMessage('ERROR: ' + args.join(' '));
                                    originalError.apply(console, args);
                                };

                                console.warn = function(...args) {
                                    window.chrome.webview.postMessage('WARN: ' + args.join(' '));
                                    originalWarn.apply(console, args);
                                };
                            })();
                        ");

                        // Collect performance metrics
                        await CollectPerformanceMetrics();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Console interceptor error: {ex.Message}");
                    }
                };

                Debug.WriteLine("CDP event listeners configured successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CDP event listener setup error: {ex.Message}");
            }
        }

        private async Task CollectPerformanceMetrics()
        {
            try
            {
                if (_webView?.CoreWebView2 == null) return;

                var metricsJson = await _webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        const timing = performance.timing;
                        const memory = performance.memory || {};

                        return JSON.stringify({
                            domContentLoaded: timing.domContentLoadedEventEnd - timing.navigationStart,
                            loadEventEnd: timing.loadEventEnd - timing.navigationStart,
                            jsHeapSizeUsed: memory.usedJSHeapSize || 0,
                            jsHeapSizeLimit: memory.jsHeapSizeLimit || 0
                        });
                    })();
                ");

                if (!string.IsNullOrEmpty(metricsJson) && metricsJson != "null")
                {
                    var metrics = JsonSerializer.Deserialize<PerformanceMetrics>(metricsJson);
                    if (metrics != null)
                    {
                        _performanceMetrics = metrics;
                        Debug.WriteLine($"Performance: DOM={metrics.DomContentLoaded}ms, Load={metrics.LoadEventEnd}ms");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Performance metrics collection error: {ex.Message}");
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
                        const dataAttributes = {{}};
                        const ariaAttributes = {{}};

                        for (let attr of element.attributes) {{
                            attributes[attr.name] = attr.value;
                            if (attr.name.startsWith('data-')) {{
                                dataAttributes[attr.name] = attr.value;
                            }}
                            if (attr.name.startsWith('aria-')) {{
                                ariaAttributes[attr.name] = attr.value;
                            }}
                        }}

                        const tableInfo = getTableInfo(element);
                        const treePath = getTreePath(element);
                        const elementPath = getElementPath(element);
                        const pageSource = document.documentElement.outerHTML;
                        const classList = element.classList ? Array.from(element.classList) : [];

                        // Get parent info
                        const parent = element.parentElement;
                        const parentInfo = parent ? {{
                            tagName: parent.tagName,
                            id: parent.id,
                            className: parent.className
                        }} : null;

                        // Get siblings info
                        const siblings = parent ? Array.from(parent.children) : [];
                        const siblingIndex = siblings.indexOf(element);
                        const childrenCount = element.children ? element.children.length : 0;

                        return {{
                            // Basic HTML DOM Attributes
                            tagName: element.tagName,
                            pageSource: pageSource,
                            id: element.id,
                            className: element.className,
                            classList: classList,
                            htmlName: element.name,
                            innerHTML: element.innerHTML ? element.innerHTML.substring(0, 1000) : '',
                            outerHTML: element.outerHTML ? element.outerHTML.substring(0, 1000) : '',
                            innerText: element.innerText ? element.innerText.substring(0, 500) : '',
                            href: element.href,
                            src: element.src,
                            alt: element.alt,
                            title: element.title,
                            type: element.type,
                            name: element.name,
                            value: element.value,
                            placeholder: element.placeholder,
                            tabIndex: element.tabIndex,
                            checked: element.checked,
                            selected: element.selected,
                            disabled: element.disabled,
                            hidden: element.hidden,
                            readOnly: element.readOnly,
                            required: element.required,
                            attributes: attributes,
                            dataAttributes: dataAttributes,

                            // ARIA Accessibility Attributes
                            role: element.getAttribute('role'),
                            ariaLabel: element.getAttribute('aria-label'),
                            ariaDescribedby: element.getAttribute('aria-describedby'),
                            ariaLabelledby: element.getAttribute('aria-labelledby'),
                            ariaDisabled: element.getAttribute('aria-disabled'),
                            ariaHidden: element.getAttribute('aria-hidden'),
                            ariaExpanded: element.getAttribute('aria-expanded'),
                            ariaSelected: element.getAttribute('aria-selected'),
                            ariaChecked: element.getAttribute('aria-checked'),
                            ariaRequired: element.getAttribute('aria-required'),
                            ariaHasPopup: element.getAttribute('aria-haspopup'),
                            ariaLevel: element.getAttribute('aria-level'),
                            ariaValueText: element.getAttribute('aria-valuetext'),
                            ariaAttributes: ariaAttributes,

                            // Layout / Box Model Properties
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
                            clientWidth: element.clientWidth,
                            clientHeight: element.clientHeight,
                            offsetWidth: element.offsetWidth,
                            offsetHeight: element.offsetHeight,
                            scrollWidth: element.scrollWidth,
                            scrollHeight: element.scrollHeight,
                            scrollTop: element.scrollTop,
                            scrollLeft: element.scrollLeft,

                            // Selectors
                            xpath: getXPath(element),
                            cssPath: getCssPath(element),
                            treePath: treePath,
                            elementPath: elementPath,
                            tableInfo: tableInfo,

                            // CSS Computed Styles (expanded)
                            styles: {{
                                display: computed.display,
                                position: computed.position,
                                visibility: computed.visibility,
                                opacity: computed.opacity,
                                color: computed.color,
                                backgroundColor: computed.backgroundColor,
                                fontSize: computed.fontSize,
                                fontWeight: computed.fontWeight,
                                fontFamily: computed.fontFamily,
                                width: computed.width,
                                height: computed.height,
                                minWidth: computed.minWidth,
                                maxWidth: computed.maxWidth,
                                minHeight: computed.minHeight,
                                maxHeight: computed.maxHeight,
                                padding: computed.padding,
                                paddingTop: computed.paddingTop,
                                paddingRight: computed.paddingRight,
                                paddingBottom: computed.paddingBottom,
                                paddingLeft: computed.paddingLeft,
                                margin: computed.margin,
                                marginTop: computed.marginTop,
                                marginRight: computed.marginRight,
                                marginBottom: computed.marginBottom,
                                marginLeft: computed.marginLeft,
                                border: computed.border,
                                borderWidth: computed.borderWidth,
                                borderStyle: computed.borderStyle,
                                borderColor: computed.borderColor,
                                borderRadius: computed.borderRadius,
                                zIndex: computed.zIndex,
                                overflow: computed.overflow,
                                overflowX: computed.overflowX,
                                overflowY: computed.overflowY,
                                pointerEvents: computed.pointerEvents,
                                cursor: computed.cursor,
                                transform: computed.transform,
                                transition: computed.transition,
                                boxShadow: computed.boxShadow,
                                textAlign: computed.textAlign,
                                lineHeight: computed.lineHeight,
                                letterSpacing: computed.letterSpacing
                            }},

                            // Hierarchy information
                            parentInfo: parentInfo,
                            siblingIndex: siblingIndex,
                            childrenCount: childrenCount,

                            // State flags
                            isVisible: element.offsetWidth > 0 && element.offsetHeight > 0,
                            isFocused: document.activeElement === element,
                            isContentEditable: element.isContentEditable
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
                // === 3.1 Basic HTML DOM Attributes ===
                if (data.ContainsKey("tagName"))
                    info.TagName = data["tagName"]?.ToString();
                if (data.ContainsKey("id"))
                    info.HtmlId = data["id"]?.ToString();
                if (data.ContainsKey("className"))
                    info.HtmlClassName = data["className"]?.ToString();
                if (data.ContainsKey("htmlName"))
                    info.HtmlName = data["htmlName"]?.ToString();
                if (data.ContainsKey("innerText"))
                    info.InnerText = data["innerText"]?.ToString();
                if (data.ContainsKey("innerHTML"))
                    info.InnerHTML = data["innerHTML"]?.ToString();
                if (data.ContainsKey("outerHTML"))
                    info.OuterHTML = data["outerHTML"]?.ToString();

                // Class list
                if (data.ContainsKey("classList") && data["classList"] is JsonElement classListElement)
                {
                    try
                    {
                        var classList = classListElement.Deserialize<List<string>>();
                        if (classList != null)
                            info.ClassList = classList;
                    }
                    catch { }
                }

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
                {
                    info.Value = data["value"]?.ToString();
                    info.InputValue = data["value"]?.ToString();
                }
                if (data.ContainsKey("placeholder"))
                    info.Placeholder = data["placeholder"]?.ToString();

                // TabIndex
                if (data.ContainsKey("tabIndex"))
                {
                    try
                    {
                        if (data["tabIndex"] is JsonElement tabElement && tabElement.ValueKind == JsonValueKind.Number)
                            info.TabIndex = tabElement.GetInt32();
                    }
                    catch { }
                }

                // State flags from HTML
                if (data.ContainsKey("checked"))
                    info.IsChecked = ConvertToBool(data["checked"]);
                if (data.ContainsKey("selected"))
                    info.IsSelected = ConvertToBool(data["selected"]);
                if (data.ContainsKey("disabled"))
                    info.IsDisabled = ConvertToBool(data["disabled"]);
                if (data.ContainsKey("hidden"))
                    info.IsHidden = ConvertToBool(data["hidden"]);
                if (data.ContainsKey("readOnly"))
                    info.ValuePattern_IsReadOnly = ConvertToBool(data["readOnly"]);
                if (data.ContainsKey("required"))
                    info.IsRequiredForForm = ConvertToBool(data["required"]);
                if (data.ContainsKey("isContentEditable"))
                    info.IsEditable = ConvertToBool(data["isContentEditable"]);

                // === 3.2 ARIA Accessibility Attributes ===
                if (data.ContainsKey("role"))
                {
                    info.Role = data["role"]?.ToString();
                    info.AriaRole = data["role"]?.ToString();
                }
                if (data.ContainsKey("ariaLabel"))
                    info.AriaLabel = data["ariaLabel"]?.ToString();
                if (data.ContainsKey("ariaDescribedby"))
                    info.AriaDescribedBy = data["ariaDescribedby"]?.ToString();
                if (data.ContainsKey("ariaLabelledby"))
                    info.AriaLabelledBy = data["ariaLabelledby"]?.ToString();
                if (data.ContainsKey("ariaDisabled"))
                    info.AriaDisabled = ConvertToNullableBool(data["ariaDisabled"]);
                if (data.ContainsKey("ariaHidden"))
                    info.AriaHidden = ConvertToNullableBool(data["ariaHidden"]);
                if (data.ContainsKey("ariaExpanded"))
                    info.AriaExpanded = ConvertToNullableBool(data["ariaExpanded"]);
                if (data.ContainsKey("ariaSelected"))
                    info.AriaSelected = ConvertToNullableBool(data["ariaSelected"]);
                if (data.ContainsKey("ariaChecked"))
                    info.AriaChecked = ConvertToNullableBool(data["ariaChecked"]);
                if (data.ContainsKey("ariaRequired"))
                    info.AriaRequired = ConvertToNullableBool(data["ariaRequired"]);
                if (data.ContainsKey("ariaHasPopup"))
                    info.AriaHasPopup = data["ariaHasPopup"]?.ToString();
                if (data.ContainsKey("ariaLevel"))
                {
                    try
                    {
                        if (data["ariaLevel"] is JsonElement levelElement && levelElement.ValueKind == JsonValueKind.Number)
                            info.AriaLevel = levelElement.GetInt32();
                    }
                    catch { }
                }
                if (data.ContainsKey("ariaValueText"))
                    info.AriaValueText = data["ariaValueText"]?.ToString();

                // ARIA attributes dictionary
                if (data.ContainsKey("ariaAttributes") && data["ariaAttributes"] is JsonElement ariaElement)
                {
                    try
                    {
                        var ariaAttrs = ariaElement.Deserialize<Dictionary<string, string>>();
                        if (ariaAttrs != null)
                            info.AriaAttributes = ariaAttrs;
                    }
                    catch { }
                }

                // === 3.3 Layout / Box Model Properties ===
                if (data.ContainsKey("clientWidth"))
                    info.ClientWidth = ConvertToDouble(data["clientWidth"]);
                if (data.ContainsKey("clientHeight"))
                    info.ClientHeight = ConvertToDouble(data["clientHeight"]);
                if (data.ContainsKey("offsetWidth"))
                    info.OffsetWidth = ConvertToDouble(data["offsetWidth"]);
                if (data.ContainsKey("offsetHeight"))
                    info.OffsetHeight = ConvertToDouble(data["offsetHeight"]);
                if (data.ContainsKey("scrollWidth"))
                    info.ScrollWidth = ConvertToDouble(data["scrollWidth"]);
                if (data.ContainsKey("scrollHeight"))
                    info.ScrollHeight = ConvertToDouble(data["scrollHeight"]);
                if (data.ContainsKey("scrollTop"))
                    info.ScrollTop = ConvertToDouble(data["scrollTop"]);
                if (data.ContainsKey("scrollLeft"))
                    info.ScrollLeft = ConvertToDouble(data["scrollLeft"]);

                // Page source code (Full profile)
                if (profile >= CollectionProfile.Full && data.ContainsKey("pageSource"))
                {
                    info.SourceCode = data["pageSource"]?.ToString();
                }

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

                // Visibility / state
                if (data.ContainsKey("isVisible"))
                    info.IsVisible = ConvertToBool(data["isVisible"]);
                if (data.ContainsKey("isFocused"))
                    info.IsFocused = ConvertToBool(data["isFocused"]);

                // Hierarchy
                if (data.ContainsKey("parentInfo") && data["parentInfo"] is JsonElement parentElement)
                {
                    try
                    {
                        var parentInfo = parentElement.Deserialize<Dictionary<string, string>>();
                        if (parentInfo != null)
                        {
                            info.ParentName = parentInfo.GetValueOrDefault("tagName");
                            info.ParentId = parentInfo.GetValueOrDefault("id");
                            info.ParentClassName = parentInfo.GetValueOrDefault("className");
                        }
                    }
                    catch { }
                }
                if (data.ContainsKey("siblingIndex"))
                {
                    try
                    {
                        if (data["siblingIndex"] is JsonElement sibElement && sibElement.ValueKind == JsonValueKind.Number)
                            info.ChildIndex = sibElement.GetInt32();
                    }
                    catch { }
                }

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
                                if (info.RowIndex >= 0)
                                    info.IsTableCell = true;
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
                if (data.ContainsKey("attributes") && data["attributes"] is JsonElement attrElement)
                {
                    try
                    {
                        var attributes = attrElement.Deserialize<Dictionary<string, string>>();
                        if (attributes != null)
                            info.HtmlAttributes = attributes;
                    }
                    catch { }
                }

                // Data attributes
                if (data.ContainsKey("dataAttributes") && data["dataAttributes"] is JsonElement dataAttrElement)
                {
                    try
                    {
                        var dataAttrs = dataAttrElement.Deserialize<Dictionary<string, string>>();
                        if (dataAttrs != null)
                            info.DataAttributes = dataAttrs;
                    }
                    catch { }
                }

                // === 3.4 CSS Computed Styles ===
                if (data.ContainsKey("styles") && data["styles"] is JsonElement styleElement)
                {
                    try
                    {
                        var styles = styleElement.Deserialize<Dictionary<string, string>>();
                        if (styles != null)
                        {
                            info.ComputedStyles = styles;

                            // Extract specific style properties
                            info.Style_Display = styles.GetValueOrDefault("display");
                            info.Style_Position = styles.GetValueOrDefault("position");
                            info.Style_Visibility = styles.GetValueOrDefault("visibility");
                            info.Style_Opacity = styles.GetValueOrDefault("opacity");
                            info.Style_Color = styles.GetValueOrDefault("color");
                            info.Style_BackgroundColor = styles.GetValueOrDefault("backgroundColor");
                            info.Style_FontSize = styles.GetValueOrDefault("fontSize");
                            info.Style_FontWeight = styles.GetValueOrDefault("fontWeight");
                            info.Style_ZIndex = styles.GetValueOrDefault("zIndex");
                            info.Style_PointerEvents = styles.GetValueOrDefault("pointerEvents");
                            info.Style_Overflow = styles.GetValueOrDefault("overflow");
                            info.Style_Transform = styles.GetValueOrDefault("transform");

                            // Box model
                            info.BoxModel_Margin = styles.GetValueOrDefault("margin");
                            info.BoxModel_Padding = styles.GetValueOrDefault("padding");
                            info.BoxModel_Border = styles.GetValueOrDefault("border");
                        }
                    }
                    catch { }
                }

                // Set element type based on tag name
                info.ElementType = info.TagName ?? "Unknown";
                info.Tag = info.TagName;

                // Add CDP Deep Features to CustomProperties (Full profile only)
                if (profile == CollectionProfile.Full)
                {
                    // Add console messages
                    if (_consoleMessages.Count > 0)
                    {
                        info.CDP_ConsoleMessages = string.Join("\n", _consoleMessages.TakeLast(50));
                        info.CustomProperties["CDP_ConsoleMessages"] = info.CDP_ConsoleMessages;
                    }

                    // Add performance metrics
                    if (_performanceMetrics != null)
                    {
                        info.CDP_DOMContentLoaded = _performanceMetrics.DomContentLoaded;
                        info.CDP_LoadEventEnd = _performanceMetrics.LoadEventEnd;
                        info.CDP_JSHeapUsed = _performanceMetrics.JsHeapSizeUsed;
                        info.CDP_JSHeapLimit = _performanceMetrics.JsHeapSizeLimit;

                        info.CustomProperties["CDP_Performance_DOMContentLoaded"] = $"{_performanceMetrics.DomContentLoaded}ms";
                        info.CustomProperties["CDP_Performance_LoadEventEnd"] = $"{_performanceMetrics.LoadEventEnd}ms";
                        info.CustomProperties["CDP_Performance_JSHeapUsed"] = $"{_performanceMetrics.JsHeapSizeUsed / 1024 / 1024:F2}MB";
                        info.CustomProperties["CDP_Performance_JSHeapLimit"] = $"{_performanceMetrics.JsHeapSizeLimit / 1024 / 1024:F2}MB";
                    }

                    // Add network requests summary
                    if (_networkRequests.Count > 0)
                    {
                        info.CDP_NetworkRequestsCount = _networkRequests.Count;
                        info.CustomProperties["CDP_NetworkRequests_Count"] = _networkRequests.Count;
                        info.CustomProperties["CDP_NetworkRequests_Failed"] = _networkRequests.Count(r => r.StatusCode >= 400);
                    }
                }

                // Mark technology used
                if (!info.TechnologiesUsed.Contains("WebView2/CDP"))
                    info.TechnologiesUsed.Add("WebView2/CDP");
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Population error: {ex.Message}");
            }
        }

        private bool ConvertToBool(object value)
        {
            if (value == null) return false;
            if (value is bool b) return b;
            if (value is JsonElement je)
            {
                if (je.ValueKind == JsonValueKind.True) return true;
                if (je.ValueKind == JsonValueKind.False) return false;
                if (je.ValueKind == JsonValueKind.String)
                    return je.GetString()?.ToLower() == "true";
            }
            return Convert.ToBoolean(value);
        }

        private bool? ConvertToNullableBool(object value)
        {
            if (value == null) return null;
            if (value is JsonElement je && je.ValueKind == JsonValueKind.Null) return null;
            return ConvertToBool(value);
        }

        private double ConvertToDouble(object value)
        {
            if (value == null) return 0;
            if (value is JsonElement je && je.ValueKind == JsonValueKind.Number)
                return je.GetDouble();
            return Convert.ToDouble(value);
        }

        public void Dispose()
        {
            _webView?.Dispose();
        }
    }
}