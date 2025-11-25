using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Playwright;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Playwright-based element detector for web browsers
    /// Uses Microsoft.Playwright for browser automation
    /// </summary>
    public class PlaywrightDetector : IElementDetector, IDisposable
    {
        private IPlaywright _playwright;
        private IBrowser _browser;
        private IPage _currentPage;
        private bool _isInitialized;
        private readonly object _lockObject = new object();

        public string Name => "Playwright";

        public PlaywrightDetector()
        {
            // Lazy initialization - browser will be started when first needed
        }

        private async Task InitializeAsync()
        {
            if (_isInitialized)
                return;

            lock (_lockObject)
            {
                if (_isInitialized)
                    return;

                try
                {
                    // This will be done synchronously in first call
                    Task.Run(async () =>
                    {
                        _playwright = await Microsoft.Playwright.Playwright.CreateAsync();
                        _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                        {
                            Headless = false, // Visible browser for inspection
                            Args = new[] { "--start-maximized" }
                        });

                        // Get or create a page
                        var pages = _browser.Contexts.SelectMany(c => c.Pages).ToList();
                        if (pages.Any())
                        {
                            _currentPage = pages.First();
                        }
                        else
                        {
                            var context = await _browser.NewContextAsync();
                            _currentPage = await context.NewPageAsync();
                        }

                        _isInitialized = true;
                    }).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Playwright initialization error: {ex.Message}");
                    _isInitialized = false;
                }
            }
        }

        public bool CanDetect(System.Windows.Point screenPoint)
        {
            // Playwright can detect elements in any open browser window
            // For now, we'll always return true if initialized
            if (!_isInitialized)
            {
                try
                {
                    InitializeAsync().Wait();
                }
                catch
                {
                    return false;
                }
            }

            return _isInitialized && _currentPage != null;
        }

        public async Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile)
        {
            if (!_isInitialized)
            {
                await InitializeAsync();
            }

            if (!_isInitialized || _currentPage == null)
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

                // Convert screen coordinates to viewport coordinates
                // Note: This is simplified - real implementation would need window position
                var viewportPoint = new { x = (int)screenPoint.X, y = (int)screenPoint.Y };

                // Execute script to find element at point
                var elementData = await _currentPage.EvaluateAsync<dynamic>($@"
                    (() => {{
                        const element = document.elementFromPoint({viewportPoint.x}, {viewportPoint.y});
                        if (!element) return null;

                        const rect = element.getBoundingClientRect();
                        const computed = window.getComputedStyle(element);
                        const attributes = {{}};

                        for (let attr of element.attributes) {{
                            attributes[attr.name] = attr.value;
                        }}

                        // Get Playwright selector
                        const getPlaywrightSelector = (el) => {{
                            if (el.id) return '#' + el.id;

                            let selector = el.tagName.toLowerCase();
                            if (el.className) {{
                                const classes = el.className.split(' ').filter(c => c).join('.');
                                if (classes) selector += '.' + classes;
                            }}

                            // Add nth-child if needed
                            const parent = el.parentElement;
                            if (parent) {{
                                const siblings = Array.from(parent.children).filter(s => s.tagName === el.tagName);
                                if (siblings.length > 1) {{
                                    const index = siblings.indexOf(el) + 1;
                                    selector += `:nth-child(${{index}})`;
                                }}
                            }}

                            return selector;
                        }};

                        // Build full selector path
                        const getFullSelector = (el) => {{
                            const path = [];
                            let current = el;
                            while (current && current !== document.body) {{
                                path.unshift(getPlaywrightSelector(current));
                                current = current.parentElement;
                            }}
                            return path.join(' > ');
                        }};

                        // Get table info
                        const getTableInfo = (el) => {{
                            let current = el;
                            while (current) {{
                                if (current.tagName === 'TD' || current.tagName === 'TH') {{
                                    const row = current.closest('tr');
                                    const table = current.closest('table');
                                    if (row && table) {{
                                        const rows = Array.from(table.querySelectorAll('tr'));
                                        const cells = Array.from(row.children);
                                        return {{
                                            rowIndex: rows.indexOf(row),
                                            columnIndex: cells.indexOf(current),
                                            tableSelector: getFullSelector(table) + ` >> tr:nth-child(${{rows.indexOf(row) + 1}}) >> td:nth-child(${{cells.indexOf(current) + 1}})`
                                        }};
                                    }}
                                }}
                                current = current.parentElement;
                            }}
                            return {{ rowIndex: -1, columnIndex: -1, tableSelector: '' }};
                        }};

                        const tableInfo = getTableInfo(element);
                        const pageSource = document.documentElement.outerHTML;

                        return {{
                            tagName: element.tagName,
                            id: element.id,
                            className: element.className,
                            name: element.name,
                            value: element.value,
                            innerText: element.innerText ? element.innerText.substring(0, 500) : '',
                            innerHTML: element.innerHTML ? element.innerHTML.substring(0, 1000) : '',
                            outerHTML: element.outerHTML ? element.outerHTML.substring(0, 1000) : '',
                            href: element.href,
                            src: element.src,
                            alt: element.alt,
                            title: element.title,
                            type: element.type,
                            role: element.getAttribute('role'),
                            ariaLabel: element.getAttribute('aria-label'),
                            playwrightSelector: getFullSelector(element),
                            tableInfo: tableInfo,
                            pageSource: pageSource,
                            attributes: attributes,
                            rect: {{
                                x: rect.x,
                                y: rect.y,
                                width: rect.width,
                                height: rect.height
                            }},
                            isVisible: computed.visibility !== 'hidden' && computed.display !== 'none' && element.offsetWidth > 0,
                            isEnabled: !element.disabled,
                            isChecked: element.checked || false,
                            isEditable: !element.readOnly && !element.disabled,
                            styles: {{
                                display: computed.display,
                                visibility: computed.visibility,
                                position: computed.position,
                                color: computed.color,
                                backgroundColor: computed.backgroundColor,
                                fontSize: computed.fontSize
                            }}
                        }};
                    }})()
                ");

                if (elementData == null)
                    return null;

                // Populate ElementInfo from script result
                PopulateElementInfo(info, elementData, profile);

                stopwatch.Stop();
                info.CollectionDuration = stopwatch.Elapsed;

                return info;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Playwright GetElementAtPoint error: {ex.Message}");
                return null;
            }
        }

        private void PopulateElementInfo(ElementInfo info, dynamic data, CollectionProfile profile)
        {
            try
            {
                // Basic properties
                info.TagName = data.tagName?.ToString();
                info.HtmlId = data.id?.ToString();
                info.HtmlClassName = data.className?.ToString();
                info.Name = data.name?.ToString();
                info.Value = data.value?.ToString();
                info.ElementType = info.TagName ?? "Unknown";

                // Text content
                info.InnerText = data.innerText?.ToString();
                info.InnerHTML = data.innerHTML?.ToString();
                info.OuterHTML = data.outerHTML?.ToString();

                // Links and media
                info.Href = data.href?.ToString();
                info.Src = data.src?.ToString();
                info.Alt = data.alt?.ToString();
                info.Title = data.title?.ToString();
                info.InputType = data.type?.ToString();

                // ARIA
                info.Role = data.role?.ToString();
                info.AriaLabel = data.ariaLabel?.ToString();

                // Playwright selector
                info.PlaywrightSelector = data.playwrightSelector?.ToString();

                // Page source (Full profile)
                if (profile >= CollectionProfile.Full)
                {
                    info.SourceCode = data.pageSource?.ToString();
                }

                // Position and size
                if (data.rect != null)
                {
                    info.X = data.rect.x;
                    info.Y = data.rect.y;
                    info.Width = data.rect.width;
                    info.Height = data.rect.height;
                    info.BoundingRectangle = new Rect(info.X, info.Y, info.Width, info.Height);
                }

                // Table info
                if (data.tableInfo != null)
                {
                    try
                    {
                        var ri = data.tableInfo.rowIndex;
                        info.RowIndex = ri != null ? (int)ri : -1;
                    }
                    catch { info.RowIndex = -1; }
                    try
                    {
                        var ci = data.tableInfo.columnIndex;
                        info.ColumnIndex = ci != null ? (int)ci : -1;
                    }
                    catch { info.ColumnIndex = -1; }
                    info.PlaywrightTableSelector = data.tableInfo.tableSelector?.ToString();
                }

                // Boolean properties
                info.IsVisible = data.isVisible ?? false;
                info.IsEnabled = data.isEnabled ?? true;
                info.IsChecked = data.isChecked ?? false;
                info.IsEditable = data.isEditable ?? false;

                // Attributes (Full profile)
                if (profile >= CollectionProfile.Full && data.attributes != null)
                {
                    foreach (var attr in data.attributes)
                    {
                        info.HtmlAttributes[attr.Name] = attr.Value?.ToString();
                    }
                }

                // Styles (Full profile)
                if (profile >= CollectionProfile.Full && data.styles != null)
                {
                    info.ComputedStyles["display"] = data.styles.display?.ToString();
                    info.ComputedStyles["visibility"] = data.styles.visibility?.ToString();
                    info.ComputedStyles["position"] = data.styles.position?.ToString();
                    info.ComputedStyles["color"] = data.styles.color?.ToString();
                    info.ComputedStyles["backgroundColor"] = data.styles.backgroundColor?.ToString();
                    info.ComputedStyles["fontSize"] = data.styles.fontSize?.ToString();
                }

                // Document info
                info.DocumentUrl = _currentPage?.Url;
                info.DocumentTitle = Task.Run(async () => await _currentPage.TitleAsync()).Result;
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Playwright property extraction error: {ex.Message}");
            }
        }

        public async Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile)
        {
            if (!_isInitialized)
            {
                await InitializeAsync();
            }

            if (!_isInitialized || _currentPage == null)
                return new List<ElementInfo>();

            var elements = new List<ElementInfo>();

            try
            {
                // Get all interactive elements on page
                var locator = _currentPage.Locator("a, button, input, select, textarea, [role='button'], [onclick]");
                var count = await locator.CountAsync();

                for (int i = 0; i < Math.Min(count, 1000); i++) // Limit to 1000 elements
                {
                    try
                    {
                        var element = locator.Nth(i);
                        var box = await element.BoundingBoxAsync();

                        if (box != null)
                        {
                            var point = new System.Windows.Point(box.X + box.Width / 2, box.Y + box.Height / 2);
                            var elementInfo = await GetElementAtPoint(point, profile);
                            if (elementInfo != null)
                            {
                                elements.Add(elementInfo);
                            }
                        }
                    }
                    catch
                    {
                        // Skip elements that can't be inspected
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Playwright GetAllElements error: {ex.Message}");
            }

            return elements;
        }

        public async Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile)
        {
            if (!_isInitialized)
            {
                await InitializeAsync();
            }

            if (!_isInitialized || _currentPage == null)
                return new List<ElementInfo>();

            var elements = new List<ElementInfo>();

            try
            {
                // Get all elements
                var allElements = await GetAllElements(IntPtr.Zero, profile);

                // Filter by region
                foreach (var element in allElements)
                {
                    if (region.Contains(element.X, element.Y))
                    {
                        elements.Add(element);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Playwright GetElementsInRegion error: {ex.Message}");
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
            if (!_isInitialized || _currentPage == null || element == null)
                return element;

            try
            {
                // Try to find element using Playwright selector
                string selector = null;

                // Priority 1: Use PlaywrightSelector if available
                if (!string.IsNullOrEmpty(element.PlaywrightSelector))
                {
                    selector = element.PlaywrightSelector;
                }
                // Priority 2: Use ID
                else if (!string.IsNullOrEmpty(element.HtmlId))
                {
                    selector = $"#{element.HtmlId}";
                }
                // Priority 3: Use tag name with class
                else if (!string.IsNullOrEmpty(element.TagName))
                {
                    selector = element.TagName.ToLower();
                    if (!string.IsNullOrEmpty(element.HtmlClassName))
                    {
                        var classes = element.HtmlClassName.Split(' ').Where(c => !string.IsNullOrWhiteSpace(c));
                        selector += "." + string.Join(".", classes);
                    }
                }

                if (string.IsNullOrEmpty(selector))
                    return element;

                // Execute script to get element data
                var elementData = await _currentPage.EvaluateAsync<dynamic>($@"
                    (() => {{
                        const element = document.querySelector('{selector}');
                        if (!element) return null;

                        const rect = element.getBoundingClientRect();
                        const computed = window.getComputedStyle(element);
                        const attributes = {{}};

                        for (let attr of element.attributes) {{
                            attributes[attr.name] = attr.value;
                        }}

                        // Get full selector path
                        const getFullSelector = (el) => {{
                            const path = [];
                            let current = el;
                            while (current && current !== document.body) {{
                                let selector = current.tagName.toLowerCase();
                                if (current.id) {{
                                    selector = '#' + current.id;
                                    path.unshift(selector);
                                    break;
                                }}
                                if (current.className) {{
                                    const classes = current.className.split(' ').filter(c => c).join('.');
                                    if (classes) selector += '.' + classes;
                                }}
                                path.unshift(selector);
                                current = current.parentElement;
                            }}
                            return path.join(' > ');
                        }};

                        // Get table info
                        const getTableInfo = (el) => {{
                            let current = el;
                            while (current) {{
                                if (current.tagName === 'TD' || current.tagName === 'TH') {{
                                    const row = current.closest('tr');
                                    const table = current.closest('table');
                                    if (row && table) {{
                                        const rows = Array.from(table.querySelectorAll('tr'));
                                        const cells = Array.from(row.children);
                                        return {{
                                            rowIndex: rows.indexOf(row),
                                            columnIndex: cells.indexOf(current)
                                        }};
                                    }}
                                }}
                                current = current.parentElement;
                            }}
                            return {{ rowIndex: -1, columnIndex: -1 }};
                        }};

                        const tableInfo = getTableInfo(element);

                        return {{
                            tagName: element.tagName,
                            id: element.id,
                            className: element.className,
                            name: element.name,
                            value: element.value,
                            innerText: element.innerText ? element.innerText.substring(0, 500) : '',
                            innerHTML: element.innerHTML ? element.innerHTML.substring(0, 1000) : '',
                            outerHTML: element.outerHTML ? element.outerHTML.substring(0, 1000) : '',
                            href: element.href,
                            src: element.src,
                            alt: element.alt,
                            title: element.title,
                            type: element.type,
                            role: element.getAttribute('role'),
                            ariaLabel: element.getAttribute('aria-label'),
                            playwrightSelector: getFullSelector(element),
                            tableInfo: tableInfo,
                            attributes: attributes,
                            rect: {{
                                x: rect.x,
                                y: rect.y,
                                width: rect.width,
                                height: rect.height
                            }},
                            isVisible: computed.visibility !== 'hidden' && computed.display !== 'none' && element.offsetWidth > 0,
                            isEnabled: !element.disabled,
                            isChecked: element.checked || false,
                            isEditable: !element.readOnly && !element.disabled,
                            styles: {{
                                display: computed.display,
                                visibility: computed.visibility,
                                position: computed.position,
                                color: computed.color,
                                backgroundColor: computed.backgroundColor,
                                fontSize: computed.fontSize
                            }}
                        }};
                    }})()
                ");

                if (elementData == null)
                    return element;

                // Create refreshed ElementInfo
                var refreshedInfo = new ElementInfo
                {
                    DetectionMethod = Name,
                    CollectionProfile = profile.ToString(),
                    CaptureTime = DateTime.Now
                };

                PopulateElementInfo(refreshedInfo, elementData, profile);
                return refreshedInfo;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Playwright RefreshElement error: {ex.Message}");
                return element;
            }
        }

        public void Dispose()
        {
            try
            {
                _currentPage?.CloseAsync().Wait();
                _browser?.CloseAsync().Wait();
                _browser?.DisposeAsync().AsTask().Wait();
                _playwright?.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Playwright disposal error: {ex.Message}");
            }
        }
    }
}
