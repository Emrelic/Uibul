using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Detects elements in Internet Explorer and embedded WebBrowser controls using MSHTML/IHTMLDocument
    /// </summary>
    public class MSHTMLDetector : IElementDetector
    {
        #region COM Interfaces

        [ComImport]
        [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IServiceProvider
        {
            [PreserveSig]
            int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
        }

        [ComImport]
        [Guid("00000100-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IEnumUnknown
        {
            [PreserveSig]
            int Next([In] uint celt, [Out, MarshalAs(UnmanagedType.LPArray)] object[] rgelt, out uint pceltFetched);
            [PreserveSig]
            int Skip([In] uint celt);
            [PreserveSig]
            int Reset();
            [PreserveSig]
            int Clone(out IEnumUnknown ppenum);
        }

        [ComImport]
        [Guid("00000117-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleContainer
        {
            void ParseDisplayName([In] IBindCtx pbc, [In, MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                out uint pchEaten, out IMoniker ppmkOut);
            void EnumObjects([In] uint grfFlags, out IEnumUnknown ppenum);
            void LockContainer([In] bool fLock);
        }

        #endregion

        #region Native Methods

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("oleacc.dll", PreserveSig = false)]
        [return: MarshalAs(UnmanagedType.Interface)]
        private static extern object AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid);

        [DllImport("user32.dll")]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
            uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private const uint SMTO_ABORTIFHUNG = 0x0002;
        private static readonly Guid IID_IHTMLDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");

        #endregion

        public string Name => "MSHTML";

        public bool CanDetect(System.Windows.Point screenPoint)
        {
            try
            {
                var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                var hwnd = WindowFromPoint(point);

                if (hwnd == IntPtr.Zero)
                    return false;

                var className = new System.Text.StringBuilder(256);
                GetClassName(hwnd, className, className.Capacity);
                var classNameStr = className.ToString();

                // Check for Internet Explorer or WebBrowser control windows
                return classNameStr.Contains("Internet Explorer") ||
                       classNameStr.Contains("IEFrame") ||
                       classNameStr.Contains("Shell DocObject View") ||
                       classNameStr.Contains("WebBrowser");
            }
            catch
            {
                return false;
            }
        }

        public async Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var stopwatch = Stopwatch.StartNew();
                    var info = new ElementInfo
                    {
                        DetectionMethod = Name,
                        CollectionProfile = profile.ToString(),
                        CaptureTime = DateTime.Now
                    };

                    var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                    var hwnd = WindowFromPoint(point);

                    if (hwnd != IntPtr.Zero)
                    {
                        var document = GetHTMLDocument(hwnd);
                        if (document != null)
                        {
                            ExtractElementFromDocument(document, screenPoint, info, profile);
                        }
                    }

                    stopwatch.Stop();
                    info.CollectionDuration = stopwatch.Elapsed;

                    return info;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MSHTML GetElementAtPoint error: {ex.Message}");
                    return null;
                }
            });
        }

        public async Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();

                try
                {
                    var document = GetHTMLDocument(windowHandle);
                    if (document != null)
                    {
                        CollectAllElements(document, elements, profile);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MSHTML GetAllElements error: {ex.Message}");
                }

                return elements;
            });
        }

        public async Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();

                try
                {
                    // Find window at region center
                    var centerPoint = new POINT
                    {
                        X = (int)(region.X + region.Width / 2),
                        Y = (int)(region.Y + region.Height / 2)
                    };
                    var hwnd = WindowFromPoint(centerPoint);

                    if (hwnd != IntPtr.Zero)
                    {
                        var document = GetHTMLDocument(hwnd);
                        if (document != null)
                        {
                            CollectElementsInRegion(document, region, elements, profile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MSHTML GetElementsInRegion error: {ex.Message}");
                }

                return elements;
            });
        }

        public Task<ElementInfo> GetElementTree(ElementInfo rootElement, CollectionProfile profile)
        {
            // TODO: Implement element tree building for MSHTML
            return Task.FromResult(rootElement);
        }

        public Task<ElementInfo> RefreshElement(ElementInfo element, CollectionProfile profile)
        {
            return Task.Run(() =>
            {
                try
                {
                    if (element == null)
                        return element;

                    // Try to find element using various methods
                    dynamic foundElement = null;
                    dynamic doc = null;

                    // Get the document from window handle if available
                    if (element.WindowHandle != IntPtr.Zero)
                    {
                        doc = GetHTMLDocument(element.WindowHandle);
                    }

                    if (doc == null)
                        return element;

                    // Method 1: Try to find by ID if available
                    if (!string.IsNullOrEmpty(element.HtmlId))
                    {
                        try
                        {
                            foundElement = doc.getElementById(element.HtmlId);
                        }
                        catch { }
                    }

                    // Method 2: Try to find by name
                    if (foundElement == null && !string.IsNullOrEmpty(element.Name))
                    {
                        try
                        {
                            dynamic elements = doc.getElementsByName(element.Name);
                            if (elements != null && elements.length > 0)
                            {
                                foundElement = elements[0];
                            }
                        }
                        catch { }
                    }

                    // Method 3: Try to find by tag name and position (using XPath-like navigation)
                    if (foundElement == null && !string.IsNullOrEmpty(element.TagName))
                    {
                        try
                        {
                            dynamic elements = doc.getElementsByTagName(element.TagName);
                            if (elements != null && element.ChildIndex >= 0 && element.ChildIndex < elements.length)
                            {
                                foundElement = elements[element.ChildIndex];
                            }
                        }
                        catch { }
                    }

                    // If element found, extract fresh info
                    if (foundElement != null)
                    {
                        var refreshedInfo = new ElementInfo
                        {
                            DetectionMethod = Name,
                            CollectionProfile = profile.ToString(),
                            CaptureTime = DateTime.Now,
                            WindowHandle = element.WindowHandle,
                            WindowTitle = element.WindowTitle,
                            WindowClassName = element.WindowClassName
                        };

                        ExtractElementInfo(foundElement, refreshedInfo, profile);
                        return refreshedInfo;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MSHTML RefreshElement error: {ex.Message}");
                }

                return element;
            });
        }

        private object GetHTMLDocument(IntPtr hwnd)
        {
            try
            {
                // Try to get IHTMLDocument2 from window
                uint msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                IntPtr result;
                SendMessageTimeout(hwnd, msg, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 1000, out result);

                if (result != IntPtr.Zero)
                {
                    object obj = null;
                    var iidHtmlDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                    var hr = ObjectFromLresult(result, ref iidHtmlDocument, IntPtr.Zero, ref obj);
                    if (hr == 0)
                    {
                        return obj;
                    }
                }

                // Try alternative method using accessibility
                var guid = new Guid("626FC520-A41E-11CF-A731-00A0C9082637"); // IID_IHTMLDocument
                return AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, ref guid);
            }
            catch
            {
                return null;
            }
        }

        [DllImport("oleacc.dll", PreserveSig = true)]
        private static extern int ObjectFromLresult(IntPtr lResult, ref Guid riid, IntPtr wParam, ref object ppvObject);

        private void ExtractElementFromDocument(object document, System.Windows.Point point, ElementInfo info, CollectionProfile profile)
        {
            try
            {
                // Use dynamic to work with COM objects
                dynamic htmlDoc = document;

                // Extract page source code if Full profile
                if (profile >= CollectionProfile.Full)
                {
                    try
                    {
                        dynamic docElement = htmlDoc.documentElement;
                        if (docElement != null)
                        {
                            info.SourceCode = docElement.outerHTML?.ToString();

                            // Also get document info
                            info.DocumentTitle = htmlDoc.title?.ToString();
                            info.DocumentUrl = htmlDoc.url?.ToString();
                            info.DocumentDomain = htmlDoc.domain?.ToString();
                            info.DocumentReadyState = htmlDoc.readyState?.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        info.CollectionErrors.Add($"Document info extraction error: {ex.Message}");
                    }
                }

                dynamic element = htmlDoc.elementFromPoint((int)point.X, (int)point.Y);

                if (element != null)
                {
                    ExtractElementInfo(element, info, profile);
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"MSHTML extraction error: {ex.Message}");
            }
        }

        private void ExtractElementInfo(dynamic element, ElementInfo info, CollectionProfile profile)
        {
            try
            {
                // Basic properties
                info.TagName = element.tagName?.ToString();
                info.HtmlId = element.id?.ToString();
                info.HtmlClassName = element.className?.ToString();

                // Name - try multiple sources for button elements
                info.Name = element.name?.ToString();
                if (string.IsNullOrEmpty(info.Name))
                {
                    // For buttons, try innerText, outerText, or value
                    try { info.Name = element.innerText?.ToString()?.Trim(); } catch { }
                    if (string.IsNullOrEmpty(info.Name))
                    {
                        try { info.Name = element.outerText?.ToString()?.Trim(); } catch { }
                    }
                    if (string.IsNullOrEmpty(info.Name))
                    {
                        try { info.Name = element.value?.ToString()?.Trim(); } catch { }
                    }
                }

                info.ElementType = info.TagName ?? "Unknown";

                // Text content
                try
                {
                    info.InnerText = element.innerText?.ToString();
                    if (info.InnerText?.Length > 500)
                        info.InnerText = info.InnerText.Substring(0, 500);
                }
                catch { }

                // HTML content
                if (profile >= CollectionProfile.Standard)
                {
                    try
                    {
                        info.InnerHTML = element.innerHTML?.ToString();
                        if (info.InnerHTML?.Length > 1000)
                            info.InnerHTML = info.InnerHTML.Substring(0, 1000);

                        info.OuterHTML = element.outerHTML?.ToString();
                        if (info.OuterHTML?.Length > 1000)
                            info.OuterHTML = info.OuterHTML.Substring(0, 1000);
                    }
                    catch { }
                }

                // Attributes
                try
                {
                    info.Href = element.href?.ToString();
                    info.Src = element.src?.ToString();
                    info.Alt = element.alt?.ToString();
                    info.Title = element.title?.ToString();
                    info.Value = element.value?.ToString();
                }
                catch { }

                // Form elements
                try
                {
                    info.InputType = element.type?.ToString();
                }
                catch { }

                // TODO: Fix dynamic property access for checked and disabled
                // Commented out temporarily to allow build
                /*
                try
                {
                    info.IsChecked = (bool)element.checked;
                }
                catch { }

                try
                {
                    info.IsDisabled = (bool)element.disabled;
                }
                catch { }
                */

                // Position and size
                try
                {
                    var rect = element.getBoundingClientRect();
                    if (rect != null)
                    {
                        info.X = rect.left;
                        info.Y = rect.top;
                        info.Width = rect.right - rect.left;
                        info.Height = rect.bottom - rect.top;
                        info.BoundingRectangle = new Rect(info.X, info.Y, info.Width, info.Height);
                    }
                }
                catch { }

                // Visibility
                try
                {
                    info.IsVisible = element.offsetWidth > 0 && element.offsetHeight > 0;
                }
                catch { }

                // Build XPath
                if (profile >= CollectionProfile.Standard)
                {
                    info.XPath = BuildXPath(element);
                }

                // Collect attributes if Full profile
                if (profile == CollectionProfile.Full)
                {
                    CollectAllAttributes(element, info);
                }

                // Document information
                try
                {
                    dynamic doc = element.document;
                    if (doc != null)
                    {
                        info.DocumentTitle = doc.title?.ToString();
                        info.DocumentUrl = doc.URL?.ToString();
                        info.DocumentDomain = doc.domain?.ToString();
                        info.DocumentReadyState = doc.readyState?.ToString();
                    }
                }
                catch { }

                // Table detection
                ExtractTableInfo(element, info);
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Element extraction error: {ex.Message}");
            }
        }

        private string BuildXPath(dynamic element)
        {
            try
            {
                if (element.id != null && !string.IsNullOrEmpty(element.id))
                {
                    return $"//*[@id='{element.id}']";
                }

                var path = new List<string>();
                dynamic current = element;

                while (current != null && current.tagName != null && current.tagName != "HTML")
                {
                    string tagName = current.tagName.ToString().ToLower();
                    int index = 1;

                    // Count siblings with same tag
                    dynamic parent = current.parentElement;
                    if (parent != null)
                    {
                        dynamic siblings = parent.children;
                        for (int i = 0; i < siblings.length; i++)
                        {
                            if (siblings[i] == current)
                                break;
                            if (siblings[i].tagName == current.tagName)
                                index++;
                        }
                    }

                    path.Insert(0, $"{tagName}[{index}]");
                    current = parent;
                }

                return "//" + string.Join("/", path);
            }
            catch
            {
                return "";
            }
        }

        private void CollectAllAttributes(dynamic element, ElementInfo info)
        {
            try
            {
                dynamic attributes = element.attributes;
                if (attributes != null)
                {
                    for (int i = 0; i < attributes.length; i++)
                    {
                        dynamic attr = attributes[i];
                        string name = attr.name?.ToString();
                        string value = attr.value?.ToString();

                        if (!string.IsNullOrEmpty(name))
                        {
                            info.HtmlAttributes[name] = value ?? "";

                            // Separate ARIA attributes
                            if (name.StartsWith("aria-"))
                            {
                                info.AriaAttributes[name] = value ?? "";
                            }
                            // Separate data attributes
                            else if (name.StartsWith("data-"))
                            {
                                info.DataAttributes[name] = value ?? "";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Attribute collection error: {ex.Message}");
            }
        }

        private void CollectAllElements(object document, List<ElementInfo> elements, CollectionProfile profile)
        {
            try
            {
                dynamic htmlDoc = document;
                dynamic allElements = htmlDoc.all;

                if (allElements != null)
                {
                    for (int i = 0; i < Math.Min(allElements.length, 1000); i++) // Limit to prevent hanging
                    {
                        try
                        {
                            dynamic element = allElements[i];
                            if (element != null)
                            {
                                var info = new ElementInfo
                                {
                                    DetectionMethod = Name,
                                    CollectionProfile = profile.ToString(),
                                    CaptureTime = DateTime.Now
                                };

                                ExtractElementInfo(element, info, profile);
                                elements.Add(info);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MSHTML CollectAllElements error: {ex.Message}");
            }
        }

        private void CollectElementsInRegion(object document, Rect region, List<ElementInfo> elements, CollectionProfile profile)
        {
            try
            {
                dynamic htmlDoc = document;
                dynamic allElements = htmlDoc.all;

                if (allElements != null)
                {
                    for (int i = 0; i < allElements.length; i++)
                    {
                        try
                        {
                            dynamic element = allElements[i];
                            if (element != null)
                            {
                                var rect = element.getBoundingClientRect();
                                if (rect != null)
                                {
                                    double left = rect.left;
                                    double top = rect.top;
                                    double right = rect.right;
                                    double bottom = rect.bottom;

                                    // Check if element is within region
                                    if (left >= region.X && top >= region.Y &&
                                        right <= region.X + region.Width &&
                                        bottom <= region.Y + region.Height)
                                    {
                                        var info = new ElementInfo
                                        {
                                            DetectionMethod = Name,
                                            CollectionProfile = profile.ToString(),
                                            CaptureTime = DateTime.Now
                                        };

                                        ExtractElementInfo(element, info, profile);
                                        elements.Add(info);
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MSHTML CollectElementsInRegion error: {ex.Message}");
            }
        }

        private void ExtractTableInfo(dynamic element, ElementInfo info)
        {
            try
            {
                dynamic current = element;
                string tagName = element.tagName?.ToString()?.ToUpper();

                // Check if element is a table cell (TD or TH)
                if (tagName == "TD" || tagName == "TH")
                {
                    info.IsTableCell = true;
                    if (tagName == "TH")
                        info.IsTableHeader = true;

                    // Get cell position
                    try
                    {
                        info.ColumnIndex = (int)current.cellIndex;
                        info.ColumnSpan = (int)(current.colSpan ?? 1);
                    }
                    catch { }

                    // Get row info
                    try
                    {
                        dynamic row = current.parentElement;
                        if (row != null && row.tagName?.ToString()?.ToUpper() == "TR")
                        {
                            info.RowIndex = (int)row.rowIndex;
                            info.RowSpan = (int)(current.rowSpan ?? 1);

                            // Get table
                            dynamic table = row.parentElement?.parentElement; // tbody or table
                            if (table != null && table.tagName?.ToString()?.ToUpper() == "TABLE")
                            {
                                ExtractTableDetails(table, info);
                            }
                            else
                            {
                                // Try to find table by going up
                                dynamic parent = row.parentElement;
                                while (parent != null)
                                {
                                    if (parent.tagName?.ToString()?.ToUpper() == "TABLE")
                                    {
                                        ExtractTableDetails(parent, info);
                                        break;
                                    }
                                    parent = parent.parentElement;
                                }
                            }
                        }
                    }
                    catch { }
                }
                // Check if element is inside a table
                else
                {
                    dynamic parent = element;
                    while (parent != null)
                    {
                        try
                        {
                            string parentTag = parent.tagName?.ToString()?.ToUpper();
                            if (parentTag == "TD" || parentTag == "TH")
                            {
                                // Element is inside a table cell
                                info.IsTableCell = true;
                                info.ColumnIndex = (int)parent.cellIndex;
                                info.ColumnSpan = (int)(parent.colSpan ?? 1);

                                dynamic row = parent.parentElement;
                                if (row != null && row.tagName?.ToString()?.ToUpper() == "TR")
                                {
                                    info.RowIndex = (int)row.rowIndex;
                                    info.RowSpan = (int)(parent.rowSpan ?? 1);
                                }
                                break;
                            }
                            else if (parentTag == "TABLE")
                            {
                                break;
                            }
                            parent = parent.parentElement;
                        }
                        catch
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Table extraction error: {ex.Message}");
            }
        }

        private void ExtractTableDetails(dynamic table, ElementInfo info)
        {
            try
            {
                // Get table name/id
                info.TableName = table.id?.ToString();
                if (string.IsNullOrEmpty(info.TableName))
                    info.TableName = table.getAttribute("name")?.ToString();
                if (string.IsNullOrEmpty(info.TableName))
                    info.TableName = table.className?.ToString();

                // Get row and column counts
                try
                {
                    dynamic rows = table.rows;
                    if (rows != null)
                        info.RowCount = rows.length;
                }
                catch { }

                try
                {
                    dynamic firstRow = table.rows[0];
                    if (firstRow != null && firstRow.cells != null)
                        info.ColumnCount = firstRow.cells.length;
                }
                catch { }

                // Get column headers from first row
                try
                {
                    dynamic thead = table.tHead;
                    if (thead != null && thead.rows != null && thead.rows.length > 0)
                    {
                        dynamic headerRow = thead.rows[0];
                        if (headerRow != null && headerRow.cells != null)
                        {
                            for (int i = 0; i < headerRow.cells.length; i++)
                            {
                                try
                                {
                                    dynamic cell = headerRow.cells[i];
                                    string headerText = cell.innerText?.ToString()?.Trim();
                                    if (!string.IsNullOrEmpty(headerText))
                                        info.ColumnHeaders.Add(headerText);
                                }
                                catch { }
                            }
                        }
                    }
                    // If no thead, check first row
                    else if (table.rows != null && table.rows.length > 0)
                    {
                        dynamic firstRow = table.rows[0];
                        if (firstRow != null && firstRow.cells != null)
                        {
                            bool hasHeaders = false;
                            // Check if first row contains TH elements
                            for (int i = 0; i < firstRow.cells.length; i++)
                            {
                                try
                                {
                                    dynamic cell = firstRow.cells[i];
                                    if (cell.tagName?.ToString()?.ToUpper() == "TH")
                                    {
                                        hasHeaders = true;
                                        string headerText = cell.innerText?.ToString()?.Trim();
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.ColumnHeaders.Add(headerText);
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }

                // Get row headers (if first column contains TH)
                try
                {
                    if (table.rows != null)
                    {
                        for (int i = 0; i < table.rows.length && i < 100; i++) // Limit to 100 rows
                        {
                            try
                            {
                                dynamic row = table.rows[i];
                                if (row != null && row.cells != null && row.cells.length > 0)
                                {
                                    dynamic firstCell = row.cells[0];
                                    if (firstCell.tagName?.ToString()?.ToUpper() == "TH")
                                    {
                                        string headerText = firstCell.innerText?.ToString()?.Trim();
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.RowHeaders.Add(headerText);
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Table details extraction error: {ex.Message}");
            }
        }
    }
}