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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

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

                // Debug: Log the class name to help identify the window type
                Debug.WriteLine($"MSHTML CanDetect - Window Class Name: '{classNameStr}'");

                // Check for Internet Explorer or WebBrowser control windows
                // "Internet Explorer_Server" is the actual class name for IE embedded browser content
                bool canDetect = classNameStr == "Internet Explorer_Server" ||  // Exact match for IE content area
                       classNameStr.Contains("Internet Explorer") ||
                       classNameStr.Contains("IEFrame") ||
                       classNameStr.Contains("Shell DocObject View") ||
                       classNameStr.Contains("Shell Embedding") ||
                       classNameStr.Contains("WebBrowser") ||
                       classNameStr.Contains("HTML") ||
                       classNameStr.Contains("Browser") ||
                       classNameStr.Contains("Mozilla") ||
                       classNameStr.Contains("Chrome") ||
                       classNameStr.Contains("Trident");

                Debug.WriteLine($"MSHTML CanDetect - Class='{classNameStr}', Result={canDetect}");
                return canDetect;
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
                var stopwatch = Stopwatch.StartNew();
                var info = new ElementInfo
                {
                    DetectionMethod = Name,
                    CollectionProfile = profile.ToString(),
                    CaptureTime = DateTime.Now
                };

                try
                {
                    var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                    var hwnd = WindowFromPoint(point);

                    info.CollectionErrors.Add($"[MSHTML] GetElementAtPoint started at ({screenPoint.X}, {screenPoint.Y})");
                    info.CollectionErrors.Add($"[MSHTML] Window handle: 0x{hwnd.ToInt64():X8}");

                    // Get window class name
                    var className = new System.Text.StringBuilder(256);
                    GetClassName(hwnd, className, className.Capacity);
                    info.CollectionErrors.Add($"[MSHTML] Window class: '{className}'");

                    if (hwnd != IntPtr.Zero)
                    {
                        var document = GetHTMLDocument(hwnd);
                        if (document != null)
                        {
                            info.CollectionErrors.Add($"[MSHTML] Got HTML document successfully, extracting element...");
                            ExtractElementFromDocument(document, screenPoint, info, profile);
                        }
                        else
                        {
                            info.CollectionErrors.Add($"[MSHTML] ERROR: GetHTMLDocument returned NULL");
                            info.CollectionErrors.Add($"[MSHTML] Details: {_lastDocumentError}");
                        }
                    }
                    else
                    {
                        info.CollectionErrors.Add($"[MSHTML] ERROR: WindowFromPoint returned NULL");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MSHTML GetElementAtPoint error: {ex.Message}");
                    info.CollectionErrors.Add($"[MSHTML] Critical error: {ex.Message}");
                }

                stopwatch.Stop();
                info.CollectionDuration = stopwatch.Elapsed;

                return info;
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
            Debug.WriteLine($"[MSHTML] GetHTMLDocument called for hwnd=0x{hwnd.ToInt64():X8}");
            _lastDocumentError = "";

            // Method 1: WM_HTML_GETOBJECT message (standard IE method)
            try
            {
                uint msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                Debug.WriteLine($"[MSHTML] Method 1: WM_HTML_GETOBJECT message={msg}");

                IntPtr result;
                var sendResult = SendMessageTimeout(hwnd, msg, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out result);
                Debug.WriteLine($"[MSHTML] SendMessageTimeout: sendResult=0x{sendResult.ToInt64():X}, lResult=0x{result.ToInt64():X}");

                if (result != IntPtr.Zero)
                {
                    object obj;
                    var iidHtmlDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                    var hr = ObjectFromLresult(result, ref iidHtmlDocument, IntPtr.Zero, out obj);
                    Debug.WriteLine($"[MSHTML] ObjectFromLresult: hr={hr}, obj={obj != null}");

                    if (hr == 0 && obj != null)
                    {
                        Debug.WriteLine($"[MSHTML] SUCCESS via WM_HTML_GETOBJECT");
                        return obj;
                    }
                    else
                    {
                        _lastDocumentError += $"ObjectFromLresult hr={hr}; ";
                    }
                }
                else
                {
                    _lastDocumentError += $"WM_HTML_GETOBJECT returned 0; ";
                }
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method1: {ex.Message}; ";
                Debug.WriteLine($"[MSHTML] Method 1 exception: {ex.Message}");
            }

            // Method 2: Try with different IID (IHTMLDocument3)
            try
            {
                uint msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                IntPtr result;
                SendMessageTimeout(hwnd, msg, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out result);

                if (result != IntPtr.Zero)
                {
                    object obj;
                    // Try IHTMLDocument3
                    var iidHtmlDocument3 = new Guid("3050F485-98B5-11CF-BB82-00AA00BDCE0B");
                    var hr = ObjectFromLresult(result, ref iidHtmlDocument3, IntPtr.Zero, out obj);
                    if (hr == 0 && obj != null)
                    {
                        Debug.WriteLine($"[MSHTML] SUCCESS via IHTMLDocument3");
                        return obj;
                    }
                }
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method2: {ex.Message}; ";
            }

            // Method 3: AccessibleObjectFromWindow with OBJID_NATIVEOM
            try
            {
                Debug.WriteLine($"[MSHTML] Method 3: OBJID_NATIVEOM");
                var guid = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                var obj = AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, ref guid);
                if (obj != null)
                {
                    Debug.WriteLine($"[MSHTML] SUCCESS via OBJID_NATIVEOM");
                    return obj;
                }
                _lastDocumentError += "OBJID_NATIVEOM returned null; ";
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method3: {ex.Message}; ";
                Debug.WriteLine($"[MSHTML] Method 3 exception: {ex.Message}");
            }

            // Method 4: Use UI Automation to get to the document
            try
            {
                Debug.WriteLine($"[MSHTML] Method 4: UI Automation");
                var element = AutomationElement.FromHandle(hwnd);
                if (element != null)
                {
                    // Try to get the document pattern or navigate to get COM object
                    object pattern;
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                    {
                        Debug.WriteLine($"[MSHTML] Got ValuePattern from UI Automation");
                    }

                    // Try to access native element
                    var nativeHandle = element.Current.NativeWindowHandle;
                    if (nativeHandle != 0)
                    {
                        // Re-try with the native handle
                        uint msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                        IntPtr result;
                        SendMessageTimeout(new IntPtr(nativeHandle), msg, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out result);
                        if (result != IntPtr.Zero)
                        {
                            object obj;
                            var iidHtmlDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                            var hr = ObjectFromLresult(result, ref iidHtmlDocument, IntPtr.Zero, out obj);
                            if (hr == 0 && obj != null)
                            {
                                Debug.WriteLine($"[MSHTML] SUCCESS via UI Automation native handle");
                                return obj;
                            }
                        }
                    }
                }
                _lastDocumentError += "UIA method failed; ";
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method4: {ex.Message}; ";
                Debug.WriteLine($"[MSHTML] Method 4 exception: {ex.Message}");
            }

            // Method 5: Try IServiceProvider from IAccessible
            try
            {
                Debug.WriteLine($"[MSHTML] Method 5: IServiceProvider");
                var iidAccessible = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
                var accessible = AccessibleObjectFromWindow(hwnd, 0, ref iidAccessible);
                if (accessible != null && accessible is IServiceProvider sp)
                {
                    var iidDoc = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                    var sidDoc = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                    IntPtr pDoc;
                    int hr = sp.QueryService(ref sidDoc, ref iidDoc, out pDoc);
                    if (hr == 0 && pDoc != IntPtr.Zero)
                    {
                        var doc = Marshal.GetObjectForIUnknown(pDoc);
                        Marshal.Release(pDoc);
                        Debug.WriteLine($"[MSHTML] SUCCESS via IServiceProvider");
                        return doc;
                    }
                    _lastDocumentError += $"IServiceProvider.QueryService hr={hr}; ";
                }
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method5: {ex.Message}; ";
            }

            // Method 6: Try finding child window Shell DocObject View
            try
            {
                Debug.WriteLine($"[MSHTML] Method 6: Finding Shell DocObject View");
                IntPtr shellDocView = FindWindowEx(hwnd, IntPtr.Zero, "Shell DocObject View", null);
                if (shellDocView != IntPtr.Zero)
                {
                    IntPtr ieServer = FindWindowEx(shellDocView, IntPtr.Zero, "Internet Explorer_Server", null);
                    if (ieServer != IntPtr.Zero && ieServer != hwnd)
                    {
                        // Recursively try with the found IE Server
                        var doc = GetHTMLDocumentDirect(ieServer);
                        if (doc != null)
                        {
                            Debug.WriteLine($"[MSHTML] SUCCESS via child window search");
                            return doc;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _lastDocumentError += $"Method6: {ex.Message}; ";
            }

            Debug.WriteLine($"[MSHTML] FAILED: All methods exhausted for hwnd=0x{hwnd.ToInt64():X8}");
            Debug.WriteLine($"[MSHTML] Errors: {_lastDocumentError}");
            return null;
        }

        private string _lastDocumentError = "";

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        private object GetHTMLDocumentDirect(IntPtr hwnd)
        {
            try
            {
                uint msg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                IntPtr result;
                SendMessageTimeout(hwnd, msg, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out result);
                if (result != IntPtr.Zero)
                {
                    object obj;
                    var iidHtmlDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
                    var hr = ObjectFromLresult(result, ref iidHtmlDocument, IntPtr.Zero, out obj);
                    if (hr == 0 && obj != null)
                        return obj;
                }
            }
            catch { }
            return null;
        }

        [DllImport("oleacc.dll")]
        private static extern int ObjectFromLresult(IntPtr lResult, ref Guid riid, IntPtr wParam, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        private void ExtractElementFromDocument(object document, System.Windows.Point point, ElementInfo info, CollectionProfile profile)
        {
            try
            {
                // Use dynamic to work with COM objects
                dynamic htmlDoc = document;
                bool hasAccessDenied = false;

                // Get the window handle for coordinate conversion first
                POINT screenPoint = new POINT { X = (int)point.X, Y = (int)point.Y };
                var hwnd = WindowFromPoint(screenPoint);

                // Get window rect to calculate proper offset
                RECT windowRect;
                GetWindowRect(hwnd, out windowRect);

                // Calculate client coordinates relative to the IE window
                int clientX = (int)point.X - windowRect.Left;
                int clientY = (int)point.Y - windowRect.Top;

                info.CollectionErrors.Add($"[MSHTML] Screen: ({point.X}, {point.Y}) -> Window: ({windowRect.Left}, {windowRect.Top})");
                info.CollectionErrors.Add($"[MSHTML] Client coords: ({clientX}, {clientY})");

                // Try to access frames first - often the content is in a frame
                dynamic targetDoc = htmlDoc;
                try
                {
                    // Check if there are frames
                    dynamic frames = htmlDoc.frames;
                    if (frames != null)
                    {
                        int frameCount = 0;
                        try { frameCount = (int)frames.length; } catch { }
                        info.CollectionErrors.Add($"[MSHTML] Found {frameCount} frames");

                        // Try each frame to find one that might contain our point
                        for (int f = 0; f < frameCount; f++)
                        {
                            try
                            {
                                dynamic frame = frames[f];
                                dynamic frameDoc = frame.document;
                                if (frameDoc != null)
                                {
                                    // Try to access this frame's document
                                    string testTitle = frameDoc.title?.ToString();
                                    info.CollectionErrors.Add($"[MSHTML] Frame {f} accessible, title: '{testTitle}'");
                                    targetDoc = frameDoc;
                                    break;
                                }
                            }
                            catch (Exception frameEx)
                            {
                                info.CollectionErrors.Add($"[MSHTML] Frame {f} access error: {frameEx.Message}");
                            }
                        }
                    }
                }
                catch { }

                // Extract document info - wrap each in try/catch to handle E_ACCESSDENIED
                try
                {
                    info.DocumentTitle = targetDoc.title?.ToString();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("ACCESSDENIED") || ex.Message.Contains("Erişim engellendi"))
                        hasAccessDenied = true;
                }

                try { info.DocumentUrl = targetDoc.URL?.ToString(); } catch { }
                try { info.DocumentDomain = targetDoc.domain?.ToString(); } catch { }
                try { info.DocumentReadyState = targetDoc.readyState?.ToString(); } catch { }
                try { info.DocumentCharset = targetDoc.charset?.ToString(); } catch { }
                try { info.DocumentCompatMode = targetDoc.compatMode?.ToString(); } catch { }

                // Get element counts with individual error handling
                try { info.MSHTML_Doc_All_Count = (int?)targetDoc.all?.length; } catch { }
                try { info.DocumentFormsCount = (int?)targetDoc.forms?.length; } catch { }
                try { info.DocumentImagesCount = (int?)targetDoc.images?.length; } catch { }
                try { info.DocumentLinksCount = (int?)targetDoc.links?.length; } catch { }
                try { info.DocumentScriptsCount = (int?)targetDoc.scripts?.length; } catch { }

                info.CollectionErrors.Add($"[MSHTML] Document info: Title='{info.DocumentTitle}', URL='{info.DocumentUrl}'");
                info.CollectionErrors.Add($"[MSHTML] Elements: All={info.MSHTML_Doc_All_Count}, Forms={info.DocumentFormsCount}");

                // Account for scroll position in the document
                int scrollLeft = 0, scrollTop = 0;
                try
                {
                    dynamic body = targetDoc.body;
                    if (body != null)
                    {
                        try { scrollLeft = (int)(body.scrollLeft ?? 0); } catch { }
                        try { scrollTop = (int)(body.scrollTop ?? 0); } catch { }
                    }
                }
                catch { }

                // Also try document element scroll
                try
                {
                    dynamic docEl = targetDoc.documentElement;
                    if (docEl != null)
                    {
                        int docScrollLeft = 0, docScrollTop = 0;
                        try { docScrollLeft = (int)(docEl.scrollLeft ?? 0); } catch { }
                        try { docScrollTop = (int)(docEl.scrollTop ?? 0); } catch { }
                        if (docScrollLeft > scrollLeft) scrollLeft = docScrollLeft;
                        if (docScrollTop > scrollTop) scrollTop = docScrollTop;
                    }
                }
                catch { }

                info.CollectionErrors.Add($"[MSHTML] Scroll: ({scrollLeft}, {scrollTop})");

                // Try multiple approaches to find element
                dynamic element = null;
                string foundMethod = "";

                // Method 1: Direct elementFromPoint with client coords
                try
                {
                    element = targetDoc.elementFromPoint(clientX, clientY);
                    if (element != null)
                    {
                        string tag = null;
                        try { tag = element.tagName?.ToString(); } catch { }
                        if (tag != null && tag.ToUpper() != "BODY" && tag.ToUpper() != "HTML")
                        {
                            foundMethod = "elementFromPoint(client)";
                        }
                        else
                        {
                            element = null; // Reset if we got BODY/HTML
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("ACCESSDENIED") || ex.Message.Contains("Erişim engellendi"))
                        hasAccessDenied = true;
                    info.CollectionErrors.Add($"[MSHTML] elementFromPoint error: {ex.Message}");
                }

                // Method 2: Try with scroll offset adjustment
                if (element == null)
                {
                    try
                    {
                        element = targetDoc.elementFromPoint(clientX + scrollLeft, clientY + scrollTop);
                        if (element != null)
                        {
                            string tag = null;
                            try { tag = element.tagName?.ToString(); } catch { }
                            if (tag != null && tag.ToUpper() != "BODY" && tag.ToUpper() != "HTML")
                            {
                                foundMethod = "elementFromPoint(scroll-adjusted)";
                            }
                            else
                            {
                                element = null;
                            }
                        }
                    }
                    catch { }
                }

                // Method 3: Use ScreenToClient API
                if (element == null)
                {
                    try
                    {
                        POINT clientPt = new POINT { X = (int)point.X, Y = (int)point.Y };
                        ScreenToClient(hwnd, ref clientPt);
                        element = targetDoc.elementFromPoint(clientPt.X, clientPt.Y);
                        if (element != null)
                        {
                            string tag = null;
                            try { tag = element.tagName?.ToString(); } catch { }
                            if (tag != null && tag.ToUpper() != "BODY" && tag.ToUpper() != "HTML")
                            {
                                foundMethod = "elementFromPoint(ScreenToClient)";
                            }
                            else
                            {
                                element = null;
                            }
                        }
                    }
                    catch { }
                }

                // Method 4: If still BODY, try to find element under cursor by iterating
                if (element == null)
                {
                    try
                    {
                        // Get all elements and find one that contains the point
                        dynamic allElements = targetDoc.all;
                        if (allElements != null)
                        {
                            int count = 0;
                            try { count = (int)allElements.length; } catch { }
                            info.CollectionErrors.Add($"[MSHTML] Searching through {count} elements...");

                            dynamic bestMatch = null;
                            double bestArea = double.MaxValue;

                            for (int i = count - 1; i >= 0 && i > count - 500; i--) // Check last 500 elements (most likely on top)
                            {
                                try
                                {
                                    dynamic el = allElements[i];
                                    if (el == null) continue;

                                    string tag = null;
                                    try { tag = el.tagName?.ToString()?.ToUpper(); } catch { continue; }
                                    if (tag == "BODY" || tag == "HTML" || tag == "HEAD" || tag == "SCRIPT" || tag == "STYLE")
                                        continue;

                                    // Get bounding rect
                                    dynamic rect = null;
                                    try { rect = el.getBoundingClientRect(); } catch { continue; }
                                    if (rect == null) continue;

                                    double left = 0, top = 0, right = 0, bottom = 0;
                                    try
                                    {
                                        left = (double)rect.left;
                                        top = (double)rect.top;
                                        right = (double)rect.right;
                                        bottom = (double)rect.bottom;
                                    }
                                    catch { continue; }

                                    double width = right - left;
                                    double height = bottom - top;

                                    // Check if point is inside element
                                    if (clientX >= left && clientX <= right && clientY >= top && clientY <= bottom)
                                    {
                                        double area = width * height;
                                        // Prefer smaller elements (more specific)
                                        if (area < bestArea && area > 0)
                                        {
                                            bestArea = area;
                                            bestMatch = el;
                                        }
                                    }
                                }
                                catch { }
                            }

                            if (bestMatch != null)
                            {
                                element = bestMatch;
                                foundMethod = "bounding rect search";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("ACCESSDENIED") || ex.Message.Contains("Erişim engellendi"))
                            hasAccessDenied = true;
                        info.CollectionErrors.Add($"[MSHTML] Element search error: {ex.Message}");
                    }
                }

                // If we have access denied, try IAccessible approach
                if (hasAccessDenied && element == null)
                {
                    info.CollectionErrors.Add($"[MSHTML] Access denied detected, trying IAccessible fallback...");
                    try
                    {
                        element = GetElementViaAccessibility(hwnd, point, info);
                        if (element != null)
                            foundMethod = "IAccessible fallback";
                    }
                    catch (Exception ex)
                    {
                        info.CollectionErrors.Add($"[MSHTML] IAccessible fallback error: {ex.Message}");
                    }
                }

                // Final fallback: just get the body element
                if (element == null)
                {
                    try
                    {
                        element = targetDoc.body;
                        foundMethod = "fallback to body";
                    }
                    catch { }
                }

                if (element != null)
                {
                    string tagName = "unknown";
                    string elementId = "";
                    string elementName = "";
                    try { tagName = element.tagName?.ToString() ?? "null"; } catch { }
                    try { elementId = element.id?.ToString() ?? ""; } catch { }
                    try { elementName = element.name?.ToString() ?? ""; } catch { }

                    info.CollectionErrors.Add($"[MSHTML] Found element via {foundMethod}: <{tagName}> id='{elementId}' name='{elementName}'");

                    // Only call ExtractElementInfo for real HTML elements, not IAccessible fallback markers
                    if (foundMethod != "IAccessible fallback")
                    {
                        ExtractElementInfo(element, info, profile);
                    }
                    else
                    {
                        // For IAccessible fallback, set basic info from the marker object
                        info.TagName = tagName;
                        info.Tag = tagName;
                        info.HtmlId = elementId;
                        // Name was already set in GetElementViaAccessibility
                    }

                    // Log what was extracted
                    info.CollectionErrors.Add($"[MSHTML] Extracted: Name='{info.Name}', HtmlId='{info.HtmlId}', Type='{info.InputType}'");
                    info.CollectionErrors.Add($"[MSHTML] Form info: FormId='{info.MSHTML_Form_Id}', FormAction='{info.MSHTML_Form_Action}'");
                }
                else
                {
                    info.CollectionErrors.Add($"[MSHTML] ERROR: Could not find any element at point");

                    // If complete access denied, provide helpful message
                    if (hasAccessDenied)
                    {
                        info.CollectionErrors.Add($"[MSHTML] NOTE: IE Protected Mode is blocking access. Try running as Administrator or disabling Protected Mode in IE settings.");
                    }
                }

                // Extract page source code if Full profile (try last, as it's large)
                if (profile >= CollectionProfile.Full && !hasAccessDenied)
                {
                    try
                    {
                        dynamic docElement = targetDoc.documentElement;
                        if (docElement != null)
                        {
                            info.SourceCode = docElement.outerHTML?.ToString();
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"[MSHTML] Critical error: {ex.Message}");
            }
        }

        private dynamic GetElementViaAccessibility(IntPtr hwnd, System.Windows.Point point, ElementInfo info)
        {
            // Try to use IAccessible to get element info even when direct DOM access is denied
            try
            {
                var accElement = AutomationElement.FromHandle(hwnd);
                if (accElement != null)
                {
                    // Try to find element at point using UI Automation
                    var foundElement = AutomationElement.FromPoint(point);
                    if (foundElement != null)
                    {
                        // Extract what we can from the automation element
                        try { info.Name = foundElement.Current.Name; } catch { }
                        try { info.AutomationId = foundElement.Current.AutomationId; } catch { }
                        try { info.HtmlId = foundElement.Current.AutomationId; } catch { } // AutomationId often contains HTML id
                        try { info.ControlType = foundElement.Current.ControlType.ProgrammaticName; } catch { }
                        try { info.HelpText = foundElement.Current.HelpText; } catch { }
                        try { info.AcceleratorKey = foundElement.Current.AcceleratorKey; } catch { }
                        try { info.AccessKey = foundElement.Current.AccessKey; } catch { }
                        try
                        {
                            var rect = foundElement.Current.BoundingRectangle;
                            info.X = rect.X;
                            info.Y = rect.Y;
                            info.Width = rect.Width;
                            info.Height = rect.Height;
                            info.BoundingRectangle = rect;
                        }
                        catch { }

                        // Try to get value pattern
                        try
                        {
                            object valuePattern;
                            if (foundElement.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
                            {
                                var vp = (ValuePattern)valuePattern;
                                info.Value = vp.Current.Value;
                                info.ValuePattern_IsReadOnly = vp.Current.IsReadOnly;
                            }
                        }
                        catch { }

                        info.CollectionErrors.Add($"[MSHTML] IAccessible fallback collected: Name='{info.Name}', AutomationId='{info.AutomationId}'");

                        // Mark technology used
                        if (!info.TechnologiesUsed.Contains("IAccessible"))
                            info.TechnologiesUsed.Add("IAccessible");

                        // Return a marker object to indicate we found something via accessibility
                        return new { tagName = "ACCESSIBLE_ELEMENT", id = info.AutomationId, name = info.Name };
                    }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"[MSHTML] Accessibility extraction error: {ex.Message}");
            }
            return null;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private void ExtractElementInfo(dynamic element, ElementInfo info, CollectionProfile profile)
        {
            try
            {
                // === 4.1 IHTMLElement Properties ===
                info.TagName = element.tagName?.ToString();
                info.Tag = info.TagName;
                info.HtmlId = element.id?.ToString();
                info.HtmlClassName = element.className?.ToString();

                // For web elements, HTML id maps to AutomationId (as shown in Inspect.exe)
                info.AutomationId = info.HtmlId;

                Debug.WriteLine($"MSHTML Extract - TagName: '{info.TagName}', HtmlId: '{info.HtmlId}', AutomationId: '{info.AutomationId}'");

                // Name - try multiple sources for button elements
                string htmlName = element.name?.ToString();
                info.Name = htmlName;
                info.HtmlName = htmlName;

                // Store HTML name attribute separately for reference
                if (!string.IsNullOrEmpty(htmlName))
                {
                    info.CustomProperties["HtmlName"] = htmlName;
                }

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

                // === 4.1 Extended IHTMLElement Properties ===
                try { info.MSHTML_OuterText = element.outerText?.ToString(); } catch { }
                try { info.MSHTML_Language = element.language?.ToString(); } catch { }
                try { info.MSHTML_Lang = element.lang?.ToString(); } catch { }
                try { info.MSHTML_SourceIndex = (int?)element.sourceIndex; } catch { }
                try { info.MSHTML_ScopeName = element.scopeName?.ToString(); } catch { }
                try { info.MSHTML_CanHaveChildren = element.canHaveChildren; } catch { }
                try { info.MSHTML_CanHaveHTML = element.canHaveHTML; } catch { }
                try { info.MSHTML_IsContentEditable = element.isContentEditable; } catch { }
                try { info.MSHTML_HideFocus = element.hideFocus; } catch { }
                try { info.MSHTML_ContentEditable = element.contentEditable?.ToString(); } catch { }
                try { info.MSHTML_TabIndex = (int?)element.tabIndex; } catch { }
                try { info.MSHTML_AccessKey = element.accessKey?.ToString(); } catch { }

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

                        // Copy full OuterHTML to SourceCode (no truncation for source code view)
                        info.SourceCode = info.OuterHTML;

                        // Truncate display version only
                        if (info.OuterHTML?.Length > 1000)
                            info.OuterHTML = info.OuterHTML.Substring(0, 1000);
                    }
                    catch { }
                }

                // Parent element
                try
                {
                    dynamic parentEl = element.parentElement;
                    if (parentEl != null)
                    {
                        info.MSHTML_ParentElementTag = parentEl.tagName?.ToString();
                        info.ParentName = info.MSHTML_ParentElementTag;
                        try { info.ParentId = parentEl.id?.ToString(); } catch { }
                        try { info.ParentClassName = parentEl.className?.ToString(); } catch { }
                    }
                }
                catch { }

                // Children count
                try
                {
                    dynamic children = element.children;
                    if (children != null)
                        info.MSHTML_ChildrenCount = (int)children.length;
                }
                catch { }

                // Attributes
                try
                {
                    info.Href = element.href?.ToString();
                    info.Src = element.src?.ToString();
                    info.Alt = element.alt?.ToString();
                    info.Title = element.title?.ToString();
                    info.Value = element.value?.ToString();
                    info.InputValue = info.Value;
                }
                catch { }

                // Form elements
                try
                {
                    info.InputType = element.type?.ToString();
                }
                catch { }

                // === 4.2 IHTMLElement2 Properties ===
                try
                {
                    info.ClientWidth = (double)(element.clientWidth ?? 0);
                    info.ClientHeight = (double)(element.clientHeight ?? 0);
                    info.MSHTML_ClientLeft = (double)(element.clientLeft ?? 0);
                    info.MSHTML_ClientTop = (double)(element.clientTop ?? 0);
                    info.ScrollWidth = (double)(element.scrollWidth ?? 0);
                    info.ScrollHeight = (double)(element.scrollHeight ?? 0);
                    info.ScrollLeft = (double)(element.scrollLeft ?? 0);
                    info.ScrollTop = (double)(element.scrollTop ?? 0);
                    info.OffsetWidth = (double)(element.offsetWidth ?? 0);
                    info.OffsetHeight = (double)(element.offsetHeight ?? 0);
                }
                catch { }

                // Current style (computed style)
                try
                {
                    dynamic currentStyle = element.currentStyle;
                    if (currentStyle != null)
                    {
                        info.Style_Display = currentStyle.display?.ToString();
                        info.Style_Position = currentStyle.position?.ToString();
                        info.Style_Visibility = currentStyle.visibility?.ToString();
                        info.Style_Color = currentStyle.color?.ToString();
                        info.Style_BackgroundColor = currentStyle.backgroundColor?.ToString();
                        info.Style_FontSize = currentStyle.fontSize?.ToString();
                        info.Style_FontWeight = currentStyle.fontWeight?.ToString();
                        info.Style_ZIndex = currentStyle.zIndex?.ToString();
                        info.Style_Overflow = currentStyle.overflow?.ToString();

                        // Store as custom property for debugging
                        info.MSHTML_CurrentStyle = $"display:{info.Style_Display}, position:{info.Style_Position}, visibility:{info.Style_Visibility}";
                    }
                }
                catch { }

                // === 4.3 IHTMLInputElement Properties ===
                if (info.TagName?.ToUpper() == "INPUT")
                {
                    try
                    {
                        info.MSHTML_DefaultValue = element.defaultValue?.ToString();
                        info.Placeholder = element.placeholder?.ToString();
                        try { info.MSHTML_DefaultChecked = element.defaultChecked; } catch { }
                        try { info.MSHTML_MaxLength = (int?)element.maxLength; } catch { }
                        try { info.MSHTML_Size = (int?)element.size; } catch { }
                        try { info.MSHTML_ReadOnly = element.readOnly; } catch { }
                        try { info.MSHTML_Accept = element.accept?.ToString(); } catch { }
                        try { info.MSHTML_Align = element.align?.ToString(); } catch { }
                        try { info.MSHTML_UseMap = element.useMap?.ToString(); } catch { }
                        try { info.MSHTML_IndeterminateState = element.indeterminate; } catch { }

                        // Get parent form info
                        try
                        {
                            dynamic form = element.form;
                            if (form != null)
                            {
                                info.MSHTML_Form_Id = form.id?.ToString();
                                info.MSHTML_Form_Name = form.name?.ToString();
                                info.MSHTML_Form_Action = form.action?.ToString();
                                info.MSHTML_Form_Method = form.method?.ToString();
                            }
                        }
                        catch { }
                    }
                    catch { }
                }

                // === 4.4 IHTMLSelectElement Properties ===
                if (info.TagName?.ToUpper() == "SELECT")
                {
                    try
                    {
                        info.MSHTML_SelectedIndex = (int)(element.selectedIndex ?? -1);
                        try { info.MSHTML_Multiple = element.multiple; } catch { }

                        dynamic options = element.options;
                        if (options != null)
                        {
                            info.MSHTML_OptionsLength = (int)options.length;
                            for (int i = 0; i < Math.Min(options.length, 50); i++) // Limit to 50
                            {
                                try
                                {
                                    string optionText = options[i].text?.ToString();
                                    string optionValue = options[i].value?.ToString();
                                    if (!string.IsNullOrEmpty(optionText))
                                        info.MSHTML_Options.Add(optionText);
                                    if (!string.IsNullOrEmpty(optionValue))
                                        info.MSHTML_OptionValues.Add(optionValue);

                                    // Get selected option
                                    if (i == info.MSHTML_SelectedIndex)
                                    {
                                        info.MSHTML_SelectedText = optionText;
                                        info.MSHTML_SelectedValue = optionValue;
                                    }
                                }
                                catch { }
                            }
                        }

                        // Get parent form info
                        try
                        {
                            dynamic form = element.form;
                            if (form != null)
                            {
                                info.MSHTML_Form_Id = form.id?.ToString();
                                info.MSHTML_Form_Name = form.name?.ToString();
                                info.MSHTML_Form_Action = form.action?.ToString();
                                info.MSHTML_Form_Method = form.method?.ToString();
                            }
                        }
                        catch { }
                    }
                    catch { }
                }

                // === 4.5 IHTMLTextAreaElement Properties ===
                if (info.TagName?.ToUpper() == "TEXTAREA")
                {
                    try
                    {
                        info.MSHTML_Cols = (int?)element.cols;
                        info.MSHTML_Rows = (int?)element.rows;
                        try { info.MSHTML_Wrap = element.wrap?.ToString(); } catch { }
                        try { info.MSHTML_DefaultValue = element.defaultValue?.ToString(); } catch { }
                        try { info.MSHTML_ReadOnly = element.readOnly; } catch { }
                        try { info.MSHTML_MaxLength = (int?)element.maxLength; } catch { }
                        info.Placeholder = element.placeholder?.ToString();

                        // Get parent form info
                        try
                        {
                            dynamic form = element.form;
                            if (form != null)
                            {
                                info.MSHTML_Form_Id = form.id?.ToString();
                                info.MSHTML_Form_Name = form.name?.ToString();
                                info.MSHTML_Form_Action = form.action?.ToString();
                                info.MSHTML_Form_Method = form.method?.ToString();
                            }
                        }
                        catch { }
                    }
                    catch { }
                }

                // === 4.6 IHTMLButtonElement Properties ===
                if (info.TagName?.ToUpper() == "BUTTON")
                {
                    try
                    {
                        info.MSHTML_ButtonType = element.type?.ToString();
                        try { info.MSHTML_FormAction = element.formAction?.ToString(); } catch { }
                        try { info.MSHTML_FormMethod = element.formMethod?.ToString(); } catch { }

                        // Get parent form info
                        try
                        {
                            dynamic form = element.form;
                            if (form != null)
                            {
                                info.MSHTML_Form_Id = form.id?.ToString();
                                info.MSHTML_Form_Name = form.name?.ToString();
                                info.MSHTML_Form_Action = form.action?.ToString();
                                info.MSHTML_Form_Method = form.method?.ToString();
                            }
                        }
                        catch { }
                    }
                    catch { }
                }

                // === 4.7 IHTMLAnchorElement Properties ===
                if (info.TagName?.ToUpper() == "A")
                {
                    try
                    {
                        info.MSHTML_Target = element.target?.ToString();
                        info.MSHTML_Protocol = element.protocol?.ToString();
                        info.MSHTML_Host = element.host?.ToString();
                        info.MSHTML_Hostname = element.hostname?.ToString();
                        info.MSHTML_Port = element.port?.ToString();
                        info.MSHTML_Pathname = element.pathname?.ToString();
                        info.MSHTML_Search = element.search?.ToString();
                        info.MSHTML_Hash = element.hash?.ToString();
                        try { info.MSHTML_Rel = element.rel?.ToString(); } catch { }
                    }
                    catch { }
                }

                // === 4.8 IHTMLImageElement Properties ===
                if (info.TagName?.ToUpper() == "IMG")
                {
                    try
                    {
                        try { info.MSHTML_IsMap = element.isMap; } catch { }
                        try { info.MSHTML_NaturalWidth = (int?)element.naturalWidth; } catch { }
                        try { info.MSHTML_NaturalHeight = (int?)element.naturalHeight; } catch { }
                        try { info.MSHTML_Complete = element.complete; } catch { }
                        try { info.MSHTML_LongDesc = element.longDesc?.ToString(); } catch { }
                        try { info.MSHTML_Vspace = (int?)element.vspace; } catch { }
                        try { info.MSHTML_Hspace = (int?)element.hspace; } catch { }
                        try { info.MSHTML_LowSrc = element.lowsrc?.ToString(); } catch { }
                    }
                    catch { }
                }

                // === 4.9 IHTMLTableElement Properties ===
                if (info.TagName?.ToUpper() == "TABLE")
                {
                    try
                    {
                        try { info.MSHTML_Table_Caption = element.caption?.innerText?.ToString(); } catch { }
                        try { info.MSHTML_Table_Summary = element.summary?.ToString(); } catch { }
                        try { info.MSHTML_Table_Border = element.border?.ToString(); } catch { }
                        try { info.MSHTML_Table_CellPadding = element.cellPadding?.ToString(); } catch { }
                        try { info.MSHTML_Table_CellSpacing = element.cellSpacing?.ToString(); } catch { }
                        try { info.MSHTML_Table_Width = element.width?.ToString(); } catch { }
                        try { info.MSHTML_Table_BgColor = element.bgColor?.ToString(); } catch { }
                        try { info.MSHTML_Table_RowsCount = (int?)element.rows?.length; } catch { }
                        try { info.MSHTML_Table_TBodiesCount = (int?)element.tBodies?.length; } catch { }
                    }
                    catch { }
                }

                // === 4.10 IHTMLTableRowElement Properties ===
                if (info.TagName?.ToUpper() == "TR")
                {
                    try
                    {
                        try { info.MSHTML_Row_RowIndex = (int?)element.rowIndex; } catch { }
                        try { info.MSHTML_Row_SectionRowIndex = (int?)element.sectionRowIndex; } catch { }
                        try { info.MSHTML_Row_CellsCount = (int?)element.cells?.length; } catch { }
                        try { info.MSHTML_Row_BgColor = element.bgColor?.ToString(); } catch { }
                        try { info.MSHTML_Row_VAlign = element.vAlign?.ToString(); } catch { }
                        try { info.MSHTML_Row_Align = element.align?.ToString(); } catch { }
                    }
                    catch { }
                }

                // === 4.11 IHTMLTableCellElement Properties ===
                if (info.TagName?.ToUpper() == "TD" || info.TagName?.ToUpper() == "TH")
                {
                    try
                    {
                        try { info.MSHTML_Cell_CellIndex = (int?)element.cellIndex; } catch { }
                        try { info.MSHTML_Cell_Abbr = element.abbr?.ToString(); } catch { }
                        try { info.MSHTML_Cell_Axis = element.axis?.ToString(); } catch { }
                        try { info.MSHTML_Cell_Headers = element.headers?.ToString(); } catch { }
                        try { info.MSHTML_Cell_Scope = element.scope?.ToString(); } catch { }
                        try { info.MSHTML_Cell_NoWrap = element.noWrap?.ToString(); } catch { }
                        try { info.MSHTML_Cell_BgColor = element.bgColor?.ToString(); } catch { }
                        try { info.MSHTML_Cell_VAlign = element.vAlign?.ToString(); } catch { }
                        try { info.MSHTML_Cell_Align = element.align?.ToString(); } catch { }
                    }
                    catch { }
                }

                // === 4.12 IHTMLFrameElement / IHTMLIFrameElement Properties ===
                if (info.TagName?.ToUpper() == "FRAME" || info.TagName?.ToUpper() == "IFRAME")
                {
                    try
                    {
                        info.MSHTML_Frame_Src = element.src?.ToString();
                        info.MSHTML_Frame_Name = element.name?.ToString();
                        try { info.MSHTML_Frame_Scrolling = element.scrolling?.ToString(); } catch { }
                        try { info.MSHTML_Frame_FrameBorder = element.frameBorder?.ToString(); } catch { }
                        try { info.MSHTML_Frame_MarginWidth = element.marginWidth?.ToString(); } catch { }
                        try { info.MSHTML_Frame_MarginHeight = element.marginHeight?.ToString(); } catch { }
                        try { info.MSHTML_Frame_NoResize = element.noResize; } catch { }
                    }
                    catch { }
                }

                // Form element states - Fixed dynamic property access
                try
                {
                    // Use getAttribute for safer property access
                    var checkedAttr = element.getAttribute("checked");
                    if (checkedAttr != null)
                    {
                        info.IsChecked = true;
                    }
                    else
                    {
                        // Try direct property access with proper conversion
                        try
                        {
                            object checkedValue = element.GetType().InvokeMember("checked",
                                System.Reflection.BindingFlags.GetProperty,
                                null, element, null);
                            if (checkedValue != null && checkedValue is bool)
                            {
                                info.IsChecked = (bool)checkedValue;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                try
                {
                    // Check disabled attribute
                    var disabledAttr = element.getAttribute("disabled");
                    if (disabledAttr != null)
                    {
                        info.IsDisabled = true;
                    }
                    else
                    {
                        // Try direct property access
                        try
                        {
                            object disabledValue = element.GetType().InvokeMember("disabled",
                                System.Reflection.BindingFlags.GetProperty,
                                null, element, null);
                            if (disabledValue != null && disabledValue is bool)
                            {
                                info.IsDisabled = (bool)disabledValue;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // Position and size from getBoundingClientRect
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
                    info.IsHidden = !info.IsVisible;
                }
                catch { }

                // Build XPath
                if (profile >= CollectionProfile.Standard)
                {
                    info.XPath = BuildXPath(element);
                    info.CssSelector = BuildCssSelector(element);
                }

                // Collect attributes if Full profile
                if (profile == CollectionProfile.Full)
                {
                    CollectAllAttributes(element, info);
                }

                // === 4.13-4.14 IHTMLDocument Properties ===
                try
                {
                    dynamic doc = element.document;
                    if (doc != null)
                    {
                        info.DocumentTitle = doc.title?.ToString();
                        info.DocumentUrl = doc.URL?.ToString();
                        info.DocumentDomain = doc.domain?.ToString();
                        info.DocumentReadyState = doc.readyState?.ToString();

                        // Additional document properties
                        try { info.DocumentCookie = doc.cookie?.ToString(); } catch { }
                        try { info.DocumentCharset = doc.charset?.ToString(); } catch { }
                        try { info.DocumentLastModified = doc.lastModified?.ToString(); } catch { }
                        try { info.DocumentReferrer = doc.referrer?.ToString(); } catch { }
                        try { info.DocumentCompatMode = doc.compatMode?.ToString(); } catch { }
                        try { info.DocumentDesignMode = doc.designMode?.ToString(); } catch { }
                        try { info.DocumentDir = doc.dir?.ToString(); } catch { }

                        // IHTMLDocument2/3/4/5 properties
                        try { info.MSHTML_Doc_Protocol = doc.protocol?.ToString(); } catch { }
                        try { info.MSHTML_Doc_NameProp = doc.nameProp?.ToString(); } catch { }
                        try { info.MSHTML_Doc_FileCreatedDate = doc.fileCreatedDate?.ToString(); } catch { }
                        try { info.MSHTML_Doc_FileModifiedDate = doc.fileModifiedDate?.ToString(); } catch { }
                        try { info.MSHTML_Doc_FileSize = doc.fileSize?.ToString(); } catch { }
                        try { info.MSHTML_Doc_MimeType = doc.mimeType?.ToString(); } catch { }
                        try { info.MSHTML_Doc_Security = doc.security?.ToString(); } catch { }

                        // Document type
                        try
                        {
                            dynamic docType = doc.doctype;
                            if (docType != null)
                                info.DocumentDocType = docType.name?.ToString();
                        }
                        catch { }

                        // Document element counts
                        try
                        {
                            dynamic frames = doc.frames;
                            if (frames != null)
                                info.DocumentFramesCount = (int)frames.length;
                        }
                        catch { }

                        try
                        {
                            dynamic scripts = doc.scripts;
                            if (scripts != null)
                                info.DocumentScriptsCount = (int)scripts.length;
                        }
                        catch { }

                        try
                        {
                            dynamic links = doc.links;
                            if (links != null)
                                info.DocumentLinksCount = (int)links.length;
                        }
                        catch { }

                        try
                        {
                            dynamic images = doc.images;
                            if (images != null)
                                info.DocumentImagesCount = (int)images.length;
                        }
                        catch { }

                        try
                        {
                            dynamic forms = doc.forms;
                            if (forms != null)
                                info.DocumentFormsCount = (int)forms.length;
                        }
                        catch { }

                        try
                        {
                            dynamic anchors = doc.anchors;
                            if (anchors != null)
                                info.MSHTML_Doc_Anchors_Count = (int)anchors.length;
                        }
                        catch { }

                        try
                        {
                            dynamic applets = doc.applets;
                            if (applets != null)
                                info.MSHTML_Doc_Applets_Count = (int)applets.length;
                        }
                        catch { }

                        try
                        {
                            dynamic embeds = doc.embeds;
                            if (embeds != null)
                                info.MSHTML_Doc_Embeds_Count = (int)embeds.length;
                        }
                        catch { }

                        try
                        {
                            dynamic all = doc.all;
                            if (all != null)
                                info.MSHTML_Doc_All_Count = (int)all.length;
                        }
                        catch { }

                        try
                        {
                            dynamic activeElement = doc.activeElement;
                            if (activeElement != null)
                            {
                                string activeTag = activeElement.tagName?.ToString();
                                string activeId = activeElement.id?.ToString();
                                info.DocumentActiveElement = !string.IsNullOrEmpty(activeId) ? $"{activeTag}#{activeId}" : activeTag;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // Table detection
                ExtractTableInfo(element, info);

                // Mark technology used
                if (!info.TechnologiesUsed.Contains("MSHTML"))
                    info.TechnologiesUsed.Add("MSHTML");
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Element extraction error: {ex.Message}");
            }
        }

        private string BuildCssSelector(dynamic element)
        {
            try
            {
                if (element.id != null && !string.IsNullOrEmpty(element.id))
                {
                    return $"#{element.id}";
                }

                var path = new List<string>();
                dynamic current = element;

                while (current != null && current.tagName != null && current.tagName != "HTML")
                {
                    string tagName = current.tagName.ToString().ToLower();
                    string selector = tagName;

                    // Add id if available
                    if (current.id != null && !string.IsNullOrEmpty(current.id))
                    {
                        selector += $"#{current.id}";
                        path.Insert(0, selector);
                        break; // ID is unique, no need to go further
                    }

                    // Add class if available
                    if (current.className != null && !string.IsNullOrEmpty(current.className))
                    {
                        string className = current.className.ToString().Split(' ')[0];
                        if (!string.IsNullOrEmpty(className))
                            selector += $".{className}";
                    }

                    path.Insert(0, selector);
                    current = current.parentElement;
                }

                return string.Join(" > ", path);
            }
            catch
            {
                return "";
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