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
            // TODO: Implement element refresh for MSHTML
            return Task.FromResult(element);
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
                info.Name = element.name?.ToString();
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
    }
}