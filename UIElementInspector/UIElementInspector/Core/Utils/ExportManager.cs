using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Newtonsoft.Json;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Manages exporting element data to various formats
    /// </summary>
    public class ExportManager
    {
        /// <summary>
        /// Exports elements to CSV format and returns the content
        /// </summary>
        public string ExportToCSV(IEnumerable<ElementInfo> elements)
        {
            try
            {
                var csv = new StringBuilder();

                // Write header
                csv.AppendLine("ElementType,Name,AutomationId,ClassName,HtmlId,TagName,Value,InnerText,XPath,CssSelector," +
                               "X,Y,Width,Height,IsEnabled,IsVisible,IsOffscreen,ProcessId,FrameworkId,ControlType," +
                               "ParentName,ParentId,WindowTitle,ApplicationName,DocumentUrl,Role,AriaLabel," +
                               "DetectionMethod,CollectionProfile,CaptureTime,Href,Src,Alt,Title,InputType");

                // Write data rows
                foreach (var element in elements)
                {
                    csv.AppendLine($"{EscapeCSV(element.ElementType)}," +
                                   $"{EscapeCSV(element.Name)}," +
                                   $"{EscapeCSV(element.AutomationId)}," +
                                   $"{EscapeCSV(element.ClassName)}," +
                                   $"{EscapeCSV(element.HtmlId)}," +
                                   $"{EscapeCSV(element.TagName)}," +
                                   $"{EscapeCSV(element.Value)}," +
                                   $"{EscapeCSV(element.InnerText)}," +
                                   $"{EscapeCSV(element.XPath)}," +
                                   $"{EscapeCSV(element.CssSelector)}," +
                                   $"{element.X},{element.Y},{element.Width},{element.Height}," +
                                   $"{element.IsEnabled},{element.IsVisible},{element.IsOffscreen}," +
                                   $"{element.ProcessId}," +
                                   $"{EscapeCSV(element.FrameworkId)}," +
                                   $"{EscapeCSV(element.ControlType)}," +
                                   $"{EscapeCSV(element.ParentName)}," +
                                   $"{EscapeCSV(element.ParentId)}," +
                                   $"{EscapeCSV(element.WindowTitle)}," +
                                   $"{EscapeCSV(element.ApplicationName)}," +
                                   $"{EscapeCSV(element.DocumentUrl)}," +
                                   $"{EscapeCSV(element.Role)}," +
                                   $"{EscapeCSV(element.AriaLabel)}," +
                                   $"{EscapeCSV(element.DetectionMethod)}," +
                                   $"{EscapeCSV(element.CollectionProfile)}," +
                                   $"{element.CaptureTime:yyyy-MM-dd HH:mm:ss}," +
                                   $"{EscapeCSV(element.Href)}," +
                                   $"{EscapeCSV(element.Src)}," +
                                   $"{EscapeCSV(element.Alt)}," +
                                   $"{EscapeCSV(element.Title)}," +
                                   $"{EscapeCSV(element.InputType)}");
                }

                return csv.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to CSV: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports elements to TXT format (similar to inspect.exe)
        /// </summary>
        public string ExportToText(IEnumerable<ElementInfo> elements)
        {
            try
            {
                var txt = new StringBuilder();
                int elementIndex = 1;

                foreach (var element in elements)
                {
                    txt.AppendLine($"========== Element #{elementIndex} ==========");
                    txt.AppendLine($"Capture Time: {element.CaptureTime:yyyy-MM-dd HH:mm:ss}");
                    txt.AppendLine($"Detection Method: {element.DetectionMethod}");
                    txt.AppendLine($"Collection Profile: {element.CollectionProfile}");
                    txt.AppendLine();

                    txt.AppendLine("--- Basic Properties ---");
                    AppendProperty(txt, "Element Type", element.ElementType);
                    AppendProperty(txt, "Name", element.Name);
                    AppendProperty(txt, "Class Name", element.ClassName);
                    AppendProperty(txt, "Value", element.Value);
                    AppendProperty(txt, "Description", element.Description);
                    txt.AppendLine();

                    txt.AppendLine("--- UI Automation Properties ---");
                    AppendProperty(txt, "AutomationId", element.AutomationId);
                    AppendProperty(txt, "Control Type", element.ControlType);
                    AppendProperty(txt, "Localized Control Type", element.LocalizedControlType);
                    AppendProperty(txt, "Framework Id", element.FrameworkId);
                    AppendProperty(txt, "Process Id", element.ProcessId.ToString());
                    AppendProperty(txt, "Runtime Id", element.RuntimeId);
                    AppendProperty(txt, "Native Window Handle", element.NativeWindowHandle.ToString());
                    AppendProperty(txt, "Is Enabled", element.IsEnabled.ToString());
                    AppendProperty(txt, "Is Offscreen", element.IsOffscreen.ToString());
                    AppendProperty(txt, "Has Keyboard Focus", element.HasKeyboardFocus.ToString());
                    AppendProperty(txt, "Is Keyboard Focusable", element.IsKeyboardFocusable.ToString());
                    AppendProperty(txt, "Is Password", element.IsPassword.ToString());
                    AppendProperty(txt, "Help Text", element.HelpText);
                    AppendProperty(txt, "Accelerator Key", element.AcceleratorKey);
                    AppendProperty(txt, "Access Key", element.AccessKey);

                    if (element.SupportedPatterns?.Count > 0)
                    {
                        txt.AppendLine($"Supported Patterns: {string.Join(", ", element.SupportedPatterns)}");
                    }
                    txt.AppendLine();

                    txt.AppendLine("--- Position & Size ---");
                    txt.AppendLine($"Bounding Rectangle: {element.BoundingRectangle}");
                    txt.AppendLine($"X: {element.X}, Y: {element.Y}");
                    txt.AppendLine($"Width: {element.Width}, Height: {element.Height}");
                    txt.AppendLine($"Clickable Point: {element.ClickablePoint}");
                    txt.AppendLine();

                    if (!string.IsNullOrEmpty(element.TagName))
                    {
                        txt.AppendLine("--- Web/HTML Properties ---");
                        AppendProperty(txt, "Tag Name", element.TagName);
                        AppendProperty(txt, "HTML Id", element.HtmlId);
                        AppendProperty(txt, "HTML Class", element.HtmlClassName);
                        AppendProperty(txt, "Inner Text", element.InnerText);
                        AppendProperty(txt, "Href", element.Href);
                        AppendProperty(txt, "Src", element.Src);
                        AppendProperty(txt, "Alt", element.Alt);
                        AppendProperty(txt, "Role", element.Role);
                        AppendProperty(txt, "Aria Label", element.AriaLabel);
                        txt.AppendLine();
                    }

                    txt.AppendLine("--- Selectors ---");
                    AppendProperty(txt, "XPath", element.XPath);
                    AppendProperty(txt, "CSS Selector", element.CssSelector);
                    AppendProperty(txt, "Windows Path", element.WindowsPath);
                    txt.AppendLine();

                    txt.AppendLine("--- Hierarchy ---");
                    AppendProperty(txt, "Parent Name", element.ParentName);
                    AppendProperty(txt, "Parent Id", element.ParentId);
                    AppendProperty(txt, "Parent Class", element.ParentClassName);
                    txt.AppendLine($"Children Count: {element.Children?.Count ?? 0}");
                    txt.AppendLine($"Tree Level: {element.TreeLevel}");
                    txt.AppendLine();

                    if (element.CustomProperties?.Count > 0)
                    {
                        txt.AppendLine("--- Custom Properties ---");
                        foreach (var prop in element.CustomProperties)
                        {
                            txt.AppendLine($"{prop.Key}: {prop.Value}");
                        }
                        txt.AppendLine();
                    }

                    if (element.CollectionErrors?.Count > 0)
                    {
                        txt.AppendLine("--- Collection Errors ---");
                        foreach (var error in element.CollectionErrors)
                        {
                            txt.AppendLine($"- {error}");
                        }
                        txt.AppendLine();
                    }

                    txt.AppendLine($"Collection Duration: {element.CollectionDuration.TotalMilliseconds}ms");
                    txt.AppendLine();
                    txt.AppendLine("=" + new string('=', 50));
                    txt.AppendLine();

                    elementIndex++;
                }

                return txt.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to TXT: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports elements to JSON format
        /// </summary>
        public string ExportToJSON(IEnumerable<ElementInfo> elements)
        {
            try
            {
                var json = JsonConvert.SerializeObject(elements, Newtonsoft.Json.Formatting.Indented, new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore,
                    DefaultValueHandling = DefaultValueHandling.Include
                });

                return json;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to JSON: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports elements to XML format
        /// </summary>
        public string ExportToXML(IEnumerable<ElementInfo> elements)
        {
            try
            {
                var xml = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("UIElements",
                        new XAttribute("ExportTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                        new XAttribute("Count", elements.Count()),
                        from element in elements
                        select new XElement("Element",
                            new XAttribute("Type", element.ElementType ?? ""),
                            new XAttribute("Name", element.Name ?? ""),
                            new XElement("Basic",
                                new XElement("ElementType", element.ElementType),
                                new XElement("Name", element.Name),
                                new XElement("ClassName", element.ClassName),
                                new XElement("Value", element.Value),
                                new XElement("Description", element.Description)
                            ),
                            new XElement("UIAutomation",
                                new XElement("AutomationId", element.AutomationId),
                                new XElement("ControlType", element.ControlType),
                                new XElement("FrameworkId", element.FrameworkId),
                                new XElement("ProcessId", element.ProcessId),
                                new XElement("RuntimeId", element.RuntimeId),
                                new XElement("IsEnabled", element.IsEnabled),
                                new XElement("IsOffscreen", element.IsOffscreen),
                                new XElement("HasKeyboardFocus", element.HasKeyboardFocus),
                                new XElement("SupportedPatterns",
                                    element.SupportedPatterns?.Select(p => new XElement("Pattern", p))
                                )
                            ),
                            new XElement("Position",
                                new XElement("X", element.X),
                                new XElement("Y", element.Y),
                                new XElement("Width", element.Width),
                                new XElement("Height", element.Height),
                                new XElement("BoundingRectangle", element.BoundingRectangle.ToString())
                            ),
                            element.TagName != null ? new XElement("Web",
                                new XElement("TagName", element.TagName),
                                new XElement("HtmlId", element.HtmlId),
                                new XElement("HtmlClassName", element.HtmlClassName),
                                new XElement("InnerText", element.InnerText),
                                new XElement("Href", element.Href),
                                new XElement("Src", element.Src),
                                new XElement("Role", element.Role),
                                new XElement("AriaLabel", element.AriaLabel)
                            ) : null,
                            new XElement("Selectors",
                                new XElement("XPath", element.XPath),
                                new XElement("CssSelector", element.CssSelector),
                                new XElement("WindowsPath", element.WindowsPath)
                            ),
                            new XElement("Hierarchy",
                                new XElement("ParentName", element.ParentName),
                                new XElement("ParentId", element.ParentId),
                                new XElement("TreeLevel", element.TreeLevel),
                                new XElement("ChildrenCount", element.Children?.Count ?? 0)
                            ),
                            new XElement("Metadata",
                                new XElement("DetectionMethod", element.DetectionMethod),
                                new XElement("CollectionProfile", element.CollectionProfile),
                                new XElement("CaptureTime", element.CaptureTime),
                                new XElement("CollectionDuration", element.CollectionDuration.TotalMilliseconds)
                            )
                        )
                    )
                );

                return xml.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to XML: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports elements to HTML format with interactive table
        /// </summary>
        public string ExportToHTML(IEnumerable<ElementInfo> elements)
        {
            try
            {
                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html lang=\"en\">");
                html.AppendLine("<head>");
                html.AppendLine("    <meta charset=\"UTF-8\">");
                html.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                html.AppendLine("    <title>UI Elements Report</title>");
                html.AppendLine("    <style>");
                html.AppendLine("        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }");
                html.AppendLine("        h1 { color: #333; }");
                html.AppendLine("        .info { background-color: #e3f2fd; padding: 10px; border-radius: 5px; margin-bottom: 20px; }");
                html.AppendLine("        table { width: 100%; border-collapse: collapse; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }");
                html.AppendLine("        th { background-color: #2196F3; color: white; padding: 12px; text-align: left; position: sticky; top: 0; }");
                html.AppendLine("        td { padding: 10px; border-bottom: 1px solid #ddd; }");
                html.AppendLine("        tr:hover { background-color: #f5f5f5; }");
                html.AppendLine("        .truncate { max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }");
                html.AppendLine("        .filter-input { margin: 10px 0; padding: 8px; width: 300px; border: 1px solid #ddd; border-radius: 4px; }");
                html.AppendLine("    </style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");
                html.AppendLine("    <h1>UI Elements Report</h1>");
                html.AppendLine($"    <div class=\"info\">");
                html.AppendLine($"        <strong>Export Time:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}<br>");
                html.AppendLine($"        <strong>Total Elements:</strong> {elements.Count()}<br>");
                html.AppendLine($"        <strong>Detection Methods:</strong> {string.Join(", ", elements.Select(e => e.DetectionMethod).Distinct())}");
                html.AppendLine($"    </div>");

                html.AppendLine("    <input type=\"text\" class=\"filter-input\" id=\"filterInput\" placeholder=\"Filter elements...\" onkeyup=\"filterTable()\">");

                html.AppendLine("    <table id=\"elementsTable\">");
                html.AppendLine("        <thead>");
                html.AppendLine("            <tr>");
                html.AppendLine("                <th>#</th>");
                html.AppendLine("                <th>Type</th>");
                html.AppendLine("                <th>Name</th>");
                html.AppendLine("                <th>ID</th>");
                html.AppendLine("                <th>Class</th>");
                html.AppendLine("                <th>Value</th>");
                html.AppendLine("                <th>Position</th>");
                html.AppendLine("                <th>Size</th>");
                html.AppendLine("                <th>Visible</th>");
                html.AppendLine("                <th>XPath</th>");
                html.AppendLine("                <th>Detection</th>");
                html.AppendLine("            </tr>");
                html.AppendLine("        </thead>");
                html.AppendLine("        <tbody>");

                int index = 1;
                foreach (var element in elements)
                {
                    html.AppendLine("            <tr>");
                    html.AppendLine($"                <td>{index}</td>");
                    html.AppendLine($"                <td>{HtmlEncode(element.ElementType)}</td>");
                    html.AppendLine($"                <td class=\"truncate\" title=\"{HtmlEncode(element.Name)}\">{HtmlEncode(element.Name)}</td>");
                    html.AppendLine($"                <td class=\"truncate\" title=\"{HtmlEncode(element.AutomationId ?? element.HtmlId)}\">{HtmlEncode(element.AutomationId ?? element.HtmlId)}</td>");
                    html.AppendLine($"                <td class=\"truncate\" title=\"{HtmlEncode(element.ClassName ?? element.HtmlClassName)}\">{HtmlEncode(element.ClassName ?? element.HtmlClassName)}</td>");
                    html.AppendLine($"                <td class=\"truncate\" title=\"{HtmlEncode(element.Value)}\">{HtmlEncode(element.Value)}</td>");
                    html.AppendLine($"                <td>({element.X:F0}, {element.Y:F0})</td>");
                    html.AppendLine($"                <td>{element.Width:F0} x {element.Height:F0}</td>");
                    html.AppendLine($"                <td>{(element.IsVisible ? "✓" : "✗")}</td>");
                    html.AppendLine($"                <td class=\"truncate\" title=\"{HtmlEncode(element.XPath)}\">{HtmlEncode(element.XPath)}</td>");
                    html.AppendLine($"                <td>{HtmlEncode(element.DetectionMethod)}</td>");
                    html.AppendLine("            </tr>");
                    index++;
                }

                html.AppendLine("        </tbody>");
                html.AppendLine("    </table>");

                html.AppendLine("    <script>");
                html.AppendLine("        function filterTable() {");
                html.AppendLine("            var input = document.getElementById('filterInput');");
                html.AppendLine("            var filter = input.value.toUpperCase();");
                html.AppendLine("            var table = document.getElementById('elementsTable');");
                html.AppendLine("            var tr = table.getElementsByTagName('tr');");
                html.AppendLine("            ");
                html.AppendLine("            for (var i = 1; i < tr.length; i++) {");
                html.AppendLine("                var display = false;");
                html.AppendLine("                var td = tr[i].getElementsByTagName('td');");
                html.AppendLine("                for (var j = 0; j < td.length; j++) {");
                html.AppendLine("                    if (td[j]) {");
                html.AppendLine("                        var txtValue = td[j].textContent || td[j].innerText;");
                html.AppendLine("                        if (txtValue.toUpperCase().indexOf(filter) > -1) {");
                html.AppendLine("                            display = true;");
                html.AppendLine("                            break;");
                html.AppendLine("                        }");
                html.AppendLine("                    }");
                html.AppendLine("                }");
                html.AppendLine("                tr[i].style.display = display ? '' : 'none';");
                html.AppendLine("            }");
                html.AppendLine("        }");
                html.AppendLine("    </script>");

                html.AppendLine("</body>");
                html.AppendLine("</html>");

                return html.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to HTML: {ex.Message}", ex);
            }
        }

        #region Helper Methods

        private string EscapeCSV(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            // Escape quotes and wrap in quotes if contains comma, quote, or newline
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }

            return value;
        }

        private string HtmlEncode(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            return System.Net.WebUtility.HtmlEncode(value);
        }

        private void AppendProperty(StringBuilder sb, string name, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sb.AppendLine($"{name}: {value}");
            }
        }

        #endregion
    }
}