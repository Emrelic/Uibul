using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Generates optimized XPath and CSS selectors for UI elements
    /// </summary>
    public class SelectorGenerator
    {
        /// <summary>
        /// Generates multiple XPath strategies for an element
        /// </summary>
        public static List<string> GenerateXPathStrategies(ElementInfo element)
        {
            var strategies = new List<string>();

            // Strategy 1: ID-based (highest priority)
            if (!string.IsNullOrEmpty(element.HtmlId))
            {
                strategies.Add($"//*[@id='{EscapeXPath(element.HtmlId)}']");
            }

            // Strategy 2: Unique attributes
            if (!string.IsNullOrEmpty(element.Name))
            {
                strategies.Add($"//*[@name='{EscapeXPath(element.Name)}']");
            }

            // Strategy 3: Class and tag combination
            if (!string.IsNullOrEmpty(element.TagName))
            {
                var xpath = $"//{element.TagName.ToLower()}";

                if (!string.IsNullOrEmpty(element.HtmlClassName))
                {
                    // Handle multiple classes
                    var classes = element.HtmlClassName.Split(' ').Where(c => !string.IsNullOrWhiteSpace(c));
                    foreach (var cls in classes.Take(2)) // Use first 2 classes for specificity
                    {
                        xpath += $"[contains(@class, '{EscapeXPath(cls)}')]";
                    }
                    strategies.Add(xpath);
                }

                // Strategy 4: Text content
                if (!string.IsNullOrEmpty(element.InnerText) && element.InnerText.Length < 50)
                {
                    var text = element.InnerText.Trim();
                    if (text.Length > 0 && text.Length < 50)
                    {
                        strategies.Add($"//{element.TagName.ToLower()}[contains(text(), '{EscapeXPath(text)}')]");
                        strategies.Add($"//{element.TagName.ToLower()}[normalize-space()='{EscapeXPath(text)}']");
                    }
                }

                // Strategy 5: Attribute-based
                if (!string.IsNullOrEmpty(element.Href))
                {
                    strategies.Add($"//a[@href='{EscapeXPath(element.Href)}']");
                    // Partial href match for dynamic URLs
                    var hrefPath = GetPathFromUrl(element.Href);
                    if (!string.IsNullOrEmpty(hrefPath))
                    {
                        strategies.Add($"//a[contains(@href, '{EscapeXPath(hrefPath)}')]");
                    }
                }

                if (!string.IsNullOrEmpty(element.Src))
                {
                    var srcFilename = GetFilenameFromPath(element.Src);
                    if (!string.IsNullOrEmpty(srcFilename))
                    {
                        strategies.Add($"//img[contains(@src, '{EscapeXPath(srcFilename)}')]");
                    }
                }

                // Strategy 6: ARIA attributes
                if (!string.IsNullOrEmpty(element.AriaLabel))
                {
                    strategies.Add($"//*[@aria-label='{EscapeXPath(element.AriaLabel)}']");
                }

                if (!string.IsNullOrEmpty(element.Role))
                {
                    strategies.Add($"//*[@role='{EscapeXPath(element.Role)}']");
                }
            }

            // Strategy 7: Data attributes
            if (element.DataAttributes != null && element.DataAttributes.Count > 0)
            {
                foreach (var attr in element.DataAttributes.Take(2))
                {
                    strategies.Add($"//*[@{attr.Key}='{EscapeXPath(attr.Value)}']");
                }
            }

            // Strategy 8: Position-based (least preferred)
            if (!string.IsNullOrEmpty(element.TagName) && !string.IsNullOrEmpty(element.ParentName))
            {
                strategies.Add($"//{element.ParentName.ToLower()}/{element.TagName.ToLower()}[{element.ChildIndex + 1}]");
            }

            // Strategy 9: UI Automation specific
            if (!string.IsNullOrEmpty(element.AutomationId))
            {
                strategies.Add($"//*[@AutomationId='{EscapeXPath(element.AutomationId)}']");
            }

            // Remove duplicates and return
            return strategies.Distinct().ToList();
        }

        /// <summary>
        /// Generates multiple CSS selector strategies for an element
        /// </summary>
        public static List<string> GenerateCssSelectorStrategies(ElementInfo element)
        {
            var strategies = new List<string>();

            // Strategy 1: ID selector (highest priority)
            if (!string.IsNullOrEmpty(element.HtmlId) && IsValidCssIdentifier(element.HtmlId))
            {
                strategies.Add($"#{EscapeCssId(element.HtmlId)}");
            }

            // Strategy 2: Class selectors
            if (!string.IsNullOrEmpty(element.HtmlClassName))
            {
                var classes = element.HtmlClassName.Split(' ')
                    .Where(c => !string.IsNullOrWhiteSpace(c) && IsValidCssIdentifier(c))
                    .Select(c => $".{EscapeCssClass(c)}");

                if (classes.Any())
                {
                    strategies.Add(string.Join("", classes.Take(3))); // Combine up to 3 classes
                    strategies.Add(classes.First()); // Single most specific class
                }
            }

            // Strategy 3: Tag + Class combination
            if (!string.IsNullOrEmpty(element.TagName))
            {
                var tag = element.TagName.ToLower();

                if (!string.IsNullOrEmpty(element.HtmlClassName))
                {
                    var firstClass = element.HtmlClassName.Split(' ').FirstOrDefault(c => !string.IsNullOrWhiteSpace(c));
                    if (!string.IsNullOrEmpty(firstClass) && IsValidCssIdentifier(firstClass))
                    {
                        strategies.Add($"{tag}.{EscapeCssClass(firstClass)}");
                    }
                }

                // Strategy 4: Attribute selectors
                if (!string.IsNullOrEmpty(element.Name))
                {
                    strategies.Add($"{tag}[name='{EscapeCssAttribute(element.Name)}']");
                }

                if (!string.IsNullOrEmpty(element.InputType))
                {
                    strategies.Add($"input[type='{EscapeCssAttribute(element.InputType)}']");
                }

                if (!string.IsNullOrEmpty(element.Href))
                {
                    strategies.Add($"a[href='{EscapeCssAttribute(element.Href)}']");

                    // Partial match for dynamic URLs
                    var hrefPath = GetPathFromUrl(element.Href);
                    if (!string.IsNullOrEmpty(hrefPath))
                    {
                        strategies.Add($"a[href*='{EscapeCssAttribute(hrefPath)}']");
                    }
                }

                // Strategy 5: Pseudo-selectors for position
                if (element.ChildIndex >= 0)
                {
                    strategies.Add($"{tag}:nth-of-type({element.ChildIndex + 1})");

                    if (element.ChildIndex == 0)
                    {
                        strategies.Add($"{tag}:first-of-type");
                    }
                }

                // Strategy 6: ARIA selectors
                if (!string.IsNullOrEmpty(element.AriaLabel))
                {
                    strategies.Add($"[aria-label='{EscapeCssAttribute(element.AriaLabel)}']");
                }

                if (!string.IsNullOrEmpty(element.Role))
                {
                    strategies.Add($"[role='{EscapeCssAttribute(element.Role)}']");
                }
            }

            // Strategy 7: Data attribute selectors
            if (element.DataAttributes != null && element.DataAttributes.Count > 0)
            {
                foreach (var attr in element.DataAttributes.Take(2))
                {
                    if (IsValidCssIdentifier(attr.Key))
                    {
                        strategies.Add($"[{attr.Key}='{EscapeCssAttribute(attr.Value)}']");
                    }
                }
            }

            // Strategy 8: Combined selectors for uniqueness
            if (!string.IsNullOrEmpty(element.ParentName) && !string.IsNullOrEmpty(element.TagName))
            {
                var parentTag = element.ParentName.ToLower();
                var childTag = element.TagName.ToLower();
                strategies.Add($"{parentTag} > {childTag}");

                if (!string.IsNullOrEmpty(element.HtmlClassName))
                {
                    var firstClass = element.HtmlClassName.Split(' ').FirstOrDefault();
                    if (!string.IsNullOrEmpty(firstClass) && IsValidCssIdentifier(firstClass))
                    {
                        strategies.Add($"{parentTag} > {childTag}.{EscapeCssClass(firstClass)}");
                    }
                }
            }

            // Remove duplicates and return
            return strategies.Distinct().ToList();
        }

        /// <summary>
        /// Gets the best XPath for an element
        /// </summary>
        public static string GetOptimalXPath(ElementInfo element)
        {
            var strategies = GenerateXPathStrategies(element);

            // Return ID-based if available
            if (strategies.Any(s => s.Contains("@id=")))
                return strategies.First(s => s.Contains("@id="));

            // Return shortest unique selector
            return strategies.OrderBy(s => s.Length).FirstOrDefault() ?? $"//*";
        }

        /// <summary>
        /// Gets the best CSS selector for an element
        /// </summary>
        public static string GetOptimalCssSelector(ElementInfo element)
        {
            var strategies = GenerateCssSelectorStrategies(element);

            // Return ID-based if available
            if (strategies.Any(s => s.StartsWith("#")))
                return strategies.First(s => s.StartsWith("#"));

            // Return shortest selector
            return strategies.OrderBy(s => s.Length).FirstOrDefault() ?? "*";
        }

        #region Helper Methods

        private static string EscapeXPath(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            // Handle quotes in XPath
            if (value.Contains("'"))
            {
                // Use concat for strings with single quotes
                var parts = value.Split('\'');
                var escaped = string.Join("', \"'\", '", parts);
                return $"concat('{escaped}')";
            }

            return value;
        }

        private static string EscapeCssId(string id)
        {
            // Escape special CSS characters in ID
            return Regex.Replace(id, @"([:#\.\[\]@!""$%&'()*+,/;<=>?\\^`{|}~])", @"\$1");
        }

        private static string EscapeCssClass(string className)
        {
            // Escape special CSS characters in class name
            return Regex.Replace(className, @"([:#\.\[\]@!""$%&'()*+,/;<=>?\\^`{|}~])", @"\$1");
        }

        private static string EscapeCssAttribute(string value)
        {
            // Escape quotes in attribute values
            return value?.Replace("'", "\\'") ?? "";
        }

        private static bool IsValidCssIdentifier(string identifier)
        {
            if (string.IsNullOrEmpty(identifier))
                return false;

            // Check if it's a valid CSS identifier
            return Regex.IsMatch(identifier, @"^[a-zA-Z_][-a-zA-Z0-9_]*$");
        }

        private static string GetPathFromUrl(string url)
        {
            try
            {
                var uri = new Uri(url);
                return uri.AbsolutePath;
            }
            catch
            {
                return "";
            }
        }

        private static string GetFilenameFromPath(string path)
        {
            try
            {
                return System.IO.Path.GetFileName(path);
            }
            catch
            {
                return "";
            }
        }

        #endregion

        #region UI Automation Specific

        /// <summary>
        /// Generates Windows UI Automation path
        /// </summary>
        public static string GenerateWindowsPath(AutomationElement element)
        {
            if (element == null)
                return "";

            var path = new StringBuilder();
            var current = element;
            var pathParts = new List<string>();

            try
            {
                while (current != null && current != AutomationElement.RootElement)
                {
                    var controlType = current.Current.ControlType;
                    var name = current.Current.Name;
                    var automationId = current.Current.AutomationId;
                    var className = current.Current.ClassName;

                    var part = controlType?.ProgrammaticName ?? "Unknown";

                    if (!string.IsNullOrEmpty(automationId))
                    {
                        part += $"[@AutomationId='{automationId}']";
                    }
                    else if (!string.IsNullOrEmpty(name))
                    {
                        part += $"[@Name='{name}']";
                    }
                    else if (!string.IsNullOrEmpty(className))
                    {
                        part += $"[@ClassName='{className}']";
                    }

                    pathParts.Insert(0, part);

                    // Get parent
                    current = TreeWalker.ControlViewWalker.GetParent(current);
                }

                return "/" + string.Join("/", pathParts);
            }
            catch
            {
                return "";
            }
        }

        #endregion
    }
}