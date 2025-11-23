using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Base interface for all element detection technologies
    /// </summary>
    public interface IElementDetector
    {
        /// <summary>
        /// Name of the detection technology
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Checks if this detector can work with the current target
        /// </summary>
        bool CanDetect(System.Windows.Point screenPoint);

        /// <summary>
        /// Gets element at the specified screen point
        /// </summary>
        Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile);

        /// <summary>
        /// Gets all elements in the specified window or area
        /// </summary>
        Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile);

        /// <summary>
        /// Gets elements within a specified rectangle
        /// </summary>
        Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile);

        /// <summary>
        /// Builds element hierarchy tree
        /// </summary>
        Task<ElementInfo> GetElementTree(ElementInfo rootElement, CollectionProfile profile);

        /// <summary>
        /// Refreshes element information
        /// </summary>
        Task<ElementInfo> RefreshElement(ElementInfo element, CollectionProfile profile);
    }
}