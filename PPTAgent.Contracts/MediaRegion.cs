using System.Collections.Generic;

namespace PPTAgent.Contracts
{
    /// <summary>
    /// Media control region on a PowerPoint slide (screen coordinates, pixels).
    /// </summary>
    public sealed class MediaRegion
    {
        /// <summary>Screen X (pixels)</summary>
        public double ScreenX { get; set; }
        /// <summary>Screen Y (pixels)</summary>
        public double ScreenY { get; set; }
        /// <summary>Width (pixels)</summary>
        public double ScreenWidth { get; set; }
        /// <summary>Height (pixels)</summary>
        public double ScreenHeight { get; set; }
        /// <summary>Shape name (for debugging)</summary>
        public string ShapeName { get; set; }
        /// <summary>Media type</summary>
        public int MediaType { get; set; }
    }

    /// <summary>
    /// Response containing media regions and slide show window handle.
    /// </summary>
    public sealed class MediaRegionsResponse
    {
        public List<MediaRegion> Regions { get; set; } = new List<MediaRegion>();
        public int SlideIndex { get; set; }
        /// <summary>Slide show window handle (for the host application)</summary>
        public long SlideShowWindowHandle { get; set; }
        /// <summary>Slide width (points)</summary>
        public float SlideWidth { get; set; }
        /// <summary>Slide height (points)</summary>
        public float SlideHeight { get; set; }
    }
}
