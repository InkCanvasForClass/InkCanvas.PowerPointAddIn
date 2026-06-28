using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using PPTAgent.Contracts;
using Newtonsoft.Json;

namespace PPTAgent.AddIn.Core
{
    public sealed class PPTController
    {
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }
        private readonly Microsoft.Office.Interop.PowerPoint.Application _application;
        private SynchronizationContext _syncContext;

        public PPTController(Microsoft.Office.Interop.PowerPoint.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            CaptureSyncContext();
        }

        public PPTState GetState()
        {
            var state = new PPTState();

            try
            {
                if (_application.Presentations.Count > 0)
                {
                    var pres = _application.ActivePresentation;
                    state.PresentationName = pres.Name;
                    try { state.PresentationFullName = pres.FullName; } catch { }
                    state.TotalSlides = pres.Slides.Count;
                    state.HasHiddenSlides = HasHiddenSlides(pres);
                    state.HasAutoPlayTimings = HasAutoPlayTimings(pres);
                }
            }
            catch { }

            try
            {
                if (_application.SlideShowWindows.Count > 0)
                {
                    state.IsRunning = true;
                    state.SlideIndex = _application.SlideShowWindows[1].View.CurrentShowPosition;
                }
            }
            catch { }

            return state;
        }

        public bool Next()
        {
            return Run(() =>
            {
                if (_application.SlideShowWindows.Count <= 0) return false;
                _application.SlideShowWindows[1].View.Next();
                return true;
            });
        }

        public bool Previous()
        {
            return Run(() =>
            {
                if (_application.SlideShowWindows.Count <= 0) return false;
                _application.SlideShowWindows[1].View.Previous();
                return true;
            });
        }

        public bool GotoSlide(int slideNumber)
        {
            return Run(() =>
            {
                if (slideNumber <= 0) return false;
                if (_application.SlideShowWindows.Count <= 0) return false;
                _application.SlideShowWindows[1].View.GotoSlide(slideNumber);
                return true;
            });
        }

        public bool StartSlideShow()
        {
            return Run(() =>
            {
                if (_application.Presentations.Count <= 0) return false;
                _application.ActivePresentation.SlideShowSettings.Run();
                return true;
            });
        }

        public bool EndSlideShow()
        {
            return Run(() =>
            {
                if (_application.SlideShowWindows.Count <= 0) return false;
                _application.SlideShowWindows[1].View.Exit();
                return true;
            });
        }

        public bool ShowSlideNavigation()
        {
            return Run(() =>
            {
                if (_application.SlideShowWindows.Count <= 0) return false;
                try
                {
                    dynamic nav = _application.SlideShowWindows[1].SlideNavigation;
                    if (nav == null) return false;
                    nav.Visible = true;
                    return true;
                }
                catch { return false; }
            });
        }

        public bool DisableAutoPlayTimings()
        {
            return Run(() =>
            {
                if (_application.Presentations.Count <= 0) return false;
                _application.ActivePresentation.SlideShowSettings.AdvanceMode =
                    Microsoft.Office.Interop.PowerPoint.PpSlideShowAdvanceMode.ppSlideShowManualAdvance;
                return true;
            });
        }

        public bool UnhideHiddenSlides()
        {
            return Run(() =>
            {
                if (_application.Presentations.Count <= 0) return false;
                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in _application.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.Hidden == Microsoft.Office.Core.MsoTriState.msoTrue)
                        slide.SlideShowTransition.Hidden = Microsoft.Office.Core.MsoTriState.msoFalse;
                }
                return true;
            });
        }

        public MediaRegionsResponse GetMediaRegions()
        {
            return Run(() =>
            {
                var response = new MediaRegionsResponse();

                if (_application.SlideShowWindows.Count <= 0)
                    return response;

                var ssw = _application.SlideShowWindows[1];
                var view = ssw.View;
                if (view == null) return response;

                response.SlideIndex = view.CurrentShowPosition;

                var slide = view.Slide;
                if (slide == null) return response;

                // Get slide show window physical coordinates and DPI via Win32 API
                IntPtr hwnd = FindWindow("screenClass", null);
                response.SlideShowWindowHandle = hwnd.ToInt64();

                double winLeft, winTop, winWidth, winHeight;
                uint dpi = 96;
                if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
                {
                    winLeft = rect.Left;
                    winTop = rect.Top;
                    winWidth = rect.Right - rect.Left;
                    winHeight = rect.Bottom - rect.Top;
                    try { dpi = GetDpiForWindow(hwnd); } catch { dpi = 96; }
                }
                else
                {
                    // Fallback: use COM properties (units are points, need DPI conversion)
                    double dpiFactor = 96.0 / 72.0; // Default assume 100% DPI
                    winLeft = ssw.Left * dpiFactor;
                    winTop = ssw.Top * dpiFactor;
                    winWidth = ssw.Width * dpiFactor;
                    winHeight = ssw.Height * dpiFactor;
                }

                var pres = ssw.Presentation;
                float slideWidth = pres.PageSetup.SlideWidth;   // points
                float slideHeight = pres.PageSetup.SlideHeight; // points
                response.SlideWidth = slideWidth;
                response.SlideHeight = slideHeight;

                // Return shape's raw point coordinates; host app does screen pixel conversion
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (!IsMediaShape(shape)) continue;

                    try
                    {
                        var region = new MediaRegion
                        {
                            ScreenX = shape.Left,
                            ScreenY = shape.Top,
                            ScreenWidth = shape.Width,
                            ScreenHeight = shape.Height,
                            ShapeName = shape.Name,
                            MediaType = (int)shape.MediaType
                        };
                        response.Regions.Add(region);
                    }
                    catch
                    {
                        // Some Shape properties may not be accessible, skip
                    }
                }

                return response;
            }) ?? new MediaRegionsResponse();
        }

        private static bool IsMediaShape(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            try
            {
                // msoMedia = 16
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                    return true;

                // OLE controls (ActiveX media player, etc.)
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoOLEControlObject)
                    return true;

                // Embedded video (legacy format)
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject)
                {
                    try
                    {
                        if (shape.MediaType == Microsoft.Office.Interop.PowerPoint.PpMediaType.ppMediaTypeMovie ||
                            shape.MediaType == Microsoft.Office.Interop.PowerPoint.PpMediaType.ppMediaTypeSound)
                            return true;
                    }
                    catch { }
                }
            }
            catch { }

            return false;
        }

        private void CaptureSyncContext()
        {
            if (SynchronizationContext.Current != null)
                _syncContext = SynchronizationContext.Current;
        }

        private bool Run(Func<bool> action)
        {
            if (_syncContext != null)
            {
                bool result = false;
                Exception captured = null;
                _syncContext.Send(_ =>
                {
                    try { result = action(); }
                    catch (Exception ex) { captured = ex; }
                }, null);
                if (captured != null) throw captured;
                return result;
            }
            return action.Invoke();
        }

        private T Run<T>(Func<T> action)
        {
            if (_syncContext != null)
            {
                T result = default;
                Exception captured = null;
                _syncContext.Send(_ =>
                {
                    try { result = action(); }
                    catch (Exception ex) { captured = ex; }
                }, null);
                if (captured != null) throw captured;
                return result;
            }
            return action.Invoke();
        }

        private static bool HasHiddenSlides(Microsoft.Office.Interop.PowerPoint.Presentation pres)
        {
            try
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in pres.Slides)
                {
                    if (slide.SlideShowTransition.Hidden == Microsoft.Office.Core.MsoTriState.msoTrue)
                        return true;
                }
            }
            catch { }
            return false;
        }

        private static bool HasAutoPlayTimings(Microsoft.Office.Interop.PowerPoint.Presentation pres)
        {
            try
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in pres.Slides)
                {
                    if (slide.SlideShowTransition.AdvanceOnTime == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        slide.SlideShowTransition.AdvanceTime > 0)
                        return true;
                }
            }
            catch { }
            return false;
        }
    }
}
