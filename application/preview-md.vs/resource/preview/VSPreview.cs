using Markdig;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace resource.preview
{
    internal class VSPreview : extension.AnyPreview
    {
        protected override void _Execute(atom.Trace context, int level, string url, string file)
        {
            var a_Name = atom.Trace.GetUrlTemp(file, ".png");
            {
                context.
                    SetAlignment(NAME.ALIGNMENT.TOP).
                    SetTrace(null, NAME.STATE.TRACE.BLINK).
                    SetProgress(CONSTANT.PROGRESS.INFINITE).
                    SetUrlPreview(a_Name).
                    SendPreview(NAME.EVENT.INFO, url);
            }
            {
                context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.HEADER, level, "[[[Info]]]");
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[File Name]]]", url);
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[File Size]]]", (new FileInfo(file)).Length.ToString());
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[Raw Format]]]", "Markdown");
                }
            }
            {
                var a_Context = new Thread(__BrowserThread);
                {
                    a_Context.SetApartmentState(ApartmentState.STA);
                    a_Context.Start(new Tuple<string, string, string, int>(url, file, a_Name, level));
                }
            }
        }

        private static string __GetUrl(object data)
        {
            var a_Context = data as Tuple<string, string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item1?.ToString();
            }
            return null;
        }

        private static string __GetUrlLocal(object data)
        {
            var a_Context = data as Tuple<string, string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item2?.ToString();
            }
            return null;
        }

        private static string __GetUrlProxy(object data)
        {
            var a_Context = data as Tuple<string, string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item3?.ToString();
            }
            return null;
        }

        private static int __GetLevel(object data)
        {
            var a_Context = data as Tuple<string, string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item4;
            }
            return 0;
        }

        private static string __GetHtml(object data)
        {
            var a_Context = new MarkdownPipelineBuilder().
                UseAdvancedExtensions().
                UseBootstrap().
                UseEmojiAndSmiley().
                UseSmartyPants().
                Build();
            {
                return Markdown.ToHtml(File.ReadAllText(__GetUrlLocal(data)), a_Context);
            }
        }

        private static void __BrowserThread(object context)
        {
            try
            {
                var a_Context = new WebBrowser();
                {
                    a_Context.Tag = context;
                    a_Context.ScrollBarsEnabled = false;
                    a_Context.ScriptErrorsSuppressed = true;
                    a_Context.IsWebBrowserContextMenuEnabled = true;
                    a_Context.AllowNavigation = true;
                    a_Context.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(__DocumentCompleted);
                    a_Context.DocumentText = __GetHtml(context);
                }
                while (a_Context.Tag != null)
                {
                    Application.DoEvents();
                    Thread.Sleep(50);
                }
                {
                    a_Context.Dispose();
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                atom.Trace.GetInstance().
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.EXCEPTION, __GetLevel(context), ex.Message).
                    SetAlignment(NAME.ALIGNMENT.TOP).
                    SetFont(null, 0, NAME.STATE.FONT.NONE).
                    SetUrlPreview(__GetUrlLocal(context)).
                    SendPreview(NAME.EVENT.EXCEPTION, __GetUrl(context));
            }
        }

        private static void __DocumentCompleted(object context, WebBrowserDocumentCompletedEventArgs e)
        {
            var a_Context = (WebBrowser)context;
            if (__GetUrl(a_Context?.Tag) != null)
            {
                var a_Context1 = a_Context.Tag;
                try
                {
                    {
                        a_Context.Tag = true;
                        a_Context.Width = 1000;// GetProperty(NAME.PROPERTY.PREVIEW_WIDTH, true);
                        a_Context.Height = a_Context.Document.Body.ScrollRectangle.Height;
                    }
                    {
                        var a_Context2 = new Bitmap(a_Context.Width, a_Context.Height);
                        {
                            a_Context.DrawToBitmap(a_Context2, new Rectangle(0, 0, a_Context2.Width, a_Context2.Height));
                        }
                        {
                            a_Context2.Save(__GetUrlProxy(a_Context1), System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                    {
                        var a_Count = (a_Context.Height + CONSTANT.OUTPUT.PREVIEW_ITEM_HEIGHT + CONSTANT.OUTPUT.PREVIEW_PAGE_INDENT + CONSTANT.OUTPUT.PREVIEW_PAGE_INDENT) / (CONSTANT.OUTPUT.PREVIEW_ITEM_HEIGHT + 1);
                        {
                            atom.Trace.GetInstance().
                                SetControl(NAME.CONTROL.BROWSER).
                                SetCount(Math.Max(a_Count, CONSTANT.OUTPUT.PREVIEW_MIN_SIZE)).
                                Send(NAME.SOURCE.PREVIEW, NAME.EVENT.CONTROL, __GetLevel(a_Context1));
                        }
                    }
                    {
                        atom.Trace.GetInstance().
                            Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOOTER, __GetLevel(a_Context1), "[[[Size]]]: " + (new FileInfo(__GetUrlLocal(a_Context1))).Length.ToString()).
                            SetAlignment(NAME.ALIGNMENT.TOP).
                            SetFont(null, 0, NAME.STATE.FONT.NONE).
                            SetProgress(100).
                            SetUrlPreview(__GetUrlProxy(a_Context1)).
                            SendPreview(NAME.EVENT.INFO, __GetUrl(a_Context1));
                    }
                    {
                        a_Context.Tag = null;
                    }
                }
                catch (Exception ex)
                {
                    {
                        a_Context.Tag = null;
                    }
                    {
                        atom.Trace.GetInstance().
                            Send(NAME.SOURCE.PREVIEW, NAME.EVENT.EXCEPTION, __GetLevel(a_Context1), ex.Message).
                            SetFont(null, 0, NAME.STATE.FONT.NONE).
                            SetProgress(100).
                            SendPreview(NAME.EVENT.EXCEPTION, __GetUrl(a_Context1));
                    }
                }
            }
        }
    };
}
