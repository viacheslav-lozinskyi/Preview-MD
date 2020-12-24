using Markdig;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace resource.preview
{
    internal class VSPreview : cartridge.AnyPreview
    {
        protected override void _Execute(atom.Trace context, string url, int level)
        {
            if (File.Exists(url))
            {
                var a_Name = GetUrlProxy(url, ".png");
                {
                    context.
                        SetProgress(0, true, "").
                        SetUrlAlignment(NAME.ALIGNMENT.TOP).
                        SetUrlProxy(a_Name).
                        SendPreview(NAME.TYPE.INFO, url);
                }
                {
                    context.
                        SetState(NAME.STATE.HEADER).
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, level, "[[Info]]");
                    {
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, level + 1, "[[File Name]]", url);
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, level + 1, "[[File Size]]", (new System.IO.FileInfo(url)).Length.ToString());
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, level + 1, "[[Raw Format]]", "Markdown");
                    }
                }
                {
                    var a_Context = new Thread(__BrowserThread);
                    {
                        a_Context.SetApartmentState(ApartmentState.STA);
                        a_Context.Start(new Tuple<string, string, int>(url, a_Name, level));
                    }
                }
            }
            else
            {
                context.
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.ERROR, level, "[[File not found]]").
                    SendPreview(NAME.TYPE.ERROR, url);
            }
        }

        private static string __GetUrl(object context)
        {
            var a_Context = context as Tuple<string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item1?.ToString();
            }
            return null;
        }

        private static string __GetUrlProxy(object context)
        {
            var a_Context = context as Tuple<string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item2?.ToString();
            }
            return null;
        }

        private static int __GetLevel(object context)
        {
            var a_Context = context as Tuple<string, string, int>;
            if (a_Context != null)
            {
                return a_Context.Item3;
            }
            return 0;
        }

        private static string __GetHtml(object context)
        {
            var a_Context = new MarkdownPipelineBuilder().
                UseAdvancedExtensions().
                UseBootstrap().
                UseEmojiAndSmiley().
                UseSmartyPants().
                Build();
            {
                return Markdown.ToHtml(File.ReadAllText(__GetUrl(context)), a_Context);
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
                {
                    Application.Run();
                }
                {
                    a_Context.Dispose();
                }
            }
            catch (Exception ex)
            {
                atom.Trace.GetInstance().
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.ERROR, __GetLevel(context), ex.Message).
                    SetUrlAlignment(NAME.ALIGNMENT.TOP).
                    SetUrlProxy(__GetUrlProxy(context)).
                    SendPreview(NAME.TYPE.ERROR, __GetUrl(context));
            }
        }

        private static void __DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var a_Context = (WebBrowser)sender;
            if (a_Context?.Tag != null)
            {
                var a_Context1 = a_Context.Tag;
                try
                {
                    {
                        a_Context.Tag = null;
                    }
                    {
                        var a_Size = a_Context.Document.Body.ScrollRectangle.Width;
                        {
                            a_Size = Math.Max(a_Size, GetProperty(NAME.PROPERTY.PREVIEW_WIDTH));
                        }
                        {
                            a_Context.Width = a_Size;
                            a_Context.Height = a_Context.Document.Body.ScrollRectangle.Height;
                        }
                        {
                            var a_Context2 = new Bitmap(a_Size, a_Context.Document.Body.ScrollRectangle.Height);
                            {
                                a_Context.DrawToBitmap(a_Context2, new Rectangle(0, 0, a_Context2.Width, a_Context2.Height));
                            }
                            {
                                a_Context2.Save(__GetUrlProxy(a_Context1), System.Drawing.Imaging.ImageFormat.Png);
                            }
                        }
                    }
                    {
                        Application.ExitThread();
                    }
                    {
                        var a_Size = GetProperty(NAME.PROPERTY.PREVIEW_MEDIA_SIZE);
                        {
                            a_Size = Math.Min(a_Size, a_Context.Document.Body.ScrollRectangle.Height / CONSTANT.OUTPUT_PREVIEW_ITEM_HEIGHT);
                            a_Size = Math.Max(a_Size, CONSTANT.OUTPUT_PREVIEW_MIN_SIZE);
                        }
                        for (var i = 0; i < a_Size; i++)
                        {
                            atom.Trace.GetInstance().Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PREVIEW, __GetLevel(a_Context1));
                        }
                    }
                    {
                        atom.Trace.GetInstance().
                            SetState(NAME.STATE.FOOTER).
                            Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, __GetLevel(a_Context1), "[[Size]]: " + (new FileInfo(__GetUrl(a_Context1))).Length.ToString()).
                            SetUrlAlignment(NAME.ALIGNMENT.TOP).
                            SetUrlProxy(__GetUrlProxy(a_Context1)).
                            SendPreview(NAME.TYPE.INFO, __GetUrl(a_Context1));
                    }
                }
                catch (Exception ex)
                {
                    atom.Trace.GetInstance().
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.ERROR, __GetLevel(a_Context1), ex.Message).
                        SetUrlAlignment(NAME.ALIGNMENT.TOP).
                        SetUrlProxy(__GetUrlProxy(a_Context1)).
                        SendPreview(NAME.TYPE.ERROR, __GetUrl(a_Context1));
                }
            }
        }
    };
}
