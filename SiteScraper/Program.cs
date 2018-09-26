using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using WebBrowser = System.Windows.Forms.WebBrowser;

namespace SiteScraper
{
    internal class Program
    {
        private const string PreText = "1280,\"mime\":\"video/mp4\",\"fps\":30,\"url\":\"";
        private const string PreText2 = "1280,\"mime\":\"video/mp4\",\"fps\":25,\"url\":\"";
        private const string AfterText = "\",\"cdn\"";
        private const string FileName = "resourceList.txt";

        private static void Main()
        {
            Console.WriteLine("Starting SiteScraper...");

            try
            {
                // read urls
                var lines = File.ReadLines(FileName);

                // ReSharper disable once CoVariantArrayConversion
                object[] list = lines.ToArray();

                // download each page and dump the content
                var task = MessageLoopWorker.Run(DoWorkAsync, list);
                task.Wait();

                Console.WriteLine("DoWorkAsync completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("DoWorkAsync failed: " + ex.Message);
            }

            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }

        // navigate WebBrowser to the list of urls in a loop
        private static async Task<object> DoWorkAsync(object[] args)
        {
            Console.WriteLine("Start working.");

            using (var wb = new WebBrowser())
            {
                wb.ScriptErrorsSuppressed = true;

                TaskCompletionSource<bool> tcs;

                // ReSharper disable once AccessToModifiedClosure
                void DocumentCompletedHandler(object s, WebBrowserDocumentCompletedEventArgs e) => tcs.TrySetResult(true);

                // navigate to each URL in the list
                foreach (var url in args)
                {
                    var u = url.ToString();

                    tcs = new TaskCompletionSource<bool>();
                    wb.DocumentCompleted += DocumentCompletedHandler;

                    try
                    {
                        wb.Navigate(u);
                        // await for DocumentCompleted

                        await tcs.Task;
                    }
                    finally
                    {
                        wb.DocumentCompleted -= DocumentCompletedHandler;
                    }
                    // the DOM is ready
                    Console.WriteLine(u);
                    var htmla = wb.Document?.Body?.OuterHtml;
                    Debug.WriteLine(htmla);

                    var linkit = wb.Document?.GetElementsByTagName("iframe");

                    if (linkit == null) continue;

                    for (var i = 0; i < linkit.Count; i++)
                    {
                        try
                        {
                            var doc = (IHTMLDocument2) wb.Document.DomDocument;

                            var window = (IHTMLWindow2) doc.frames.item(i);

                            while (true)
                            {
                                // wait asynchronously, this will throw if cancellation requested
                                await Task.Delay(500);

                                // continue polling if the WebBrowser is still busy
                                if (wb.IsBusy)
                                    continue;

                                var windowNow = (IHTMLWindow2) doc.frames.item(i);
                                if (window == windowNow)
                                    break; // no changes detected, end the poll loop

                                window = windowNow;
                            }

                            var frame = CrossFrameIe.GetDocumentFromWindow((IHTMLWindow2) doc.frames.item(i));
                            var html = frame.body.outerHTML;

                            var videoUrl = GetBetween(html, PreText, AfterText);
                            if (string.IsNullOrEmpty(videoUrl))
                                videoUrl = GetBetween(html, PreText2, AfterText);

                            Console.WriteLine(videoUrl);

                            Download(u, videoUrl, i);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Could not download video {i} from {u}. Error: {ex}");
                        }
                    }
                }
            }

            Console.WriteLine("End working.");
            return null;
        }

        private static void Download(string u, string videoUrl, int i)
        {
            using (var client = new WebClient())
            {
                var pos = u.LastIndexOf("/", StringComparison.Ordinal) + 1;
                var folderName = u.Substring(pos, u.Length - pos);

                Directory.CreateDirectory(folderName);

                client.DownloadFile(new Uri(videoUrl), $@"{folderName}\{folderName}-{i + 1}.mp4");
            }
        }

        public static string GetBetween(string strSource, string strStart, string strEnd)
        {
            if (!strSource.Contains(strStart) || !strSource.Contains(strEnd)) return string.Empty;

            var start = strSource.IndexOf(strStart, 0, StringComparison.InvariantCultureIgnoreCase) + strStart.Length;
            var end = strSource.IndexOf(strEnd, start, StringComparison.InvariantCultureIgnoreCase);
            return strSource.Substring(start, end - start);
        }

    }
}
