using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RenderToVideo
{
    class Program
    {
        private static string lastMsg;

        static void Main(string[] args)
        {
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:\n" + Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location) +
                    " input_powerpoint.pptx output_video.mp4");
            }
            else
            {
                try
                {
                    Console.WriteLine($"Processing: {args[1]}");
                    ProcessPowerPoint(args[0], args[1]);
                    Console.WriteLine("Done.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"General error: {ex.Message}");
                }
            }
        }

        private static void ProcessPowerPoint(string presentation, string output)
        {
            Console.WriteLine("Starting PowerPoint ...");
            Application ppApp = InitializePowerPoint();

            if (ppApp == null)
                return;

            Presentation p = null;

            try
            {
                p = ppApp.Presentations.Open(Path.GetFullPath(presentation),
                    MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

                Console.WriteLine("Creating video ...");
                p.CreateVideo(Path.GetFullPath(output), true, 0, 2160, 45, 100);

                var processing = true;
                do
                {
                    Thread.Sleep(100);
                    switch (p.CreateVideoStatus)
                    {
                        case PpMediaTaskStatus.ppMediaTaskStatusInProgress:
                            WriteStatus("Working ...");
                            break;
                        case PpMediaTaskStatus.ppMediaTaskStatusDone:
                            WriteStatus("Done!");
                            processing = false;
                            break;
                        case PpMediaTaskStatus.ppMediaTaskStatusFailed:
                            WriteStatus("Failed");
                            processing = false;
                            break;
                        default:
                            WriteStatus("Waiting"); // Waiting or queueing, etc
                            break;
                    }

                } while (processing);

                // Close and terminate everything
                p.Close();
                NAR(p);

                try
                {
                    NAR(ppApp);
                    ppApp.Quit();
                }
                catch { }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while processing: {ex.Message}");
            }
        }

        private static void WriteStatus(string msg)
        {
            if (lastMsg == msg)
                return;

            lastMsg = msg;
            Console.WriteLine($"[Render] {msg}");
        }

        private static Microsoft.Office.Interop.PowerPoint.Application InitializePowerPoint()
        {
            Microsoft.Office.Interop.PowerPoint.Application ppApp;
            try
            {
                ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            }
            catch
            {
                ppApp = null;
            }
            return ppApp;
        }
 
        private static void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }
    }
}
