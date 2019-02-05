using Args;
using Args.Help;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Generator
{
    class Program
    {
        private static string extTextReplacer = "ReplaceText.exe", argsTextReplacer = "",
            extVideoRenderer = "RenderToVideo.exe", argsVideoRenderer = "",
            extMusicAttacher = "env\\Scripts\\python.exe", argsMusicAttacher = "compose.py";


        static void Main(string[] args)
        {
            try
            {
                var command = Configuration.Configure<CommandObject>().CreateAndBind(args);

                if (command.output != null && command.input != null && command.music != null && command.template != null)
                {
                    foreach (var txt in Directory.EnumerateFiles(command.input, "*.txt"))
                    {
                        Console.WriteLine("Preparing: " + txt);
                        var tmp = Path.GetTempFileName();
                        File.Delete(tmp);
                        Directory.CreateDirectory(tmp);

                        // Copy all files with the same name
                        foreach (var f in Directory.EnumerateFiles(command.input, Path.GetFileNameWithoutExtension(txt) + "*.*", SearchOption.TopDirectoryOnly))
                        {
                            var filename = Path.GetFileName(f);
                            var r = Path.GetFileNameWithoutExtension(txt) + "-";
                            if (filename.StartsWith(r))
                                filename = filename.Substring(r.Length);

                            File.Copy(f, Path.Combine(tmp,filename));
                        }

                        var tmpTXT = Path.Combine(tmp, Path.GetFileName(txt));
                        var tmpPPTX = Path.ChangeExtension(tmpTXT, ".pptx");
                        var tmpMP4 = Path.ChangeExtension(tmpTXT, ".mp4");

                        // Only do 1... because powerpoint
                        Console.WriteLine("Replacing presentation variables.");
                        Process.Start(extTextReplacer, argsTextReplacer + $" \"{Path.GetFullPath(command.template)}\" \"{Path.GetFullPath(tmpPPTX)}\" \"{Path.GetFullPath(txt)}\"").WaitForExit();
                        Console.WriteLine("Rendering video.");
                        Process.Start(extVideoRenderer, argsVideoRenderer + $" \"{Path.GetFullPath(tmpPPTX)}\" \"{Path.GetFullPath(tmpMP4)}\"").WaitForExit();
                        Console.WriteLine("Creating output video.");
                        Process.Start(extMusicAttacher, argsMusicAttacher +
                            $" \"{Path.GetFullPath(tmpMP4)}\" \"{Path.GetFullPath(command.music)}\" \"{Path.Combine(Path.GetFullPath(command.output), Path.GetFileName(tmpMP4))}\"");
                    }

                    Console.WriteLine("All done. Wait for the processes to exit");
                }
                else
                {
                    var result = new HelpProvider().GenerateModelHelp(Configuration.Configure<CommandObject>());
                    Console.WriteLine("Not enough parameters. Please provide:");

                    foreach (var t in result.Members)
                    {

                        Console.WriteLine("\t" + result.SwitchDelimiter + String.Join(" " + result.SwitchDelimiter, t.Switches.Reverse()) + "\t" + t.HelpText);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }       
        }
    }

    public class CommandObject
    {
        [System.ComponentModel.Description("Folder containing the input files")]
        public string input { get; set; }
        [System.ComponentModel.Description("Folder to output the result files")]
        public string output { get; set; }
        [System.ComponentModel.Description("Template PPTX")]
        public string template { get; set; }
        [System.ComponentModel.Description("Audio track")]
        public string music { get; set; }
    }
}
