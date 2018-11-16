using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReplaceText
{
    class Program
    {
        private const bool caseSensitive = true;

        static void Main(string[] args)
        {
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (args.Length < 3)
            {
                Console.WriteLine("Usage:\n" + Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location) +
                    " input_powerpoint.pptx output_powerpoint.pptx input_variables.txt");
            }
            else
            {
                try
                {
                    Console.WriteLine($"Processing: {args[0]}");
                    ProcessPowerPoint(ParseReplacements(args[2]), args[0], args[1]);
                    Console.WriteLine("Done.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"General error: {ex.Message}");
                }
            }
        }

        private static void ProcessPowerPoint(IEnumerable<SearchReplace> replacements, string presentation, string output)
        {
            Console.WriteLine("Starting PowerPoint ...");
            Application ppApp = InitializePowerPoint();

            if (ppApp == null)
                return;

            // Replace strings
            var processed = false;
            Presentation p = null;

            try
            {
                p = ppApp.Presentations.Open(Path.GetFullPath(presentation),
                    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                try
                {
                    foreach (var replacement in replacements)
                        for (int slide = 1; slide <= p.Slides.Count; slide++)
                            for (int shape = 1; shape <= p.Slides[slide].Shapes.Count; shape++)
                            {
                                if (p.Slides[slide].Shapes[shape].HasTextFrame == MsoTriState.msoTrue)
                                {
                                    p.Slides[slide].Shapes[shape].TextFrame.TextRange.Replace(replacement.Search,
                                        replacement.Replace, 0, caseSensitive ? MsoTriState.msoTrue : MsoTriState.msoFalse);

                                    processed = true;
                                }

                                /*if (p.Slides[slide].Shapes[shape].HasTable == MsoTriState.msoTrue)
                                {
                                    for (int row = 1; row <= p.Slides[slide].Shapes[shape].Table.Rows.Count; row++)
                                        for (int column = 1; column <= p.Slides[slide].Shapes[shape].Table.Columns.Count; column++)

                                            if (p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.HasTextFrame == MsoTriState.msoTrue)
                                            {
                                                previous = p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.TextFrame.TextRange.Text;

                                                if (!previous.StartsWith(secret))
                                                {
                                                    p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.TextFrame.TextRange.Replace(replacement.Search,
                                                        replacement.Replace, 0, caseSensitive ? MsoTriState.msoTrue : MsoTriState.msoFalse);

                                                    if (previous.CompareTo(p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.TextFrame.TextRange.Text) != 0)
                                                    {
                                                        // There are changes
                                                        p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.TextFrame.TextRange.InsertBefore(secret);
                                                        processed = true;
                                                    }
                                                }
                                            }

                                    // Cleaning
                                    for (int row = 1; row <= p.Slides[slide].Shapes[shape].Table.Rows.Count; row++)
                                        for (int column = 1; column <= p.Slides[slide].Shapes[shape].Table.Columns.Count; column++)

                                            if (p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.HasTextFrame == MsoTriState.msoTrue)
                                            {
                                                p.Slides[slide].Shapes[shape].Table.Cell(row, column).Shape.TextFrame.TextRange.Replace(secret,
                                                        "", 0, MsoTriState.msoTrue);
                                            }
                                }*/
                            }

                    if (processed)
                        if (output == presentation)
                            p.Save();
                        else
                            p.SaveAs(Path.GetFullPath(output));
                }
                catch
                {
                    processed = false;
                }

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

        private static IEnumerable<SearchReplace> ParseReplacements(string v)
        {
            foreach (var line in File.ReadAllLines(v))
            {
                if (line.StartsWith("#"))
                    continue;

                var s = line.Split('=');
                yield return new SearchReplace { Search = s[0], Replace = s[1] };
            }
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

        class SearchReplace
        {
            public string Search, Replace;
        }
    }
}
