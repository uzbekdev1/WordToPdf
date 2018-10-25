using System;
using System.Diagnostics;
using System.IO;

namespace WordToPdf.docto
{
    internal static class ProcessHelper
    {
        internal static void Run(string fileName, string arguments, bool showResultsInConsole)
        {
            // Start the child process.
            using (var p = new Process
            {
                StartInfo =
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    FileName = fileName
                }
            })
            {
                // Redirect the output stream of the child process.
                if (!string.IsNullOrEmpty(arguments))
                    p.StartInfo.Arguments = arguments;

                if (!string.IsNullOrEmpty(Path.GetDirectoryName(fileName)))
                    p.StartInfo.WorkingDirectory = Path.GetDirectoryName(fileName);
                p.Start();
                // Do not wait for the child process to exit before
                // reading to the end of its redirected stream.
                // p.WaitForExit();

                // Read the output stream first and then wait.
                var output = p.StandardOutput.ReadToEnd();
                if (showResultsInConsole)
                    Console.WriteLine(output);
                p.WaitForExit();
            }
        }
    }
}