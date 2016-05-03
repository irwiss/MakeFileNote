namespace MakeFileNote
{
    using System;
    using File = System.IO.File;
    using Path = System.IO.Path;
    using Assembly = System.Reflection.Assembly;
    using Process = System.Diagnostics.Process;
    using RegistryKey = Microsoft.Win32.RegistryKey;

    internal class MakeFileNote
    {
        private const string HKCU_Classes_Shell_MakeFileNote = @"Software\Classes\*\shell\MakeFileNote";
        private const string HKCU_Classes_Shell_MakeFileNote_Command = HKCU_Classes_Shell_MakeFileNote + @"\command";
        private const string Title = "Make File Note";

        private static readonly RegistryKey HKCU = Microsoft.Win32.Registry.CurrentUser;
        private static string TargetExecutable = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "makefilenote.exe");

        /// <summary>Flag to break out of user input loop when uninstalling.</summary>
        private static bool keepRunning = true;

        /// <summary>This class exists to keep linter quiet.</summary>
        internal static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("kernel32.dll")]
            internal static extern bool AllocConsole();
        }

        /// <summary>Entry point, if given a file name it will operate on it, otherwise it'll open command prompt and ask for commands.</summary>
        /// <param name="args">Argument representing the file name, extra arguments are ignored</param>
        [STAThread]
        internal static void Main(string[] args)
        {
            try
            {
                if (args.Length > 0)
                {
                    ProcessFile(args[0]);
                    return;
                }

                NativeMethods.AllocConsole();
                while (keepRunning) // loop over user input
                {
                    Console.Write("Type install, uninstall or exit and press enter> ");

                    switch (Console.ReadLine())
                    {
                        case "install": Install(); break;
                        case "uninstall": Uninstall(); break;
                        case "exit": return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }

        /// <summary>Given a filename it creates a text file in the same directory with .txt extension and opens it with default editor.</summary>
        /// <param name="filename">Name of the file to make a note for</param>
        private static void ProcessFile(string filename)
        {
            // replace extension
            string textFileName = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + ".txt");

            if (!File.Exists(textFileName))
            {
                try
                {
                    using (File.Create(textFileName))
                    {
                        // touch file and close handle
                    }
                }
                catch (Exception)
                {
                    // swallow exceptions from touching file
                }
            }

            Process.Start(textFileName); // Open program associated with .txts
        }

        /// <summary>Creates registry keys and copies executable.</summary>
        /// <param name="quiet">If true does not print to console.</param>
        private static void Install()
        {
            Uninstall(quiet: true); // quietly uninstall (but not quit)

            using (var rkMenu = HKCU.CreateSubKey(HKCU_Classes_Shell_MakeFileNote))
            {
                rkMenu.SetValue(string.Empty, Title);
                rkMenu.SetValue("Icon", TargetExecutable + ",0", Microsoft.Win32.RegistryValueKind.String);
            }

            using (var rkCommand = HKCU.CreateSubKey(HKCU_Classes_Shell_MakeFileNote_Command))
            {
                rkCommand.SetValue(string.Empty, "\"" + TargetExecutable + "\" \"%1\"");
            }

            File.Copy(Assembly.GetExecutingAssembly().Location, TargetExecutable, true);

            Console.WriteLine("Installed to " + TargetExecutable + ".");
        }

        /// <summary>Deletes registry keys and the executable</summary>
        /// <param name="quiet">If true does not print to console and does not quit.</param>
        private static void Uninstall(bool quiet = false)
        {
            HKCU.DeleteSubKey(HKCU_Classes_Shell_MakeFileNote_Command, throwOnMissingSubKey: false);
            HKCU.DeleteSubKey(HKCU_Classes_Shell_MakeFileNote, throwOnMissingSubKey: false);

            // Delete target executable after 3 seconds if it exists
            if (File.Exists(TargetExecutable))
            {
                var info = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    CreateNoWindow = true,
                    Arguments = "/C choice /C Y /N /D Y /T 2 & del \"" + TargetExecutable + "\"",
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                };

                Process.Start(info);
            }

            if (!quiet)
            {
                keepRunning = false;
                Console.WriteLine("Uninstalled.");
            }
        }
    }
}
