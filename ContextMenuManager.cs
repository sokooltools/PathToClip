using System;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace PathToClip
{
    //----------------------------------------------------------------------------------------------------------------------------
    /// <summary>
    /// Enumerated types of context menus.
    /// </summary>
    //----------------------------------------------------------------------------------------------------------------------------
    public enum ContextMenu
    {
        FilePath = 0,
        FolderPath = 1,
        FileCmdPrompt = 2,
        FolderCmdPrompt = 3
    }

    //----------------------------------------------------------------------------------------------------------------------------
    /// <summary>
    /// Manages the Windows Explorer right-click context menus by adding to the registry and
    /// creating an additional support file.
    /// </summary>
    //----------------------------------------------------------------------------------------------------------------------------
    public static class ContextMenuManager
    {
        private static string _vs2019Batch;
        private static string _vs2017Batch;
        private static string _vs2015Batch;
        private static string _vs2013Batch;
        private static string _vs2010Batch;
        private static string _vs2008Batch;
        
        private const string FOLDER90 = "Microsoft Visual Studio 9.0";
        private const string BATFILE90 = "vcvarsall.bat";
        private const string FOLDER100 = "Microsoft Visual Studio 10.0";
        private const string BATFILE100 = "vcvarsall.bat";
        private const string FOLDER120 = "Microsoft Visual Studio 12.0";
        private const string BATFILE120 = "vcvarsall.bat";
        private const string FOLDER140 = "Microsoft Visual Studio 14.0";
        private const string BATFILE140 = "vsdevcmd.bat";
        private const string BATFILE150 = "vsdevcmd.bat";

        private const string VSWHERE_EXE = @"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe";

        private static class Hive
        {
            public const string FilePath = @"*\shell\PathToClip";
            public const string FolderPath = @"Folder\shell\PathToClip";
            public const string FileCmdPrompt = @"*\shell\CmdHere";
            public const string FolderCmdPrompt = @"Folder\shell\CmdHere";
        }

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Initializes the <see cref="ContextMenuManager"/> class.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		static ContextMenuManager()
        {
            FindVisualStudioBatchFile();
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Gets the Command Prompt text.
        /// </summary>
        //------------------------------------------------------------------------------------------------------------------------
        internal static string CommandPromptText
        {
            get
            {
                if (_vs2019Batch != null)
                    return "Visual Studio 2019 Command &Prompt";
                if (_vs2017Batch != null)
                    return "Visual Studio 2017 Command &Prompt";
                if (_vs2015Batch != null)
                    return "Visual Studio 2015 Command &Prompt";
                if (_vs2013Batch != null)
                    return "Visual Studio 2013 Command &Prompt";
                if (_vs2010Batch != null)
                    return "Visual Studio 2010 Command &Prompt";
                return _vs2008Batch != null ? "Visual Studio 2008 Command &Prompt" : "Command &Prompt";
            }
        }

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Gets the Visual Studio Batch File pertinent to the version.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		private static string VisualStudioBatchFile => _vs2019Batch ?? _vs2017Batch ?? _vs2015Batch ?? _vs2013Batch ?? _vs2010Batch ?? _vs2008Batch;

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Indicates whether the specified right-click context menu currently exists in the registry.
        /// </summary>
        /// <param name="eContextMenu"></param>
        /// <returns>True if the context menu exists in the registry; false otherwise.</returns>
        //------------------------------------------------------------------------------------------------------------------------
        internal static bool Exists(ContextMenu eContextMenu)
        {
            string sHive;

            switch (eContextMenu)
            {
                case ContextMenu.FilePath:
                    sHive = Hive.FilePath;
                    break;
                case ContextMenu.FolderPath:
                    sHive = Hive.FolderPath;
                    break;
                case ContextMenu.FileCmdPrompt:
                    sHive = Hive.FileCmdPrompt;
                    break;
                case ContextMenu.FolderCmdPrompt:
                    sHive = Hive.FolderCmdPrompt;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(eContextMenu), eContextMenu, null);
            }
            using (RegistryKey tstKey = Registry.ClassesRoot.OpenSubKey(sHive, false))
            {
                return tstKey != null;
            }
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Adds the specified right-click context menu to Windows Explorer.
        /// </summary>
        /// <param name="eContextMenu"></param>
        //------------------------------------------------------------------------------------------------------------------------
        internal static void Add(ContextMenu eContextMenu)
        {
            switch (eContextMenu)
            {
                case ContextMenu.FilePath:
                    if (!Exists(ContextMenu.FilePath))
                    {
                        using (RegistryKey newkey = Registry.ClassesRoot.CreateSubKey(Hive.FilePath))
                        {
                            if (newkey != null)
                            {
                                newkey.SetValue("", "Copy FilePath to Clipboard");
                                using (RegistryKey subkey = newkey.CreateSubKey("command"))
                                {
                                    subkey?.SetValue("", $"\"{AssemblyLocation()}\" \"%1\"");
                                }
                            }
                        }
                    }
                    break;
                case ContextMenu.FolderPath:
                    if (!Exists(ContextMenu.FolderPath))
                    {
                        using (RegistryKey newkey = Registry.ClassesRoot.CreateSubKey(Hive.FolderPath))
                        {
                            if (newkey != null)
                            {
                                newkey.SetValue("", "Copy FolderPath to Clipboard");
                                using (RegistryKey subkey = newkey.CreateSubKey("command"))
                                {
                                    subkey?.SetValue("", $"\"{AssemblyLocation()}\" \"%1\"");
                                }
                            }
                        }
                    }
                    break;
                case ContextMenu.FileCmdPrompt:
                    // Always remove the previously installed context menu.
                    Remove(ContextMenu.FileCmdPrompt);
                    // Create the command file.
                    CreateCmdHereCmdFile();
                    using (RegistryKey newkey = Registry.ClassesRoot.CreateSubKey(Hive.FileCmdPrompt))
                    {
                        if (newkey != null)
                        {
                            newkey.SetValue("", CommandPromptText + @" Here");
                            using (RegistryKey subkey = newkey.CreateSubKey("command"))
                            {
                                subkey?.SetValue("", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DevTools", "CmdHere.cmd") + " \"%~dp1\" *");
                            }
                        }
                    }
                    break;
                case ContextMenu.FolderCmdPrompt:
                    // Always remove the previously installed context menu.
                    Remove(ContextMenu.FolderCmdPrompt);
                    using (RegistryKey newkey = Registry.ClassesRoot.CreateSubKey(Hive.FolderCmdPrompt))
                    {
                        if (newkey != null)
                        {
                            newkey.SetValue("", CommandPromptText + @" Here");
                            using (RegistryKey subkey = newkey.CreateSubKey("command"))
                            {
                                subkey?.SetValue("", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DevTools", "CmdHere.cmd") + " \"%1\" f");
                            }
                        }
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(eContextMenu), eContextMenu, null);
            }
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Removes the specified right-click context menu from Windows Explorer.
        /// </summary>
        /// <param name="eContextMenu"></param>
        //------------------------------------------------------------------------------------------------------------------------
        internal static void Remove(ContextMenu eContextMenu)
        {
            switch (eContextMenu)
            {
                case ContextMenu.FilePath:
                    if (Exists(eContextMenu))
                        Registry.ClassesRoot.DeleteSubKeyTree(Hive.FilePath); // <-- make absolutely certain this subkey is correct!!!
                    break;
                case ContextMenu.FolderPath:
                    if (Exists(eContextMenu))
                        Registry.ClassesRoot.DeleteSubKeyTree(Hive.FolderPath); // <-- make absolutely certain this subkey is correct!!!
                    break;
                case ContextMenu.FileCmdPrompt:
                    if (Exists(eContextMenu))
                        Registry.ClassesRoot.DeleteSubKeyTree(Hive.FileCmdPrompt); // <-- make absolutely certain this subkey is correct!!!
                    break;
                case ContextMenu.FolderCmdPrompt:
                    if (Exists(eContextMenu))
                        Registry.ClassesRoot.DeleteSubKeyTree(Hive.FolderCmdPrompt); // <-- make absolutely certain this subkey is correct!!!
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(eContextMenu), eContextMenu, null);
            }
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// AssemblyLocation
        /// </summary>
        /// <returns></returns>
        //------------------------------------------------------------------------------------------------------------------------
        private static string AssemblyLocation()
        {
            return Path.Combine(Environment.CurrentDirectory, Assembly.GetExecutingAssembly().GetName().Name + ".exe");
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Finds the most recent Visual Studio batch file if one exists.
        /// </summary>
        //------------------------------------------------------------------------------------------------------------------------
        private static void FindVisualStudioBatchFile()
        {
            int loc;
            string sFile;
            // Look for VS2017 or VS2019.
            string sPath = Get2017InstallationPath();
            if (!string.IsNullOrEmpty(sPath))
            {
                sFile = Path.Combine(sPath, Path.Combine(@"Common7\Tools", BATFILE150));
                if (File.Exists(sFile))
                {
                    string text = File.ReadAllText(sFile);
                    if (text.Contains("VisualStudioVersion=16.0"))
                        _vs2019Batch = sFile;
                    else
                        _vs2017Batch = sFile;
                    return;
                }
            }
            // Look for VS2015.
            sPath = Environment.GetEnvironmentVariable("VS140COMNTOOLS");
            if (sPath != null)
            {
                loc = sPath.IndexOf(FOLDER140, StringComparison.OrdinalIgnoreCase);
                if (loc > -1)
                    sPath = sPath.Substring(0, loc + FOLDER140.Length);
                sFile = Path.Combine(sPath, Path.Combine(@"Common7\Tools", BATFILE140));
                if (File.Exists(sFile))
                {
                    _vs2015Batch = sFile;
                    return;
                }
                sFile = FindFile(sPath, BATFILE140);
                if (sFile != null)
                {
                    _vs2015Batch = sFile;
                    return;
                }
            }
            // Look for VS2013.
            sPath = Environment.GetEnvironmentVariable("VS120COMNTOOLS");
            if (sPath != null)
            {
                loc = sPath.IndexOf(FOLDER120, StringComparison.OrdinalIgnoreCase);
                if (loc > -1)
                    sPath = sPath.Substring(0, loc + FOLDER120.Length);
                sFile = Path.Combine(sPath, Path.Combine("vc", BATFILE120));
                if (File.Exists(sFile))
                {
                    _vs2013Batch = sFile;
                    return;
                }
                sFile = FindFile(sPath, BATFILE120);
                if (sFile != null)
                {
                    _vs2013Batch = sFile;
                    return;
                }
            }
            // Look for VS2010
            sPath = Environment.GetEnvironmentVariable("VS100COMNTOOLS");
            if (sPath != null)
            {
                loc = sPath.IndexOf(FOLDER100, StringComparison.OrdinalIgnoreCase);
                if (loc > -1)
                    sPath = sPath.Substring(0, loc + FOLDER100.Length);
                sFile = Path.Combine(sPath, Path.Combine("vc", BATFILE100));
                if (File.Exists(sFile))
                {
                    _vs2010Batch = sFile;
                    return;
                }
                sFile = FindFile(sPath, BATFILE100);
                if (sFile != null)
                {
                    _vs2010Batch = sFile;
                    return;
                }
            }
            // Look for VS2008
            sPath = Environment.GetEnvironmentVariable("VS90COMNTOOLS");
            if (sPath == null)
                return;
            loc = sPath.IndexOf(FOLDER90, StringComparison.OrdinalIgnoreCase);
            if (loc > -1)
                sPath = sPath.Substring(0, loc + FOLDER90.Length);
            sFile = Path.Combine(sPath, Path.Combine("vc", BATFILE90));
            if (File.Exists(sFile))
            {
                _vs2008Batch = sFile;
                return;
            }
            sFile = FindFile(sPath, BATFILE90);
            if (sFile != null)
                _vs2008Batch = sFile;
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Get the Visual Studio 2017 or 2019 'Installation Path' using the 'Visual Studio Locator' application.
        /// </summary>
        /// <returns></returns>
        //------------------------------------------------------------------------------------------------------------------------
        private static string Get2017InstallationPath()
        {
            using (var process = new System.Diagnostics.Process())
            {
                process.StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    CreateNoWindow = true,
                    FileName = Environment.ExpandEnvironmentVariables(VSWHERE_EXE),
                    Arguments = "-prerelease -latest -property installationPath",
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                };
                process.Start();
                string output = process.StandardOutput.ReadToEnd().Trim('\r', '\n');
                process.WaitForExit();
                return output;
            }
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Searches for a file recursively beginning with a starting folder.
        /// </summary>
        /// <param name="sStartPath"></param>
        /// <param name="sFileToFind"></param>
        /// <returns></returns>
        //------------------------------------------------------------------------------------------------------------------------
        private static string FindFile(string sStartPath, string sFileToFind)
        {
            string[] file = Directory.GetFiles(sStartPath, sFileToFind, SearchOption.AllDirectories);
            return file.Length > 0 ? file[0] : null;
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Creates a command file used for opening a Visual Studio or standard command prompt when
        /// the user right-clicks a file or a folder.
        /// </summary>
        //------------------------------------------------------------------------------------------------------------------------
        private static void CreateCmdHereCmdFile()
        {
            
            using (StreamWriter sw = File.CreateText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"DevTools", "CmdHere.cmd")))
            {
                sw.WriteLine("::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
                sw.WriteLine(":: This command file is used to open a command prompt in the containing   ");
                sw.WriteLine(":: folder of a right-clicked file or folder within Windows Explorer.      ");
                sw.WriteLine(":: Concept conceived and created by Ronn A. Sokol, (c) 2007-2019.         ");
                sw.WriteLine("::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
                sw.WriteLine("@echo off");
                sw.WriteLine("@title VSCMD");
                sw.WriteLine("mode con lines=1 cols=20");
                sw.WriteLine("if '%2'=='*' (chdir %~dp1) else (chdir %1)");
                if (VisualStudioBatchFile == null)
                    sw.WriteLine("start %comspec% /k");
                else
                    sw.WriteLine("start %comspec% /k \"set VSCMD_START_DIR=%CD% && \"{0}\"\"", VisualStudioBatchFile);
                sw.WriteLine("exit /B 0");
            }
        }

        //public static string ProgramFilesx86()
        //{
        //    return IntPtr.Size == 8 || !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))
        //    ? Environment.GetEnvironmentVariable("ProgramFiles(x86)")
        //    : Environment.GetEnvironmentVariable("ProgramFiles");
        //}
    }
}