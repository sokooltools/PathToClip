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
	internal enum ContextMenu
	{
		FilePath = 0,
		FolderPath = 1,
		FileCmdPrompt = 2,
		FolderCmdPrompt = 3,
		FileVsCmdPrompt = 4,
		FolderVsCmdPrompt = 5
	}

	//----------------------------------------------------------------------------------------------------------------------------
	/// <summary>
	/// Manages the Windows Explorer right-click context menus by adding to the registry and
	/// creating an additional support file.
	/// </summary>
	//----------------------------------------------------------------------------------------------------------------------------
	internal static class ContextMenuManager
	{
		private static string _vs2019Batch;
		private static string _vsLatestBatch;

		private static readonly string CmdFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SokoolTools");
		private static readonly string CmdHereFilePath = Path.Combine(CmdFolderPath, "CmdHere.cmd");
		private static readonly string VsCmdHereFilePath = Path.Combine(CmdFolderPath, "VsCmdHere.cmd");

		private const string BATFILE150 = "vsdevcmd.bat";
		private const string FOLDER160 = "Microsoft Visual Studio 16.0";
		private const string FOLDER180 = "Microsoft Visual Studio 18.0";

		private const string VSWHERE_EXE = @"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe";

		private static class Hive
		{
			public const string FILE_PATH = @"*\shell\PathToClip";
			public const string FOLDER_PATH = @"Folder\shell\PathToClip";
			public const string FILE_CMD_PROMPT = @"*\shell\CmdHere";
			public const string FOLDER_CMD_PROMPT = @"Folder\shell\CmdHere";
			public const string FILE_VSCMD_PROMPT = @"*\shell\VsCmdHere";
			public const string FOLDER_VSCMD_PROMPT = @"Folder\shell\VsCmdHere";
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
		/// Gets the Visual Studio Batch File pertinent to the version (2019 or later).
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		public static string VisualStudioBatchFile => _vsLatestBatch ?? _vs2019Batch;

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
					sHive = Hive.FILE_PATH;
					break;
				case ContextMenu.FolderPath:
					sHive = Hive.FOLDER_PATH;
					break;
				case ContextMenu.FileCmdPrompt:
					sHive = Hive.FILE_CMD_PROMPT;
					break;
				case ContextMenu.FolderCmdPrompt:
					sHive = Hive.FOLDER_CMD_PROMPT;
					break;
				case ContextMenu.FileVsCmdPrompt:
					sHive = Hive.FILE_VSCMD_PROMPT;
					break;
				case ContextMenu.FolderVsCmdPrompt:
					sHive = Hive.FOLDER_VSCMD_PROMPT;
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
						using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FILE_PATH))
						{
							if (newKey != null)
							{
								newKey.SetValue(String.Empty, "Copy File Path to Clipboard");
								using (RegistryKey subkey = newKey.CreateSubKey("command"))
								{
									subkey?.SetValue(String.Empty, $"\"{AssemblyLocation}\" \"%1\"");
								}
							}
						}
					}
					break;
				case ContextMenu.FolderPath:
					if (!Exists(ContextMenu.FolderPath))
					{
						using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FOLDER_PATH))
						{
							if (newKey != null)
							{
								newKey.SetValue(String.Empty, "Copy Folder Path to Clipboard");
								using (RegistryKey subkey = newKey.CreateSubKey("command"))
								{
									subkey?.SetValue(String.Empty, $"\"{AssemblyLocation}\" \"%1\"");
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
					using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FILE_CMD_PROMPT))
					{
						if (newKey != null)
						{
							newKey.SetValue(String.Empty, @"Command Prompt Here");
							using (RegistryKey subkey = newKey.CreateSubKey("command"))
							{
								subkey?.SetValue(String.Empty, $"{CmdHereFilePath} \"%~dp1\" *");
							}
						}
					}
					break;
				case ContextMenu.FolderCmdPrompt:
					// Always remove the previously installed context menu.
					Remove(ContextMenu.FolderCmdPrompt);
					using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FOLDER_CMD_PROMPT))
					{
						if (newKey != null)
						{
							newKey.SetValue(String.Empty, @"Command Prompt Here");
							using (RegistryKey subkey = newKey.CreateSubKey("command"))
							{
								subkey?.SetValue(String.Empty, $"{CmdHereFilePath} \"%1\" f");
							}
						}
					}
					break;
				case ContextMenu.FileVsCmdPrompt:
					// Always remove the previously installed context menu.
					Remove(ContextMenu.FileVsCmdPrompt);
					// Create the command file.
					CreateVSCmdHereCmdFile();
					using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FILE_VSCMD_PROMPT))
					{
						if (newKey != null)
						{
							newKey.SetValue(String.Empty, @"Visual Studio Command Prompt Here");
							using (RegistryKey subkey = newKey.CreateSubKey("command"))
							{
								subkey?.SetValue(String.Empty, $"{VsCmdHereFilePath} \"%~dp1\" *");
							}
						}
					}
					break;
				case ContextMenu.FolderVsCmdPrompt:
					// Always remove the previously installed context menu.
					Remove(ContextMenu.FolderVsCmdPrompt);
					using (RegistryKey newKey = Registry.ClassesRoot.CreateSubKey(Hive.FOLDER_VSCMD_PROMPT))
					{
						if (newKey != null)
						{
							newKey.SetValue(String.Empty, @"Visual Studio Command Prompt Here");
							using (RegistryKey subkey = newKey.CreateSubKey("command"))
							{
								subkey?.SetValue(String.Empty, $"{VsCmdHereFilePath} \"%1\" f");
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
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FILE_PATH); // <-- make absolutely certain this subkey is correct!!!
					break;
				case ContextMenu.FolderPath:
					if (Exists(eContextMenu))
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FOLDER_PATH); // <-- make absolutely certain this subkey is correct!!!
					break;
				case ContextMenu.FileCmdPrompt:
					if (Exists(eContextMenu))
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FILE_CMD_PROMPT); // <-- make absolutely certain this subkey is correct!!!
					break;
				case ContextMenu.FolderCmdPrompt:
					if (Exists(eContextMenu))
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FOLDER_CMD_PROMPT); // <-- make absolutely certain this subkey is correct!!!
					break;
				case ContextMenu.FileVsCmdPrompt:
					if (Exists(eContextMenu))
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FILE_VSCMD_PROMPT); // <-- make absolutely certain this subkey is correct!!!
					break;
				case ContextMenu.FolderVsCmdPrompt:
					if (Exists(eContextMenu))
						Registry.ClassesRoot.DeleteSubKeyTree(Hive.FOLDER_VSCMD_PROMPT); // <-- make absolutely certain this subkey is correct!!!
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
		private static string AssemblyLocation => Path.Combine(Environment.CurrentDirectory, Assembly.GetExecutingAssembly().GetName().Name + ".exe");

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Finds the most recent Visual Studio batch file if one exists.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		private static void FindVisualStudioBatchFile()
		{
			string sFile;
			// Try latest VS via vswhere.
			string sPath = GetLatestVsInstallationPath();
			if (!string.IsNullOrEmpty(sPath))
			{
				sFile = Path.Combine(sPath, Path.Combine(@"Common7\Tools", BATFILE150));
				if (File.Exists(sFile))
				{
					_vsLatestBatch = sFile;
					return;
				}
			}
			// Fallback: explicit scans for VS 16.0 and 18.0 standard folders.
			foreach (var baseDir in new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) })
			{
				if (string.IsNullOrEmpty(baseDir)) continue;
				// VS 16.0
				string sPath16 = Path.Combine(baseDir, FOLDER160);
				sFile = Path.Combine(sPath16, Path.Combine(@"Common7\Tools", BATFILE150));
				if (File.Exists(sFile)) { _vs2019Batch = sFile; return; }
				// VS 18.0
				string sPath18 = Path.Combine(baseDir, FOLDER180);
				sFile = Path.Combine(sPath18, Path.Combine(@"Common7\Tools", BATFILE150));
				if (File.Exists(sFile)) { _vsLatestBatch = sFile; return; }
			}
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Get the latest Visual Studio 'Installation Path' using the 'Visual Studio Locator' application.
		/// </summary>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		private static string GetLatestVsInstallationPath()
		{
			string vsWhereFile = Environment.ExpandEnvironmentVariables(VSWHERE_EXE);
			if (!File.Exists(vsWhereFile))
				return null;
			using (var process = new System.Diagnostics.Process())
			{
				process.StartInfo = new System.Diagnostics.ProcessStartInfo
				{
					CreateNoWindow = true,
					FileName = vsWhereFile,
					Arguments = "-latest -products * -requires Microsoft.Component.MSBuild -property installationPath",
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
		/// Searches recursively for a Visual Studio Batch file beginning at the specified starting folder.
		/// </summary>
		/// <param name="sStartPath">Starting folder path</param>
		/// <param name="sFileToFind">Batch file to find</param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		private static string FindVsBatchFile(string sStartPath, string sFileToFind)
		{
			string[] file = Directory.GetFiles(sStartPath, sFileToFind, SearchOption.AllDirectories);
			return file.Length > 0 ? file[0] : null;
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Creates a command file used for opening a Visual Studio or standard command prompt when the user right-clicks a file 
		/// or a folder.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		private static void CreateCmdHereCmdFile()
		{
			if (!Directory.Exists(CmdFolderPath))
				Directory.CreateDirectory(CmdFolderPath);
			using (StreamWriter sw = File.CreateText(CmdHereFilePath))
			{
				sw.WriteLine(":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
				sw.WriteLine(":: This command file is used to open a standard windows command prompt in the  ");
				sw.WriteLine(":: containing folder of a right-clicked file or folder within Windows Explorer.");
				sw.WriteLine(":: Concept conceived and created by SokoolTools, (c) 2007-2019.                ");
				sw.WriteLine(":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
				sw.WriteLine("@echo off");
				sw.WriteLine("mode con lines=1 cols=20");
				sw.WriteLine("if '%2'=='*' (chdir %~dp1) else (chdir %1)");
				sw.WriteLine("start %comspec% /k \"TITLE Command Prompt\"");
				sw.WriteLine("exit /B 0");
			}
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Creates a command file used for opening a Visual Studio or standard command prompt when the user right-clicks a file 
		/// or a folder.
		/// </summary>
		//------------------------------------------------------------------------------------------------------------------------
		private static void CreateVSCmdHereCmdFile()
		{
			if (!Directory.Exists(CmdFolderPath))
				Directory.CreateDirectory(CmdFolderPath);
			using (StreamWriter sw = File.CreateText(VsCmdHereFilePath))
			{
				sw.WriteLine(":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
				sw.WriteLine(":: This command file is used to open a Visual Studio command prompt in the     ");
				sw.WriteLine(":: containing folder of a right-clicked file or folder within Windows Explorer.");
				sw.WriteLine(":: Concept conceived and created by SokoolTools, (c) 2007-2019.                ");
				sw.WriteLine(":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::");
				sw.WriteLine("@echo off");
				sw.WriteLine("mode con lines=1 cols=20");
				sw.WriteLine("if '%2'=='*' (chdir %~dp1) else (chdir %1)");
				sw.WriteLine("start %comspec% /k \"set VSCMD_START_DIR=%CD% && TITLE Visual Studio Command Prompt && \"{0}\"\"", VisualStudioBatchFile);
				sw.WriteLine("exit /B 0");
			}
		}
	}
}