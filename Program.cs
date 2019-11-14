using System;
using System.Windows.Forms;

namespace PathToClip
{
    internal static class Program
    {
        //----------------------------------------------------------------------------------------------------
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        //----------------------------------------------------------------------------------------------------
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Clipboard.Clear();
                Clipboard.SetDataObject(args[0].Trim('\"'), true, 3, 150);
                Application.Exit();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new AboutBox());
            }
        }
    }
}