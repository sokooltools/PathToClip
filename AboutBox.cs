using System;
using System.IO;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;

namespace PathToClip
{
    internal partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();

            //  Initialize the AboutBox to display the product information from the assembly information.
            //  Change assembly information settings for your application through either:
            //  - Project->Properties->Application->Assembly Information
            //  - AssemblyInfo.cs
        }

        private void AboutBox_Load(object sender, EventArgs e)
        {
            Text = $@"About {AssemblyTitle}";
            labelProductName.Text = AssemblyProduct;
            labelVersion.Text = $@"Version {AssemblyVersion}";
            labelCopyright.Text = AssemblyCopyright;
            labelCompanyName.Text = AssemblyCompany;
            textBoxDescription.Text = AssemblyDescription;
        }

        private void AboutBox_Shown(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
				chkFolderPathContextMenu.Checked = ContextMenuManager.Exists(PathToClip.ContextMenu.FolderPath);
				chkFilePathContextMenu.Checked = ContextMenuManager.Exists(PathToClip.ContextMenu.FilePath);
                chkCmdPromptContextMenu.Checked = ContextMenuManager.Exists(PathToClip.ContextMenu.FolderCmdPrompt);
				chkVsCmdPromptContextMenu.Checked = ContextMenuManager.Exists(PathToClip.ContextMenu.FolderVsCmdPrompt);
                chkVsCmdPromptContextMenu.Visible = ContextMenuManager.VisualStudioBatchFile != null;
            }
			catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                if (!IsWindowsAdministrator())
                    throw new ApplicationException("Sorry, 'Windows Administrative' rights are required to perform this function...");

                if (chkFolderPathContextMenu.Checked)
					ContextMenuManager.Add(PathToClip.ContextMenu.FolderPath);
                else
					ContextMenuManager.Remove(PathToClip.ContextMenu.FolderPath);

                if (chkFilePathContextMenu.Checked)
					ContextMenuManager.Add(PathToClip.ContextMenu.FilePath);
                else
					ContextMenuManager.Remove(PathToClip.ContextMenu.FilePath);

                if (chkCmdPromptContextMenu.Checked)
                {
                    ContextMenuManager.Add(PathToClip.ContextMenu.FileCmdPrompt);
                    ContextMenuManager.Add(PathToClip.ContextMenu.FolderCmdPrompt);
                }
                else
                {
                    ContextMenuManager.Remove(PathToClip.ContextMenu.FileCmdPrompt);
                    ContextMenuManager.Remove(PathToClip.ContextMenu.FolderCmdPrompt);
                }

				if (chkVsCmdPromptContextMenu.Visible && chkVsCmdPromptContextMenu.Checked)
                {
                    ContextMenuManager.Add(PathToClip.ContextMenu.FileVsCmdPrompt);
                    ContextMenuManager.Add(PathToClip.ContextMenu.FolderVsCmdPrompt);
                }
                else
                {
                    ContextMenuManager.Remove(PathToClip.ContextMenu.FileVsCmdPrompt);
                    ContextMenuManager.Remove(PathToClip.ContextMenu.FolderVsCmdPrompt);
                }
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        #region Assembly Attribute Accessors

        private static string AssemblyTitle
        {
            get
            {
                // Get all Title attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                // If there is at least one Title attribute
                if (attributes.Length <= 0)
                    return Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
                // Select the first one
                var titleAttribute = (AssemblyTitleAttribute)attributes[0];
                // If it is not an empty string, return it
                return
                    titleAttribute.Title != string.Empty
                    ? titleAttribute.Title
                    : Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
                // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
            }
        }

        private static string AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        private static string AssemblyDescription
        {
            get
            {
                // Get all Description attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute),
                                                                                                                  false);
                // If there aren't any Description attributes, return an empty string, otherwise return its value.
                return attributes.Length == 0 ? string.Empty : ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        private static string AssemblyProduct
        {
            get
            {
                // Get all Product attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                // If there aren't any Product attributes, return an empty string
                return attributes.Length == 0 ? string.Empty : ((AssemblyProductAttribute)attributes[0]).Product;
                // If there is a Product attribute, return its value
            }
        }

        private static string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                // If there aren't any Copyright attributes, return an empty string
                return attributes.Length == 0 ? string.Empty : ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
                // If there is a Copyright attribute, return its value
            }
        }

        private static string AssemblyCompany
        {
            get
            {
                // Get all Company attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                // If there aren't any Company attributes, return an empty string
                return attributes.Length == 0 ? string.Empty : ((AssemblyCompanyAttribute)attributes[0]).Company;
                // If there is a Company attribute, return its value
            }
        }

        //------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Returns an indication as to whether the current user belongs to the Windows user group with an administrator role.
        /// </summary>
        /// <returns><c>true</c> if the current user is a Windows administrator; otherwise, <c>false</c>.</returns>
        //------------------------------------------------------------------------------------------------------------------------
        private static bool IsWindowsAdministrator()
        {
            return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
        }

        #endregion
    }
}