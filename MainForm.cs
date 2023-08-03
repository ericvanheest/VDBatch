using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace VDBatch
{
	/// <summary>
	/// Summary description for VDBAForm.
	/// </summary>
	public class VDBAForm : System.Windows.Forms.Form
	{
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.StatusBar statusBar1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem mi_File;
        private System.Windows.Forms.MenuItem mi_File_Exit;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Splitter splitter2;
        private System.Windows.Forms.OpenFileDialog ofdBrowse;
        private System.Windows.Forms.ListView lvFiles;
        private System.Windows.Forms.ColumnHeader chFilename;
        private System.Windows.Forms.ContextMenu cmEdit;
        private System.Windows.Forms.MenuItem cm_Delete;
        private System.Windows.Forms.MenuItem mi_Edit;
        private System.Windows.Forms.MenuItem mi_Edit_CreateJobs;
        private System.Windows.Forms.SaveFileDialog sfdJobs;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.MenuItem mi_File_Open;
        private System.Windows.Forms.MenuItem mi_File_Save;
        private System.Windows.Forms.SaveFileDialog sfdTemplate;
        private System.Windows.Forms.OpenFileDialog ofdTemplate;
        private System.Windows.Forms.TextBox tbTemplate;
        private System.Windows.Forms.Button btnCreateJobsFile;
        private System.Windows.Forms.TextBox tbInstructions;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Splitter splitter3;
        private IContainer components;
        private System.Windows.Forms.MenuItem mi_Edit_ClearFileList;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem mi_Edit_Options;

        private const string sJobSeparator = "// $job";

        private Options m_optionsForm = new Options();
        private CLOptions m_options;

		public VDBAForm(CLOptions options)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
            CreateInstructions();

            m_options = options;
            lvFiles.Columns[0].Width = -2;
		}

        public void CreateInstructions()
        {
            m_optionsForm.JobPrefix = sPrefix;
            m_optionsForm.JobSuffix = sSuffix;
            m_optionsForm.ReadFromRegistry();
            sPrefix = m_optionsForm.JobPrefix;
            sSuffix = m_optionsForm.JobSuffix;

            tbInstructions.Text = Resource1.Instructions;

            tbTemplate.Text = Resource1.Prefix + Resource1.TemplateBase + Resource1.Suffix;
        }

        public string sPrefix = Resource1.Prefix;
        public string sSuffix = Resource1.Suffix;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lvFiles = new System.Windows.Forms.ListView();
            this.chFilename = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.cmEdit = new System.Windows.Forms.ContextMenu();
            this.cm_Delete = new System.Windows.Forms.MenuItem();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.mi_File = new System.Windows.Forms.MenuItem();
            this.mi_File_Open = new System.Windows.Forms.MenuItem();
            this.mi_File_Save = new System.Windows.Forms.MenuItem();
            this.mi_File_Exit = new System.Windows.Forms.MenuItem();
            this.mi_Edit = new System.Windows.Forms.MenuItem();
            this.mi_Edit_CreateJobs = new System.Windows.Forms.MenuItem();
            this.mi_Edit_ClearFileList = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.mi_Edit_Options = new System.Windows.Forms.MenuItem();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tbTemplate = new System.Windows.Forms.TextBox();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbInstructions = new System.Windows.Forms.TextBox();
            this.splitter3 = new System.Windows.Forms.Splitter();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnCreateJobsFile = new System.Windows.Forms.Button();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.ofdBrowse = new System.Windows.Forms.OpenFileDialog();
            this.sfdJobs = new System.Windows.Forms.SaveFileDialog();
            this.sfdTemplate = new System.Windows.Forms.SaveFileDialog();
            this.ofdTemplate = new System.Windows.Forms.OpenFileDialog();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(0, 0);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 1;
            this.btnBrowse.Text = "&Browse";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lvFiles
            // 
            this.lvFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chFilename});
            this.lvFiles.ContextMenu = this.cmEdit;
            this.lvFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvFiles.FullRowSelect = true;
            this.lvFiles.LabelEdit = true;
            this.lvFiles.Location = new System.Drawing.Point(0, 0);
            this.lvFiles.Name = "lvFiles";
            this.lvFiles.Size = new System.Drawing.Size(656, 152);
            this.lvFiles.TabIndex = 2;
            this.lvFiles.UseCompatibleStateImageBehavior = false;
            this.lvFiles.View = System.Windows.Forms.View.Details;
            this.lvFiles.Resize += new System.EventHandler(this.lvFiles_Resize);
            // 
            // chFilename
            // 
            this.chFilename.Text = "Filename";
            // 
            // cmEdit
            // 
            this.cmEdit.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.cm_Delete});
            // 
            // cm_Delete
            // 
            this.cm_Delete.Index = 0;
            this.cm_Delete.Shortcut = System.Windows.Forms.Shortcut.Del;
            this.cm_Delete.Text = "&Delete";
            this.cm_Delete.Click += new System.EventHandler(this.cm_Delete_Click);
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 555);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(656, 22);
            this.statusBar1.TabIndex = 4;
            this.statusBar1.Text = "Ready.";
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mi_File,
            this.mi_Edit});
            // 
            // mi_File
            // 
            this.mi_File.Index = 0;
            this.mi_File.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mi_File_Open,
            this.mi_File_Save,
            this.mi_File_Exit});
            this.mi_File.Text = "&File";
            // 
            // mi_File_Open
            // 
            this.mi_File_Open.Index = 0;
            this.mi_File_Open.Shortcut = System.Windows.Forms.Shortcut.CtrlO;
            this.mi_File_Open.Text = "&Open job file template...";
            this.mi_File_Open.Click += new System.EventHandler(this.mi_File_Open_Click);
            // 
            // mi_File_Save
            // 
            this.mi_File_Save.Index = 1;
            this.mi_File_Save.Shortcut = System.Windows.Forms.Shortcut.CtrlS;
            this.mi_File_Save.Text = "&Save job file template...";
            this.mi_File_Save.Click += new System.EventHandler(this.mi_File_Save_Click);
            // 
            // mi_File_Exit
            // 
            this.mi_File_Exit.Index = 2;
            this.mi_File_Exit.Text = "E&xit";
            this.mi_File_Exit.Click += new System.EventHandler(this.mi_File_Exit_Click);
            // 
            // mi_Edit
            // 
            this.mi_Edit.Index = 1;
            this.mi_Edit.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mi_Edit_CreateJobs,
            this.mi_Edit_ClearFileList,
            this.menuItem1,
            this.mi_Edit_Options});
            this.mi_Edit.Text = "&Edit";
            this.mi_Edit.Popup += new System.EventHandler(this.mi_Edit_Popup);
            // 
            // mi_Edit_CreateJobs
            // 
            this.mi_Edit_CreateJobs.Index = 0;
            this.mi_Edit_CreateJobs.Shortcut = System.Windows.Forms.Shortcut.CtrlJ;
            this.mi_Edit_CreateJobs.Text = "&Create .jobs file...";
            this.mi_Edit_CreateJobs.Click += new System.EventHandler(this.mi_Edit_CreateJobs_Click);
            // 
            // mi_Edit_ClearFileList
            // 
            this.mi_Edit_ClearFileList.Index = 1;
            this.mi_Edit_ClearFileList.Text = "C&lear file list";
            this.mi_Edit_ClearFileList.Click += new System.EventHandler(this.mi_Edit_ClearFileList_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 2;
            this.menuItem1.Text = "-";
            // 
            // mi_Edit_Options
            // 
            this.mi_Edit_Options.Index = 3;
            this.mi_Edit_Options.Text = "&Options...";
            this.mi_Edit_Options.Click += new System.EventHandler(this.mi_Edit_Options_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tbTemplate);
            this.panel2.Controls.Add(this.splitter1);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 155);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(656, 400);
            this.panel2.TabIndex = 6;
            // 
            // tbTemplate
            // 
            this.tbTemplate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbTemplate.Location = new System.Drawing.Point(347, 0);
            this.tbTemplate.MaxLength = 65535;
            this.tbTemplate.Multiline = true;
            this.tbTemplate.Name = "tbTemplate";
            this.tbTemplate.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbTemplate.Size = new System.Drawing.Size(309, 400);
            this.tbTemplate.TabIndex = 4;
            this.tbTemplate.WordWrap = false;
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(344, 0);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 400);
            this.splitter1.TabIndex = 3;
            this.splitter1.TabStop = false;
            this.splitter1.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tbInstructions);
            this.panel1.Controls.Add(this.splitter3);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(344, 400);
            this.panel1.TabIndex = 2;
            // 
            // tbInstructions
            // 
            this.tbInstructions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbInstructions.Location = new System.Drawing.Point(0, 0);
            this.tbInstructions.Multiline = true;
            this.tbInstructions.Name = "tbInstructions";
            this.tbInstructions.ReadOnly = true;
            this.tbInstructions.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbInstructions.Size = new System.Drawing.Size(344, 373);
            this.tbInstructions.TabIndex = 6;
            this.tbInstructions.Text = "Instructions:";
            // 
            // splitter3
            // 
            this.splitter3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.splitter3.Location = new System.Drawing.Point(0, 373);
            this.splitter3.Name = "splitter3";
            this.splitter3.Size = new System.Drawing.Size(344, 3);
            this.splitter3.TabIndex = 8;
            this.splitter3.TabStop = false;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.btnBrowse);
            this.panel3.Controls.Add(this.btnCreateJobsFile);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 376);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(344, 24);
            this.panel3.TabIndex = 7;
            // 
            // btnCreateJobsFile
            // 
            this.btnCreateJobsFile.Location = new System.Drawing.Point(88, 0);
            this.btnCreateJobsFile.Name = "btnCreateJobsFile";
            this.btnCreateJobsFile.Size = new System.Drawing.Size(96, 23);
            this.btnCreateJobsFile.TabIndex = 5;
            this.btnCreateJobsFile.Text = "&Create Job File";
            this.btnCreateJobsFile.Click += new System.EventHandler(this.btnCreateJobsFile_Click);
            // 
            // splitter2
            // 
            this.splitter2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.splitter2.Location = new System.Drawing.Point(0, 152);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(656, 3);
            this.splitter2.TabIndex = 7;
            this.splitter2.TabStop = false;
            // 
            // ofdBrowse
            // 
            this.ofdBrowse.Filter = "AVI files (*.avi)|*.avi|All files (*.*)|*.*";
            this.ofdBrowse.Multiselect = true;
            this.ofdBrowse.Title = "Select the video files to include in the batch operation";
            // 
            // sfdJobs
            // 
            this.sfdJobs.DefaultExt = "jobs";
            this.sfdJobs.FileName = "VirtualDub.jobs";
            this.sfdJobs.Filter = "VirtualDub Job file (*.jobs)|*.jobs|All files (*.*)|*.*";
            this.sfdJobs.Title = "Select the location for the .jobs file";
            // 
            // sfdTemplate
            // 
            this.sfdTemplate.DefaultExt = "job";
            this.sfdTemplate.FileName = "Template.job";
            this.sfdTemplate.Filter = "Job templates (*.job)|*.job|All files (*.*)|*.*";
            this.sfdTemplate.Title = "Select a filename for this job template";
            // 
            // ofdTemplate
            // 
            this.ofdTemplate.FileName = "Template.job";
            this.ofdTemplate.Filter = "Job templates (*.job)|*.job|All files (*.*)|*.*";
            this.ofdTemplate.Title = "Select a job template to open";
            // 
            // VDBAForm
            // 
            this.AllowDrop = true;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(656, 577);
            this.Controls.Add(this.lvFiles);
            this.Controls.Add(this.splitter2);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.statusBar1);
            this.Menu = this.mainMenu1;
            this.Name = "VDBAForm";
            this.Text = "VirtualDub Batch Assistant v1.0.4";
            this.Load += new System.EventHandler(this.VDBAForm_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.VDBAForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.VDBAForm_DragEnter);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.VDBAForm_KeyDown);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
            CLOptions options = new CLOptions(Environment.GetCommandLineArgs(), true);
			Application.Run(new VDBAForm(options));
		}

        private void mi_File_Exit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void SetReadyStatus()
        {
            SetStatus("Ready.");
        }

        private void SetStatus(string sText)
        {
            statusBar1.Text = sText;
        }

        private void SetStatusWarning(string sText)
        {
            SetStatus("Warning: " + sText);
        }

        private void SetStatusError(string sText)
        {
            SetStatus("Error: " + sText);
        }

        private void AddFileToList(string sFile)
        {
            SetReadyStatus();

            if (!File.Exists(sFile))
            {
                SetStatusWarning("File does not exist - " + sFile);
                return;
            }

            foreach(ListViewItem lvItem in lvFiles.Items)
            {
                if (lvItem.SubItems[0].Text == sFile)
                    return;
            }

            lvFiles.Items.Add(sFile);
        }

        private void btnBrowse_Click(object sender, System.EventArgs e)
        {
            if (ofdBrowse.ShowDialog() == DialogResult.OK)
            {
                foreach (string sFile in ofdBrowse.FileNames)
                {
                    AddFileToList(sFile);
                }
            }
        }

        private void lvFiles_Resize(object sender, System.EventArgs e)
        {
            lvFiles.Columns[0].Width = -2;
        }

        private void AddDirectory(string sDir)
        {
            foreach (string sFile in Directory.GetFiles(sDir, "*.avi"))
            {
                AddFileToList(sFile);
            }

            foreach (string sSubdir in Directory.GetDirectories(sDir))
            {
                AddDirectory(sSubdir);
            }
        }

        private void VDBAForm_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            string[] saFiles = (string[]) e.Data.GetData(DataFormats.FileDrop);
            if (saFiles.Length == 1) 
            {
                string sTest = Path.GetExtension(saFiles[0]).ToLower();
                if (sTest == ".job")
                {
                    ReadJobTemplate(saFiles[0]);
                    return;
                }
                else if (sTest == ".vcf")
                {
                    InsertVCF(saFiles[0]);
                    return;
                }
            }

            lvFiles.SuspendLayout();
            foreach (string sFile in saFiles)
            {
                if (Directory.Exists(sFile))
                    AddDirectory(sFile);
                else
                    AddFileToList(sFile);
            }
            lvFiles.ResumeLayout();
        }

        private void VDBAForm_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            if(e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void VDBAForm_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
        
        }

        private void DeleteSelectedFiles()
        {
            lvFiles.SuspendLayout();
            foreach(ListViewItem lvItem in lvFiles.SelectedItems)
                lvItem.Remove();
            lvFiles.ResumeLayout();
        }

        private void cm_Delete_Click(object sender, System.EventArgs e)
        {
            DeleteSelectedFiles();
        }

        private void mi_Edit_CreateJobs_Click(object sender, System.EventArgs e)
        {
            CreateJobFile();
        }

        private void CreateJobFile()
        {
            sfdJobs.FileName = m_optionsForm.JobsPath;

            if (sfdJobs.ShowDialog() == DialogResult.OK)
            {
                CreateJobsFile(sfdJobs.FileName);
            }
        }

        private void mi_File_Open_Click(object sender, System.EventArgs e)
        {
            OpenTemplateFile();
        }

        private void OpenTemplateFile()
        {
            if (ofdTemplate.ShowDialog() == DialogResult.OK)
            {
                ReadJobTemplate(ofdTemplate.FileName);
            }
        }

        private void InsertTemplate(string sFile, bool bClear)
        {
            SetReadyStatus();
            StreamReader reader;
            try
            {
                reader = new StreamReader(sFile);
            }
            catch (Exception e)
            {
                DisplayReadError(e, sFile);
                return;
            }
            if (reader.BaseStream.Length > tbTemplate.MaxLength)
            {
                MessageBox.Show("The file \"" + sFile + "\" is too large to be used as a template.\r\n" +
                    "It is " +
                    reader.BaseStream.Length + " bytes, and the maximum size is " + tbTemplate.MaxLength + " bytes.\r\n\r\n" +
                    "Please verify that you selected the correct file.",
                    "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                SetStatusError("File too large - " + sFile);
            }
            else
            {
                if (bClear)
                    tbTemplate.Text = reader.ReadToEnd();
                else
                    tbTemplate.Text = sPrefix + reader.ReadToEnd() + sSuffix;
            }
            reader.Close();
        }

        private void ReadJobTemplate(string sJobFile)
        {
            InsertTemplate(sJobFile, true);
        }

        private void InsertVCF(string sVCF)
        {
            InsertTemplate(sVCF, false);
        }

        private void mi_File_Save_Click(object sender, System.EventArgs e)
        {
            SaveTemplateFile();
        }

        private void SaveTemplateFile()
        {
            if (sfdTemplate.ShowDialog() == DialogResult.OK)
            {
                SetReadyStatus();
                StreamWriter writer;
                try
                {
                    writer = new StreamWriter(sfdTemplate.FileName, false);
                }
                catch(Exception e)
                {
                    DisplayWriteError(e, sfdTemplate.FileName);
                    return;
                }
                writer.Write(tbTemplate.Text);
                writer.Close();
            }
        }

        private void btnCreateJobsFile_Click(object sender, System.EventArgs e)
        {
            CreateJobFile();
        }

        private void DisplayWriteError(Exception e, string sFile)
        {
            MessageBox.Show("Could not open file for writing:\r\n\r\n" + sFile + 
                "\r\n\r\nReason: " + e.Message, "Error opening file",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            SetStatusError("Could not create file - " + sFile);
        }

        private void DisplayReadError(Exception e, string sFile)
        {
            MessageBox.Show("Could not open file for reading:\r\n\r\n" + sFile + 
                "\r\n\r\nReason: " + e.Message, "Error opening file",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            SetStatusError("Could not open file - " + sFile);
        }

        private void CreateJobsFile(string sFile)
        {
            SetReadyStatus();
            int iIndex = 1;
            string sJob;
            StreamWriter writer;

            // Open file for writing
            try
            {
                writer = new StreamWriter(sFile, false);
            }
            catch(Exception e)
            {
                DisplayWriteError(e, sFile);
                return;
            }

            int iNumJobs = lvFiles.Items.Count;
            int iMultiplier = 0;
            int iFind = 0;
            while (iFind != -1)
            {
                if (iFind + 1 >= tbTemplate.Text.Length)
                    break;
                iFind = tbTemplate.Text.IndexOf(sJobSeparator, iFind+1);
                iMultiplier++;
            }

            // Write global header (including number of jobs)
            writer.Write(@"// VirtualDub job list (Sylia script format)
// This is a program generated file -- edit at your own risk.
//
// $numjobs " + iNumJobs * iMultiplier + @"
//

");

            foreach(ListViewItem lvItem in lvFiles.Items)
            {
                sJob = CreateOneJob(lvItem.SubItems[0].Text, ref iIndex);
                writer.Write(sJob);
            }

            // Write footer
            writer.Write("// $done\r\n");

            // Close file
            writer.Close();
        }

        private string CreateOneJob(string sFile, ref int iIndex)
        {
            int iSubJob = -1;
            int iFind = 0;
            int iOldFind = 0;
            int iNextJob = 0;

            string sJob = "";
            string sBaseFull = Path.GetDirectoryName(sFile);
            if (!sBaseFull.EndsWith("\\"))
                sBaseFull += "\\";
            sBaseFull += Path.GetFileNameWithoutExtension(sFile);
            string sExtension = Path.GetExtension(sFile);
            if (sExtension.Length > 1)
                sExtension = sExtension.Substring(1);

            iNextJob = tbTemplate.Text.IndexOf(sJobSeparator);
            while (iFind > -1)
            {
                iOldFind = iFind;
                iFind = tbTemplate.Text.IndexOf("%", iFind);
                if ((iNextJob > -1) && (iFind > iNextJob))
                {
                    iNextJob = tbTemplate.Text.IndexOf(sJobSeparator, iNextJob+1);
                    iSubJob++;
                }
                if (iFind < 0)
                    break;

                if (iFind + 1 >= tbTemplate.Text.Length)
                {
                    iOldFind = iFind;
                    break;
                }

                sJob += tbTemplate.Text.Substring(iOldFind, iFind - iOldFind);
                switch(tbTemplate.Text[iFind+1])
                {
                    case 'p':
                        sJob += sFile;
                        break;
                    case 'P':
                        sJob += sFile.Replace("\\", "\\\\");
                        break;
                    case 'b':
                        sJob += sBaseFull;
                        break;
                    case 'B':
                        sJob += sBaseFull.Replace("\\", "\\\\");
                        break;
                    case 'd':
                        sJob += Path.GetDirectoryName(sFile);
                        break;
                    case 'D':
                        sJob += Path.GetDirectoryName(sFile).Replace("\\", "\\\\");
                        break;
                    case 'f':
                        sJob += Path.GetFileName(sFile);
                        break;
                    case 'F':
                        sJob += Path.GetFileNameWithoutExtension(sFile);
                        break;
                    case 'e':
                        sJob += sExtension;
                        break;
                    case 'n':
                        sJob += (iIndex + iSubJob).ToString();
                        break;
                    default:
                        sJob += tbTemplate.Text.Substring(iFind+1,1);
                        break;
                }
                iFind += 2;
                iOldFind = iFind;
            }

            if (iOldFind < tbTemplate.Text.Length)
            {
                sJob += tbTemplate.Text.Substring(iOldFind);
            }

            if (iSubJob > -1)
                iIndex += iSubJob;
            iIndex++;

            if (!sJob.EndsWith("\r\n"))
                sJob += "\r\n";

            return sJob;
        }

        private void mi_Edit_ClearFileList_Click(object sender, System.EventArgs e)
        {
            if (MessageBox.Show("This action will remove all files from the file list.  Continue?", "Confirmation", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                lvFiles.Items.Clear();
            }
        }

        private void mi_Edit_Popup(object sender, System.EventArgs e)
        {
            mi_Edit_ClearFileList.Enabled = (lvFiles.Items.Count > 0);
        }

        private void mi_Edit_Options_Click(object sender, System.EventArgs e)
        {
            m_optionsForm.StartPosition = FormStartPosition.CenterParent;
            m_optionsForm.JobPrefix = sPrefix;
            m_optionsForm.JobSuffix = sSuffix;
            m_optionsForm.ShowDialog();

            sPrefix = m_optionsForm.JobPrefix;
            sSuffix = m_optionsForm.JobSuffix;
        }

        private void VDBAForm_Load(object sender, EventArgs e)
        {
            if (m_options.Help)
                MessageBox.Show(m_options.Usage, "Usage");

            if (m_options.UseJobFile != "")
                InsertTemplate(m_options.UseJobFile, true);

            if (m_options.InputFiles.Count > 0)
            {
                lvFiles.BeginUpdate();
                foreach (string sPath in m_options.InputFiles)
                {
                    string sDir = Path.GetDirectoryName(sPath);
                    string sFile = Path.GetFileName(sPath);
                    if (sDir == null)
                        AddFileToList(sFile);
                    else
                    {
                        foreach (string sFoundFile in Directory.GetFiles(sDir, sFile))
                        {
                            AddFileToList(sFoundFile);
                        }
                    }
                }
                lvFiles.EndUpdate();
            }

            if (m_options.UseOutputFile != "")
                CreateJobsFile(m_options.UseOutputFile);

            if (m_options.ExitAfterOutput)
                Close();
        }
    }
}
