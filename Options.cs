using System;
using System.Drawing;
using Microsoft.Win32;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;


namespace VDBatch
{
	/// <summary>
	/// Summary description for Options.
	/// </summary>
	public class Options : System.Windows.Forms.Form
	{
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox tbJobsPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Label labelJobPrefix;
        private Label labelJobSuffix;
        private TextBox tbJobPrefix;
        private TextBox tbJobSuffix;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        public string JobsPath
        {
            get { return tbJobsPath.Text; }
        }

        public string JobPrefix
        {
            get { return tbJobPrefix.Text; }
            set { tbJobPrefix.Text = value; }
        }

        public string JobSuffix
        {
            get { return tbJobSuffix.Text; }
            set { tbJobSuffix.Text = value; }
        }

        public Options()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbJobsPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labelJobPrefix = new System.Windows.Forms.Label();
            this.labelJobSuffix = new System.Windows.Forms.Label();
            this.tbJobPrefix = new System.Windows.Forms.TextBox();
            this.tbJobSuffix = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(160, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Default VirtualDub.jobs path:";
            // 
            // tbJobsPath
            // 
            this.tbJobsPath.Location = new System.Drawing.Point(160, 5);
            this.tbJobsPath.Name = "tbJobsPath";
            this.tbJobsPath.Size = new System.Drawing.Size(312, 20);
            this.tbJobsPath.TabIndex = 1;
            this.tbJobsPath.Text = "C:\\Program Files\\VirtualDub\\VirtualDub.jobs";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(480, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "&Browse";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(317, 266);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(397, 266);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            // 
            // labelJobPrefix
            // 
            this.labelJobPrefix.Location = new System.Drawing.Point(8, 31);
            this.labelJobPrefix.Name = "labelJobPrefix";
            this.labelJobPrefix.Size = new System.Drawing.Size(160, 23);
            this.labelJobPrefix.TabIndex = 0;
            this.labelJobPrefix.Text = "Job prefix:";
            // 
            // labelJobSuffix
            // 
            this.labelJobSuffix.Location = new System.Drawing.Point(8, 150);
            this.labelJobSuffix.Name = "labelJobSuffix";
            this.labelJobSuffix.Size = new System.Drawing.Size(160, 23);
            this.labelJobSuffix.TabIndex = 0;
            this.labelJobSuffix.Text = "Job suffix:";
            // 
            // tbJobPrefix
            // 
            this.tbJobPrefix.AcceptsReturn = true;
            this.tbJobPrefix.Location = new System.Drawing.Point(160, 28);
            this.tbJobPrefix.Multiline = true;
            this.tbJobPrefix.Name = "tbJobPrefix";
            this.tbJobPrefix.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbJobPrefix.Size = new System.Drawing.Size(312, 113);
            this.tbJobPrefix.TabIndex = 1;
            // 
            // tbJobSuffix
            // 
            this.tbJobSuffix.AcceptsReturn = true;
            this.tbJobSuffix.Location = new System.Drawing.Point(160, 147);
            this.tbJobSuffix.Multiline = true;
            this.tbJobSuffix.Name = "tbJobSuffix";
            this.tbJobSuffix.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbJobSuffix.Size = new System.Drawing.Size(312, 113);
            this.tbJobSuffix.TabIndex = 1;
            // 
            // Options
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(560, 293);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.tbJobSuffix);
            this.Controls.Add(this.tbJobPrefix);
            this.Controls.Add(this.tbJobsPath);
            this.Controls.Add(this.labelJobSuffix);
            this.Controls.Add(this.labelJobPrefix);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Name = "Options";
            this.Text = "Options";
            this.Load += new System.EventHandler(this.Options_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
		#endregion

        private void Options_Load(object sender, System.EventArgs e)
        {
            ReadFromRegistry();
        }

        private void btnOK_Click(object sender, System.EventArgs e)
        {
            WriteToRegistry();
            Close();
        }

        public void ReadFromRegistry()
        {
            RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\EDV Software\\VDBatch");
            tbJobsPath.Text = rk.GetValue("JobsPath", tbJobsPath.Text).ToString();
            tbJobPrefix.Text = rk.GetValue("Prefix", tbJobPrefix.Text).ToString();
            tbJobSuffix.Text = rk.GetValue("Suffix", tbJobSuffix.Text).ToString();
            rk.Close();
        }

        private void WriteToRegistry()
        {
            RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\EDV Software\\VDBatch");
            rk.SetValue("JobsPath", tbJobsPath.Text);
            rk.SetValue("Prefix", tbJobPrefix.Text);
            rk.SetValue("Suffix", tbJobSuffix.Text);
            rk.Close();
        }

        private void btnBrowse_Click(object sender, System.EventArgs e)
        {
            openFileDialog1.FileName = tbJobsPath.Text;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbJobsPath.Text = openFileDialog1.FileName;
            }
        }
	}
}
