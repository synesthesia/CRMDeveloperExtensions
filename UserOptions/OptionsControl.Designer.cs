namespace UserOptions
{
    partial class OptionsControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.DefaultSdkVersion = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.DefaultKeyFileName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.DefaultWebBrowser = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.EnableSdkSearch = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // DefaultSdkVersion
            // 
            this.DefaultSdkVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DefaultSdkVersion.FormattingEnabled = true;
            this.DefaultSdkVersion.Items.AddRange(new object[] {
            "CRM 2011 (5.0.X)",
            "CRM 2013 (6.0.X)",
            "CRM 2013 (6.1.X)",
            "CRM 2015 (7.0.X)",
            "CRM 2015 (7.1.X)",
            "CRM 2016 (8.0.X)",
            "CRM 2016 (8.1.X)",
            "CRM 2016 (8.2.X)"});
            this.DefaultSdkVersion.Location = new System.Drawing.Point(151, 26);
            this.DefaultSdkVersion.Name = "DefaultSdkVersion";
            this.DefaultSdkVersion.Size = new System.Drawing.Size(163, 21);
            this.DefaultSdkVersion.TabIndex = 1;
            this.DefaultSdkVersion.SelectedIndexChanged += new System.EventHandler(this.DefaultSdkVersion_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Default CRM SDK Version";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Default Key File Name";
            // 
            // DefaultKeyFileName
            // 
            this.DefaultKeyFileName.Location = new System.Drawing.Point(151, 52);
            this.DefaultKeyFileName.Name = "DefaultKeyFileName";
            this.DefaultKeyFileName.Size = new System.Drawing.Size(134, 20);
            this.DefaultKeyFileName.TabIndex = 4;
            this.DefaultKeyFileName.Text = "MyKey";
            this.DefaultKeyFileName.TextChanged += new System.EventHandler(this.DefaultKeyFileName_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(287, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = ".snk";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(17, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Templates";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(17, 74);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "External Content";
            // 
            // DefaultWebBrowser
            // 
            this.DefaultWebBrowser.AutoSize = true;
            this.DefaultWebBrowser.Location = new System.Drawing.Point(17, 96);
            this.DefaultWebBrowser.Name = "DefaultWebBrowser";
            this.DefaultWebBrowser.Size = new System.Drawing.Size(149, 17);
            this.DefaultWebBrowser.TabIndex = 10;
            this.DefaultWebBrowser.Text = "Use Default Web Browser";
            this.DefaultWebBrowser.UseVisualStyleBackColor = true;
            this.DefaultWebBrowser.CheckedChanged += new System.EventHandler(this.DefaultWebBrowser_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(17, 122);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(93, 13);
            this.label7.TabIndex = 11;
            this.label7.Text = "CRM SDK Search";
            // 
            // EnableSdkSearch
            // 
            this.EnableSdkSearch.AutoSize = true;
            this.EnableSdkSearch.Checked = true;
            this.EnableSdkSearch.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EnableSdkSearch.Location = new System.Drawing.Point(17, 144);
            this.EnableSdkSearch.Name = "EnableSdkSearch";
            this.EnableSdkSearch.Size = new System.Drawing.Size(148, 17);
            this.EnableSdkSearch.TabIndex = 12;
            this.EnableSdkSearch.Text = "Enable CRM SDK Search";
            this.EnableSdkSearch.UseVisualStyleBackColor = true;
            this.EnableSdkSearch.CheckedChanged += new System.EventHandler(this.EnableSdkSearch_CheckedChanged);
            // 
            // OptionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EnableSdkSearch);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.DefaultWebBrowser);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.DefaultKeyFileName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DefaultSdkVersion);
            this.Name = "OptionsControl";
            this.Size = new System.Drawing.Size(373, 352);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox DefaultSdkVersion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox DefaultKeyFileName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox DefaultWebBrowser;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox EnableSdkSearch;
    }
}
