namespace UserOptions
{
    partial class SpOptionsControl
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SpName = new System.Windows.Forms.TextBox();
            this.SaveSolution = new System.Windows.Forms.CheckBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.OpenFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Solution Packager";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "SDK Bin Folder";
            // 
            // SpName
            // 
            this.SpName.CausesValidation = false;
            this.SpName.Location = new System.Drawing.Point(151, 51);
            this.SpName.Name = "SpName";
            this.SpName.ReadOnly = true;
            this.SpName.Size = new System.Drawing.Size(140, 20);
            this.SpName.TabIndex = 2;
            this.SpName.TabStop = false;
            // 
            // SaveSolution
            // 
            this.SaveSolution.AutoSize = true;
            this.SaveSolution.Location = new System.Drawing.Point(17, 77);
            this.SaveSolution.Name = "SaveSolution";
            this.SaveSolution.Size = new System.Drawing.Size(163, 17);
            this.SaveSolution.TabIndex = 4;
            this.SaveSolution.Text = "Save Solution Files in Project";
            this.SaveSolution.UseVisualStyleBackColor = true;
            this.SaveSolution.CheckedChanged += new System.EventHandler(this.SaveSolution_CheckedChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // OpenFolder
            // 
            this.OpenFolder.Location = new System.Drawing.Point(291, 50);
            this.OpenFolder.Name = "OpenFolder";
            this.OpenFolder.Size = new System.Drawing.Size(24, 22);
            this.OpenFolder.TabIndex = 5;
            this.OpenFolder.Text = "...";
            this.OpenFolder.UseVisualStyleBackColor = true;
            this.OpenFolder.Click += new System.EventHandler(this.OpenFolder_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(17, 10);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(242, 13);
            this.label9.TabIndex = 21;
            this.label9.Text = "Changes Require Restart of Visual Studio";
            // 
            // SpOptionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label9);
            this.Controls.Add(this.OpenFolder);
            this.Controls.Add(this.SaveSolution);
            this.Controls.Add(this.SpName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "SpOptionsControl";
            this.Size = new System.Drawing.Size(373, 352);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox SpName;
        private System.Windows.Forms.CheckBox SaveSolution;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button OpenFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Label label9;


    }
}
