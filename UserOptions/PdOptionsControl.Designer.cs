namespace UserOptions
{
    partial class PdOptionsControl
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
            this.PrtName = new System.Windows.Forms.TextBox();
            this.OpenFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.EnablePdContextTemplates = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Plug-in Deployer";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Registration Tool Folder";
            // 
            // PrtName
            // 
            this.PrtName.CausesValidation = false;
            this.PrtName.Location = new System.Drawing.Point(151, 51);
            this.PrtName.Name = "PrtName";
            this.PrtName.ReadOnly = true;
            this.PrtName.Size = new System.Drawing.Size(140, 20);
            this.PrtName.TabIndex = 2;
            this.PrtName.TabStop = false;
            // 
            // OpenFolder
            // 
            this.OpenFolder.Location = new System.Drawing.Point(291, 50);
            this.OpenFolder.Name = "OpenFolder";
            this.OpenFolder.Size = new System.Drawing.Size(24, 22);
            this.OpenFolder.TabIndex = 6;
            this.OpenFolder.Text = "...";
            this.OpenFolder.UseVisualStyleBackColor = true;
            this.OpenFolder.Click += new System.EventHandler(this.OpenFolder_Click);
            // 
            // EnablePdContextTemplates
            // 
            this.EnablePdContextTemplates.AutoSize = true;
            this.EnablePdContextTemplates.Location = new System.Drawing.Point(20, 77);
            this.EnablePdContextTemplates.Name = "EnablePdContextTemplates";
            this.EnablePdContextTemplates.Size = new System.Drawing.Size(262, 17);
            this.EnablePdContextTemplates.TabIndex = 15;
            this.EnablePdContextTemplates.Text = "Enable Context Menu Add -> New Item Templates";
            this.EnablePdContextTemplates.UseVisualStyleBackColor = true;
            this.EnablePdContextTemplates.CheckedChanged += new System.EventHandler(this.EnablePdContextTemplates_CheckedChanged);
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
            // PdOptionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label9);
            this.Controls.Add(this.EnablePdContextTemplates);
            this.Controls.Add(this.OpenFolder);
            this.Controls.Add(this.PrtName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "PdOptionsControl";
            this.Size = new System.Drawing.Size(373, 352);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox PrtName;
        private System.Windows.Forms.Button OpenFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.CheckBox EnablePdContextTemplates;
        private System.Windows.Forms.Label label9;


    }
}
