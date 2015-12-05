namespace UserOptions
{
    partial class RdOptionsControl
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
            this.AllowPublishManagedRpts = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Report Deployer";
            // 
            // AllowPublishManagedRpts
            // 
            this.AllowPublishManagedRpts.AutoSize = true;
            this.AllowPublishManagedRpts.Location = new System.Drawing.Point(17, 30);
            this.AllowPublishManagedRpts.Name = "AllowPublishManagedRpts";
            this.AllowPublishManagedRpts.Size = new System.Drawing.Size(190, 17);
            this.AllowPublishManagedRpts.TabIndex = 14;
            this.AllowPublishManagedRpts.Text = "Allow Publishing Managed Reports";
            this.AllowPublishManagedRpts.UseVisualStyleBackColor = true;
            this.AllowPublishManagedRpts.CheckedChanged += new System.EventHandler(this.AllowPublishManagedRpts_CheckedChanged);
            // 
            // RdOptionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.AllowPublishManagedRpts);
            this.Controls.Add(this.label1);
            this.Name = "RdOptionsControl";
            this.Size = new System.Drawing.Size(373, 352);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox AllowPublishManagedRpts;


    }
}
