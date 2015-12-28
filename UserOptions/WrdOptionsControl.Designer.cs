namespace UserOptions
{
    partial class WrdOptionsControl
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
            this.AllowPublishManaged = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Web Resource Deployer";
            // 
            // AllowPublishManaged
            // 
            this.AllowPublishManaged.AutoSize = true;
            this.AllowPublishManaged.Location = new System.Drawing.Point(17, 30);
            this.AllowPublishManaged.Name = "AllowPublishManaged";
            this.AllowPublishManaged.Size = new System.Drawing.Size(230, 17);
            this.AllowPublishManaged.TabIndex = 7;
            this.AllowPublishManaged.Text = "Allow Publishing Managed Web Resources";
            this.AllowPublishManaged.UseVisualStyleBackColor = true;
            this.AllowPublishManaged.CheckedChanged += new System.EventHandler(this.AllowPublishManaged_CheckedChanged);
            // 
            // WrdOptionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.AllowPublishManaged);
            this.Controls.Add(this.label1);
            this.Name = "WrdOptionsControl";
            this.Size = new System.Drawing.Size(373, 352);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox AllowPublishManaged;


    }
}
