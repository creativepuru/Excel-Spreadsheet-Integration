namespace Excel_Spreadsheet_Integration
{
    partial class Form2
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonForm1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonForm1
            // 
            this.buttonForm1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonForm1.Location = new System.Drawing.Point(27, 12);
            this.buttonForm1.Name = "buttonForm1";
            this.buttonForm1.Size = new System.Drawing.Size(98, 41);
            this.buttonForm1.TabIndex = 0;
            this.buttonForm1.Text = "Form 1";
            this.buttonForm1.UseVisualStyleBackColor = true;
            this.buttonForm1.Click += new System.EventHandler(this.buttonForm1_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1262, 977);
            this.Controls.Add(this.buttonForm1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonForm1;
    }
}