namespace WindowsFormsApplication1
{
    partial class Form1
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
            this.demo = new System.Windows.Forms.Button();
            this.btnWord = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // demo
            // 
            this.demo.Location = new System.Drawing.Point(12, 12);
            this.demo.Name = "demo";
            this.demo.Size = new System.Drawing.Size(75, 23);
            this.demo.TabIndex = 0;
            this.demo.Text = "Demo";
            this.demo.UseVisualStyleBackColor = true;
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(12, 41);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(75, 23);
            this.btnWord.TabIndex = 0;
            this.btnWord.Text = "Word";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWordClick);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(12, 70);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExcel.TabIndex = 0;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcelClick);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "PowerPoint";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnPptClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.demo);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button demo;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button button1;
    }
}

