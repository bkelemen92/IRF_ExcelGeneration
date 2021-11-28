
namespace ExcelGeneration
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
            this.lbl_Magic = new System.Windows.Forms.Label();
            this.btn_Exit = new System.Windows.Forms.Button();
            this.btn_CloseExcelExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_Magic
            // 
            this.lbl_Magic.AutoSize = true;
            this.lbl_Magic.Location = new System.Drawing.Point(208, 27);
            this.lbl_Magic.Name = "lbl_Magic";
            this.lbl_Magic.Size = new System.Drawing.Size(35, 13);
            this.lbl_Magic.TabIndex = 0;
            this.lbl_Magic.Text = "label1";
            // 
            // btn_Exit
            // 
            this.btn_Exit.Location = new System.Drawing.Point(55, 102);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(75, 23);
            this.btn_Exit.TabIndex = 1;
            this.btn_Exit.Text = "button1";
            this.btn_Exit.UseVisualStyleBackColor = true;
            // 
            // btn_CloseExcelExit
            // 
            this.btn_CloseExcelExit.Location = new System.Drawing.Point(137, 102);
            this.btn_CloseExcelExit.Name = "btn_CloseExcelExit";
            this.btn_CloseExcelExit.Size = new System.Drawing.Size(75, 23);
            this.btn_CloseExcelExit.TabIndex = 2;
            this.btn_CloseExcelExit.Text = "button2";
            this.btn_CloseExcelExit.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 181);
            this.Controls.Add(this.btn_CloseExcelExit);
            this.Controls.Add(this.btn_Exit);
            this.Controls.Add(this.lbl_Magic);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_Magic;
        private System.Windows.Forms.Button btn_Exit;
        private System.Windows.Forms.Button btn_CloseExcelExit;
    }
}

