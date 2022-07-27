namespace PentruPomagaDeLaNokia
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SelectExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ColumnRead = new System.Windows.Forms.NumericUpDown();
            this.ColumnWrite = new System.Windows.Forms.NumericUpDown();
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.ColumnRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ColumnWrite)).BeginInit();
            this.SuspendLayout();
            // 
            // SelectExcel
            // 
            this.SelectExcel.Enabled = false;
            this.SelectExcel.Location = new System.Drawing.Point(12, 69);
            this.SelectExcel.Name = "SelectExcel";
            this.SelectExcel.Size = new System.Drawing.Size(138, 30);
            this.SelectExcel.TabIndex = 0;
            this.SelectExcel.Text = "Select excel to modify";
            this.SelectExcel.UseVisualStyleBackColor = true;
            this.SelectExcel.Click += new System.EventHandler(this.SelectExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(164, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Column number to read data:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(167, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = "Column number to write data:";
            // 
            // ColumnRead
            // 
            this.ColumnRead.Location = new System.Drawing.Point(182, 7);
            this.ColumnRead.Name = "ColumnRead";
            this.ColumnRead.Size = new System.Drawing.Size(120, 23);
            this.ColumnRead.TabIndex = 6;
            this.ColumnRead.ValueChanged += new System.EventHandler(this.ColumnRead_ValueChanged);
            // 
            // ColumnWrite
            // 
            this.ColumnWrite.Location = new System.Drawing.Point(182, 40);
            this.ColumnWrite.Name = "ColumnWrite";
            this.ColumnWrite.Size = new System.Drawing.Size(120, 23);
            this.ColumnWrite.TabIndex = 7;
            this.ColumnWrite.ValueChanged += new System.EventHandler(this.ColumnWrite_ValueChanged);
            // 
            // OpenFileDialog1
            // 
            this.OpenFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(333, 113);
            this.Controls.Add(this.ColumnWrite);
            this.Controls.Add(this.ColumnRead);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SelectExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.ColumnRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ColumnWrite)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button SelectExcel;
        private Label label1;
        private Label label2;
        private NumericUpDown ColumnRead;
        private NumericUpDown ColumnWrite;
        private OpenFileDialog OpenFileDialog1;
    }
}