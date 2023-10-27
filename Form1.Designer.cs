namespace Conv.Net
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
            panel1 = new Panel();
            panel3 = new Panel();
            panel5 = new Panel();
            dataGridView = new DataGridView();
            panel4 = new Panel();
            buttonDelALL = new Button();
            buttonConver = new Button();
            buttonOpenFile = new Button();
            panel2 = new Panel();
            panel1.SuspendLayout();
            panel3.SuspendLayout();
            panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            panel4.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.Controls.Add(panel3);
            panel1.Controls.Add(panel2);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(921, 447);
            panel1.TabIndex = 0;
            // 
            // panel3
            // 
            panel3.Controls.Add(panel5);
            panel3.Controls.Add(panel4);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(0, 45);
            panel3.Name = "panel3";
            panel3.Size = new Size(921, 402);
            panel3.TabIndex = 1;
            // 
            // panel5
            // 
            panel5.Controls.Add(dataGridView);
            panel5.Dock = DockStyle.Fill;
            panel5.Location = new Point(0, 0);
            panel5.Name = "panel5";
            panel5.Size = new Size(921, 352);
            panel5.TabIndex = 1;
            // 
            // dataGridView
            // 
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.Location = new Point(0, 0);
            dataGridView.Name = "dataGridView";
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersWidth = 51;
            dataGridView.RowTemplate.Height = 29;
            dataGridView.Size = new Size(921, 352);
            dataGridView.TabIndex = 0;
            // 
            // panel4
            // 
            panel4.Controls.Add(buttonDelALL);
            panel4.Controls.Add(buttonConver);
            panel4.Controls.Add(buttonOpenFile);
            panel4.Dock = DockStyle.Bottom;
            panel4.Location = new Point(0, 352);
            panel4.Name = "panel4";
            panel4.Size = new Size(921, 50);
            panel4.TabIndex = 0;
            // 
            // buttonDelALL
            // 
            buttonDelALL.Location = new Point(457, 11);
            buttonDelALL.Name = "buttonDelALL";
            buttonDelALL.Size = new Size(212, 29);
            buttonDelALL.TabIndex = 2;
            buttonDelALL.Text = "Удалить все строки";
            buttonDelALL.UseVisualStyleBackColor = true;
            buttonDelALL.Click += buttonDelALL_Click;
            // 
            // buttonConver
            // 
            buttonConver.Location = new Point(171, 11);
            buttonConver.Name = "buttonConver";
            buttonConver.Size = new Size(262, 29);
            buttonConver.TabIndex = 1;
            buttonConver.Text = "Сформировать распределение";
            buttonConver.UseVisualStyleBackColor = true;
            buttonConver.Click += buttonConver_Click;
            // 
            // buttonOpenFile
            // 
            buttonOpenFile.Location = new Point(18, 11);
            buttonOpenFile.Name = "buttonOpenFile";
            buttonOpenFile.Size = new Size(147, 29);
            buttonOpenFile.TabIndex = 0;
            buttonOpenFile.Text = "Загрузить excel";
            buttonOpenFile.UseVisualStyleBackColor = true;
            buttonOpenFile.Click += buttonOpenFile_Click;
            // 
            // panel2
            // 
            panel2.Dock = DockStyle.Top;
            panel2.Location = new Point(0, 0);
            panel2.Name = "panel2";
            panel2.Size = new Size(921, 45);
            panel2.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(921, 447);
            Controls.Add(panel1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            panel1.ResumeLayout(false);
            panel3.ResumeLayout(false);
            panel5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            panel4.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Panel panel3;
        private Panel panel5;
        private DataGridView dataGridView;
        private Panel panel4;
        private Button buttonConver;
        private Button buttonOpenFile;
        private Panel panel2;
        public Button buttonDelALL;
    }
}