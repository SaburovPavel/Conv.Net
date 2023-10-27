namespace Conv.Net
{
    partial class ButtonTwoForm
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
            panel1 = new Panel();
            panel3 = new Panel();
            buttonCancel = new Button();
            panel2 = new Panel();
            panel5 = new Panel();
            buttonRight = new Button();
            panel4 = new Panel();
            buttonLeft = new Button();
            panel1.SuspendLayout();
            panel3.SuspendLayout();
            panel2.SuspendLayout();
            panel5.SuspendLayout();
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
            panel1.Size = new Size(704, 252);
            panel1.TabIndex = 0;
            // 
            // panel3
            // 
            panel3.Controls.Add(buttonCancel);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(491, 0);
            panel3.Name = "panel3";
            panel3.Size = new Size(213, 252);
            panel3.TabIndex = 1;
            // 
            // buttonCancel
            // 
            buttonCancel.BackColor = Color.Orange;
            buttonCancel.Dock = DockStyle.Fill;
            buttonCancel.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            buttonCancel.Location = new Point(0, 0);
            buttonCancel.Name = "buttonCancel";
            buttonCancel.Size = new Size(213, 252);
            buttonCancel.TabIndex = 0;
            buttonCancel.Text = "Отмена";
            buttonCancel.UseVisualStyleBackColor = false;
            buttonCancel.Click += buttonCancel_Click;
            // 
            // panel2
            // 
            panel2.Controls.Add(panel5);
            panel2.Controls.Add(panel4);
            panel2.Dock = DockStyle.Left;
            panel2.Location = new Point(0, 0);
            panel2.Name = "panel2";
            panel2.Size = new Size(491, 252);
            panel2.TabIndex = 0;
            // 
            // panel5
            // 
            panel5.Controls.Add(buttonRight);
            panel5.Dock = DockStyle.Fill;
            panel5.Location = new Point(247, 0);
            panel5.Name = "panel5";
            panel5.Size = new Size(244, 252);
            panel5.TabIndex = 1;
            // 
            // buttonRight
            // 
            buttonRight.BackColor = Color.LightBlue;
            buttonRight.Dock = DockStyle.Fill;
            buttonRight.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            buttonRight.Location = new Point(0, 0);
            buttonRight.Name = "buttonRight";
            buttonRight.Size = new Size(244, 252);
            buttonRight.TabIndex = 0;
            buttonRight.Text = "button2";
            buttonRight.UseVisualStyleBackColor = false;
            buttonRight.Click += buttonRight_Click;
            // 
            // panel4
            // 
            panel4.Controls.Add(buttonLeft);
            panel4.Dock = DockStyle.Left;
            panel4.Location = new Point(0, 0);
            panel4.Name = "panel4";
            panel4.Size = new Size(247, 252);
            panel4.TabIndex = 0;
            // 
            // buttonLeft
            // 
            buttonLeft.BackColor = Color.LightGreen;
            buttonLeft.Dock = DockStyle.Fill;
            buttonLeft.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            buttonLeft.Location = new Point(0, 0);
            buttonLeft.Name = "buttonLeft";
            buttonLeft.Size = new Size(247, 252);
            buttonLeft.TabIndex = 0;
            buttonLeft.Text = "button1";
            buttonLeft.UseVisualStyleBackColor = false;
            buttonLeft.Click += buttonLeft_Click;
            // 
            // ButtonTwoForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = buttonCancel;
            ClientSize = new Size(704, 252);
            Controls.Add(panel1);
            Name = "ButtonTwoForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "ButtonTwoForm";
            Load += ButtonTwoForm_Load;
            panel1.ResumeLayout(false);
            panel3.ResumeLayout(false);
            panel2.ResumeLayout(false);
            panel5.ResumeLayout(false);
            panel4.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Panel panel3;
        private Button buttonCancel;
        private Panel panel2;
        private Panel panel5;
        private Button buttonRight;
        private Panel panel4;
        private Button buttonLeft;
    }
}