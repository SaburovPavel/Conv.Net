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
            panel1 = new System.Windows.Forms.Panel();
            panel3 = new System.Windows.Forms.Panel();
            buttonCancel = new System.Windows.Forms.Button();
            panel2 = new System.Windows.Forms.Panel();
            panel5 = new System.Windows.Forms.Panel();
            buttonRight = new System.Windows.Forms.Button();
            panel4 = new System.Windows.Forms.Panel();
            buttonLeft = new System.Windows.Forms.Button();
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
            panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            panel1.Location = new System.Drawing.Point(0, 0);
            panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            panel1.Name = "panel1";
            panel1.Size = new System.Drawing.Size(616, 189);
            panel1.TabIndex = 0;
            // 
            // panel3
            // 
            panel3.Controls.Add(buttonCancel);
            panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            panel3.Location = new System.Drawing.Point(430, 0);
            panel3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            panel3.Name = "panel3";
            panel3.Size = new System.Drawing.Size(186, 189);
            panel3.TabIndex = 1;
            // 
            // buttonCancel
            // 
            buttonCancel.BackColor = System.Drawing.Color.Orange;
            buttonCancel.Dock = System.Windows.Forms.DockStyle.Fill;
            buttonCancel.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            buttonCancel.Location = new System.Drawing.Point(0, 0);
            buttonCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            buttonCancel.Name = "buttonCancel";
            buttonCancel.Size = new System.Drawing.Size(186, 189);
            buttonCancel.TabIndex = 0;
            buttonCancel.Text = "Отмена";
            buttonCancel.UseVisualStyleBackColor = false;
            buttonCancel.Click += buttonCancel_Click;
            // 
            // panel2
            // 
            panel2.Controls.Add(panel5);
            panel2.Controls.Add(panel4);
            panel2.Dock = System.Windows.Forms.DockStyle.Left;
            panel2.Location = new System.Drawing.Point(0, 0);
            panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            panel2.Name = "panel2";
            panel2.Size = new System.Drawing.Size(430, 189);
            panel2.TabIndex = 0;
            // 
            // panel5
            // 
            panel5.Controls.Add(buttonRight);
            panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            panel5.Location = new System.Drawing.Point(216, 0);
            panel5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            panel5.Name = "panel5";
            panel5.Size = new System.Drawing.Size(214, 189);
            panel5.TabIndex = 1;
            // 
            // buttonRight
            // 
            buttonRight.BackColor = System.Drawing.Color.LightBlue;
            buttonRight.Dock = System.Windows.Forms.DockStyle.Fill;
            buttonRight.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            buttonRight.Location = new System.Drawing.Point(0, 0);
            buttonRight.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            buttonRight.Name = "buttonRight";
            buttonRight.Size = new System.Drawing.Size(214, 189);
            buttonRight.TabIndex = 0;
            buttonRight.Text = "button2";
            buttonRight.UseVisualStyleBackColor = false;
            buttonRight.Click += buttonRight_Click;
            // 
            // panel4
            // 
            panel4.Controls.Add(buttonLeft);
            panel4.Dock = System.Windows.Forms.DockStyle.Left;
            panel4.Location = new System.Drawing.Point(0, 0);
            panel4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            panel4.Name = "panel4";
            panel4.Size = new System.Drawing.Size(216, 189);
            panel4.TabIndex = 0;
            // 
            // buttonLeft
            // 
            buttonLeft.BackColor = System.Drawing.Color.LightGreen;
            buttonLeft.Dock = System.Windows.Forms.DockStyle.Fill;
            buttonLeft.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            buttonLeft.Location = new System.Drawing.Point(0, 0);
            buttonLeft.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            buttonLeft.Name = "buttonLeft";
            buttonLeft.Size = new System.Drawing.Size(216, 189);
            buttonLeft.TabIndex = 0;
            buttonLeft.Text = "button1";
            buttonLeft.UseVisualStyleBackColor = false;
            buttonLeft.Click += buttonLeft_Click;
            // 
            // ButtonTwoForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            CancelButton = buttonCancel;
            ClientSize = new System.Drawing.Size(616, 189);
            Controls.Add(panel1);
            Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            Name = "ButtonTwoForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "ButtonTwoForm";
            panel1.ResumeLayout(false);
            panel3.ResumeLayout(false);
            panel2.ResumeLayout(false);
            panel5.ResumeLayout(false);
            panel4.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button buttonRight;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button buttonLeft;
    }
}