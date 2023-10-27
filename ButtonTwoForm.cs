using System;
using System.Windows.Forms;

namespace Conv.Net
{
    public partial class ButtonTwoForm : Form
    {
        public bool left = false;
        public bool right = false;
        public ButtonTwoForm(string textButtonLeft, string textButtonRight)
        {
            InitializeComponent();
            buttonLeft.Text = textButtonLeft;
            buttonRight.Text = textButtonRight;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void buttonLeft_Click(object sender, EventArgs e)
        {
            left = true;
            DialogResult = DialogResult.OK;
        }

        private void buttonRight_Click(object sender, EventArgs e)
        {
            right = true;
            DialogResult = DialogResult.OK;
        }

        private void ButtonTwoForm_Load(object sender, EventArgs e)
        {

        }
    }
}
