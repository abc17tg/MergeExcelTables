using System;
using System.Windows.Forms;

namespace MergeExcelsTables
{
    public partial class DelimiterForm : Form
    {
        public string Delimiter { get; private set; }

        public DelimiterForm()
        {
            InitializeComponent();
            FormClosed += (s, e) => { if (string.IsNullOrEmpty(Delimiter)) Delimiter = "\t"; };
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            Delimiter = delimiterTextBox.Text;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
