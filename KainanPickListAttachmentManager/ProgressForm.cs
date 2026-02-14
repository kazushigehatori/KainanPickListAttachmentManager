using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KainanPickListAttachmentManager
{
    public partial class ProgressForm : Form
    {

        public ProgressForm()
        {
            InitializeComponent();
            progressBar1.Minimum = 0;
        }

        public void UpdateProgress(int value, int max, string message)
        {
            progressBar1.Maximum = max;
            progressBar1.Value = value;
            label1.Text = message;
            Application.DoEvents(); // UIを即時更新
        }

    }
}
