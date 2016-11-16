using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PSEGPortalDownload
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (var browser = new WatiN.Core.IE("http://www.google.com"))
            {
                browser.TextField(WatiN.Core.Find.ByName("q")).TypeText("WatiN");
                browser.Button(WatiN.Core.Find.ByName("btnG")).Click();

            }
        }
    }
}
