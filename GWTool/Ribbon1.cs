using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace GWTool
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f1 = new Form1("通知");
            f1.Show();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f1 = new Form1("");
            f1.Show();
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f1 = new Form1("呈批件");
            f1.Show();
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f1 = new Form1("请示");
            f1.Show();
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f1 = new Form1("上报公文");
            f1.Show();
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            Form_setup fs = new Form_setup();
            fs.Show();
        }
    }
}
