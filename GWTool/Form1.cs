using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GWTool
{
    public partial class Form1 : Form
    {
        private Label lable_qf = new Label();
        private TextBox textbox_qf = new TextBox();
        private string zhonglei = "";
        Object Nothing = System.Reflection.Missing.Value;
        public Form1(string action)
        {
            zhonglei = action;
            InitializeComponent();
            if (action == "呈批件" || action == "请示" || action == "上报公文")
            {
                ShangXingWenLayout();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void TongzhiLayout()
        {
            
        }

        private void ShangXingWenLayout()
        {
            lable_qf.Location = new Point(400, 134);
            lable_qf.AutoSize = true;
            lable_qf.Text = "签发人：";
            textbox_qf.Location = new Point(455, 129);
            groupBox1.Controls.Add(lable_qf);
            groupBox1.Controls.Add(textbox_qf);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBox_fwdw_TextChanged(object sender, EventArgs e)
        {
            label12.Text = comboBox_fwdw.Text;
        }

        /**
         * 设置通知头
         * 调用前将光标移动到最前边
         */
        private void TongZhiTou(Word.Application wordApp)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Selection ws = wordApp.Selection;
            ws.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ws.Font.Size = 16;
            ws.Font.Name = "黑体";
            ws.TypeText("01\r\n");
            ws.TypeText(comboBox1.Text + "\r\n");
            ws.TypeText(comboBox2.Text + "\r\n");
            Word.Table newTable = wordApp.ActiveDocument.Tables.Add(ws.Range, 1, 2, ref Nothing, ref Nothing);
            newTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            newTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            newTable.Cell(1, 1).Range.Font.Size = 22;
            newTable.Cell(1, 1).Range.Font.Name = "宋体";
            newTable.Cell(1, 1).Range.Font.ColorIndex = Word.WdColorIndex.wdRed;
            newTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            newTable.Cell(1, 1).Range.Text = "中国人民\r\n解放军";
            newTable.Columns[1].AutoFit();
            newTable.Cell(1, 1).Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            newTable.Cell(1, 1).Range.ParagraphFormat.LineSpacing = 24;
            newTable.Cell(1, 2).Range.Font.Size = 22;
            newTable.Cell(1, 2).Range.Font.Name = "宋体";
            newTable.Cell(1, 2).Range.Font.ColorIndex = Word.WdColorIndex.wdRed;
            newTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            newTable.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            newTable.Cell(1, 2).Range.Text = comboBox_fwdw.Text;
            newTable.Columns[2].AutoFit();
            newTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            object count = 2;
            object WdLine = Word.WdUnits.wdLine;//换一行;
            ws.MoveDown(ref WdLine, ref count, ref Nothing);
            ws.TypeParagraph();
            ws.TypeParagraph();
            ws.Paragraphs[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            ws.Paragraphs[1].Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
            ws.Paragraphs[1].Borders.OutsideColor = Word.WdColor.wdColorRed;
            ws.Paragraphs[1].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            ws.Paragraphs[1].Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            ws.Paragraphs[1].Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            ws.Font.Name = "仿宋";
            ws.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ws.TypeText(comboBox4.Text + "〔" + dateTimePicker1.Text + "〕" + textBox2.Text + "号");
        }

        /**
         * 设置通知正文格式
         * */
        private void TongZhiZhengwen(Word.Application wordApp)
        {
            ResetGuangBiao(wordApp);
            Word.Selection ws = wordApp.Selection;
            Word.Document ad = wordApp.ActiveDocument;
            ad.Paragraphs[1].Range.Font.Size = 22;
            ad.Paragraphs[1].Range.Font.Name = "宋体";
            ad.Paragraphs[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ad.Paragraphs[2].Range.Font.Size = 16;
            ad.Paragraphs[2].Range.Font.Name = "楷体";
            ad.Paragraphs[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            for (int i = 3; i <= ad.Paragraphs.Count; i++)
            {
                ad.Paragraphs[i].Range.Font.Size = 16;
                ad.Paragraphs[i].Range.Font.Name = "仿宋";
                ad.Paragraphs[i].Range.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            }
        }

        private void ResetGuangBiao(Word.Application wordApp)
        {
            object dummy = System.Reflection.Missing.Value;
            object what = Word.WdGoToItem.wdGoToLine;
            object which = Word.WdGoToDirection.wdGoToFirst;
            object count = 1;
            Word.Selection ws = wordApp.Selection;
            ws.GoTo(ref what, ref which, ref count, ref dummy);
        }

        private void GotoLastLine(Word.Application wordApp)
        {
            object dummy = System.Reflection.Missing.Value;
            object what = Word.WdGoToItem.wdGoToLine;
            object which = Word.WdGoToDirection.wdGoToLast;
            object count = 99999999;
            wordApp.Selection.GoTo(ref what, ref which, ref count, ref dummy);
        }

        public void GotoLastCharacter(Word.Selection selection)
        {
            object dummy = System.Reflection.Missing.Value;
            object count = 99999999;
            object Unit = Word.WdUnits.wdCharacter;
            selection.MoveRight(ref Unit, ref count, ref dummy);
        }

        private void TongZhi()
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;
            Word.Selection ws = wordApp.Selection;
            if (!checkBox1.Checked)
            {
                TongZhiZhengwen(wordApp);
            }
            SetPageStyle(wordApp);
            ResetGuangBiao(wordApp);
            ws.Font.Size = 16;
            ws.TypeParagraph();
            ws.TypeParagraph();
            ResetGuangBiao(wordApp);
            TongZhiTou(wordApp);
            Chengban_Tongzhi(wordApp);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TongZhi();
            Close();
        }

        private void SetPageStyle(Word.Application wordApp)
        {
            wordApp = Globals.ThisAddIn.Application;
            wordApp.ActiveDocument.PageSetup.TopMargin = wordApp.CentimetersToPoints(float.Parse("3.7"));
            wordApp.ActiveDocument.PageSetup.BottomMargin = wordApp.CentimetersToPoints(float.Parse("3.5"));
            wordApp.ActiveDocument.PageSetup.LeftMargin = wordApp.CentimetersToPoints(float.Parse("2.8"));
            wordApp.ActiveDocument.PageSetup.RightMargin = wordApp.CentimetersToPoints(float.Parse("2.6"));
        }

        private void Chengban_Tongzhi(Word.Application wordApp)
        {
            GotoLastLine(wordApp);
            Word.Selection ws = wordApp.Selection;
            GotoLastCharacter(ws);
            ws.TypeParagraph();
            ws.TypeParagraph();
            ws.Font.Size = 14;
            ws.Font.Name = "仿宋";
            ws.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            ws.TypeText("抄送：" + textBox1.Text + ChaoSong_Space("抄送：" + textBox1.Text + "（共印" + textBox3.Text + "份）") + "（共印" + textBox3.Text + "份）");
            ws.TypeParagraph();
            //计算承办信息第二行空格数
            string cSqace = ChengBan_Space("承办单位：" + comboBox5.Text + "承办人：" + textBox4.Text + "电话：" + textBox5.Text);
            ws.TypeText("承办单位：" + comboBox5.Text + cSqace + "承办人：" + textBox4.Text + cSqace + "电话：" + textBox5.Text);
            ws.TypeParagraph();
            ws.TypeText(label12.Text + YinFa_Space(label12.Text + dateTimePicker2.Text + label13.Text) + dateTimePicker2.Text + label13.Text);
            Word.Document ad = wordApp.ActiveDocument;
            int p = ad.Paragraphs.Count;
            ad.Paragraphs[p - 2].Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            ad.Paragraphs[p - 2].Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
            ad.Paragraphs[p - 1].Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            ad.Paragraphs[p - 1].Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth075pt;
            ad.Paragraphs[p - 1].Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            ad.Paragraphs[p - 1].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth075pt;
            ad.Paragraphs[p].Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            ad.Paragraphs[p].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
        }

        private string ChaoSong_Space(string str)
        {
            string restr = "";
            if (str.Length < 46)
            {
                for (int i = 0; i < 46 - str.Length; i++)
                {
                    restr += " ";
                }
            }
            return restr;
        }

        private string ChengBan_Space(string str)
        {
            string restr = "";
            if (str.Length < 46)
            {
                for (int i = 0; i < (46 - str.Length) / 2; i++)
                {
                    restr += " ";
                }
            }
            return restr;
        }

        private string YinFa_Space(string str)
        {
            //MessageBox.Show(str + "\n" + str.Length.ToString(), "输出");
            string restr = "";
            if (str.Length < 43)
            {
                for (int i = 0; i < 43 - str.Length; i++)
                {
                    restr += " ";
                }
            }
            return restr;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                radioButton2.Checked = false;
            else
                radioButton2.Checked = true;
        }
    }
}
