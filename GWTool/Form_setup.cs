using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace GWTool
{
    public partial class Form_setup : Form
    {
        public Form_setup()
        {
            InitializeComponent();
            DataInit();
        }

        private void DataInit()
        {
            string path = Path.Combine(Application.StartupPath, "config");
            string filePath_b = Path.Combine(Application.StartupPath, "config\\config_base.xml");
            string filePath_y = Path.Combine(Application.StartupPath, "config\\config_y.xml");
            //MessageBox.Show(path,"xml");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                if(!File.Exists(filePath_b))
                    CreateNewXml_b(filePath_b);
            }
            else
            {
                if (!File.Exists(filePath_b))
                    CreateNewXml_b(filePath_b);
            }
            
        }

        /**
         * 新建基本设置的xml文件
         * */
        private void CreateNewXml_b(string filePath_b)
        {
            //通过代码创建XML文档
            //1、引用命名空间   System.Xml
            //2、创建一个 xml 文档
            XmlDocument xml = new XmlDocument();
            //3、创建一行声明信息，并添加到 xml 文档顶部
            XmlDeclaration decl = xml.CreateXmlDeclaration("1.0", "utf-8", null);
            xml.AppendChild(decl);

            //4、创建根节点
            XmlElement rootEle = xml.CreateElement("config");
            xml.AppendChild(rootEle);
            //5、创建子结点|属性：发文单位数据
            XmlElement childEle = xml.CreateElement("fwdw");
            rootEle.AppendChild(childEle);

            XmlElement c2Ele = xml.CreateElement("fwdwCount");
            c2Ele.InnerText = "0";
            childEle.AppendChild(c2Ele);

            //6、创建子节点|属性：发文字号数据
            childEle = xml.CreateElement("fwzh");
            rootEle.AppendChild(childEle);

            c2Ele = xml.CreateElement("fwzhCount");
            c2Ele.InnerText = "0";
            childEle.AppendChild(c2Ele);

            //7、创建子节点|属性：承办单位数据
            childEle = xml.CreateElement("cbdw");
            rootEle.AppendChild(childEle);

            c2Ele = xml.CreateElement("cbdwCount");
            c2Ele.InnerText = "0";
            childEle.AppendChild(c2Ele);

            xml.Save(filePath_b);
        }

        /**
         * 用新配置覆盖旧配置
         * */
        private void OverwriteXml_b()
        { }

        private void CreateNode(XmlDocument xmlDoc, XmlNode parentNode, string name, string value)
        {
            XmlNode node = xmlDoc.CreateNode(XmlNodeType.Element, name, null);
            node.InnerText = value;
            parentNode.AppendChild(node);
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
