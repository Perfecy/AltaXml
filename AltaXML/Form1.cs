using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;


namespace AltaXML
{
    public partial class Form1 : Form
    {
        private string file_name;
        private string file;
        private string new_file_name;
        private string template_file_name;
        private string directory_root_name;
        
        class AltaIndPost
        {
            public string NUM;
            public string INVNUM;
            public string INVDATE;
        }

        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            DirectoryInfo directory_root = Directory.GetParent(Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName);//получаем корнейвую папку
            directory_root_name = directory_root.FullName;//получаем имя корневой папки
            //Debug.WriteLine(directory_root );
            foreach(FileInfo fi in directory_root.GetFiles())
            {
                //Debug.WriteLine(fi.FullName);
                if (fi.FullName.Contains("template.xml")) { template_file_name = fi.FullName; }
                
            }
            Debug.WriteLine(template_file_name);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                file_name = openFileDialog1.FileName;
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(template_file_name);
                // получим корневой элемент
                XmlElement root = xDoc.DocumentElement;
                Debug.WriteLine("имя рута "+root.Name);
                foreach(XmlNode node in root)
                {
                    Debug.WriteLine("имя ноды "+node.Name + " значение = "+ node.InnerText);
                    if (node.Name == "NUM")
                    {
                        node.InnerText = "Sosi zhopu";
                    }
                    
                }
                xDoc.Save(file_name);

            }
        }

        private List<string> ParseExcel(string file)
        {
            List<string> data = new List<string>();
            return data;
        }
         
        private void InsertDataToXML(List<string> data)
        {

        }

    }

}
