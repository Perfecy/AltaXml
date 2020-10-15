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


namespace AltaXML
{
    public partial class Form1 : Form
    {
        private string name;
        private string file;
        private string new_file_name;
      
        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                name = openFileDialog1.FileName;
                Debug.WriteLine(name);
                file = File.ReadAllText(name);
                Debug.WriteLine(file);
                new_file_name = "C:/Users/Kirik/Downloads/1.xml" ;
                File.Create(new_file_name).Close();
                File.WriteAllText(new_file_name, file);


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
