﻿using System;
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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace AltaXML
{
    public partial class Form1 : Form
    {
        private string file_name;

        private string template_file_name;
        private string directory_root_name;

        private XmlElement root;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;


        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            DirectoryInfo directory_root = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName);//получаем корнейвую папку
            directory_root_name = directory_root.FullName;//получаем имя корневой папки
            //Debug.WriteLine(directory_root );
            foreach(FileInfo fi in directory_root.GetFiles())
            {
                //Debug.WriteLine(fi.FullName);
                if (fi.FullName.Contains("template.xml")) { template_file_name = fi.FullName; }//получаем путь к шаблону
                
            }
            Debug.WriteLine(template_file_name);
            //FileDisplay.Text = template_file_name;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                file_name = openFileDialog1.FileName;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file_name);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                FileDisplay.Text = file_name;


                int rw = 0;
                int cl = 0;

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                List<string> cell_names = new List<string>();
          
                for (int i = 1; i <= cl; i++)
                {
                    cell_names.Add((string)(range.Cells[1, i] as Excel.Range).Value);
                }

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(template_file_name);
                // получим корневой элемент

                //   Debug.WriteLine("имя рута " + root.Name);
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {



                    List<string> cell_values = new List<string>();

                    for (int j = 2; j <= rw; j++)
                    {
                        xDoc.Load(template_file_name);
                        root = xDoc.DocumentElement;
                        root.SetAttribute("time", DateTime.Now.ToString("yyyy-mm-dd"));
                        for (int i = 1; i <= cl; i++)
                        {
                            cell_values.Add(Convert.ToString((range.Cells[j, i] as Excel.Range).Value));
                            //Debug.WriteLine(cell_values[i-1]);
                        }

                        foreach (XmlNode node in root)
                        {

                            if (node.Name == "NUM")
                            {
                                node.InnerText = cell_values[0];
                            }
                            if (node.Name == "INVNUM")
                            {
                                node.InnerText = cell_values[0];
                            }
                            if (node.Name == "INVDATE")
                            {
                                node.InnerText = DateTime.Now.ToString("yyyy-mm-dd");
                            }
                            if (node.Name == "PERSONSURNAME")
                            {
                                node.InnerText = cell_values[1];
                            }
                            if (node.Name == "PERSONNAME")
                            {
                                node.InnerText = cell_values[2];
                            }
                            if (node.Name == "PERSONMIDDLENAME")
                            {
                                node.InnerText = cell_values[3];
                            }
                            if (node.Name == "CITY")
                            {
                                node.InnerText = cell_values[7];
                            }
                            if (node.Name == "POSTALCODE")
                            {
                                node.InnerText = cell_values[6];
                            }
                            if (node.Name == "STREETHOUSE")
                            {
                                node.InnerText = cell_values[8];
                            }
                            if (node.Name == "GOODS")
                            {
                                foreach (XmlNode child in node.ChildNodes)
                                {
                                    if (child.Name == "DESCR")
                                    {
                                        child.InnerText = cell_values[9];
                                    }
                                    if (child.Name == "TNVED")
                                    {
                                        child.InnerText = cell_values[10];
                                    }
                                    if (child.Name == "PRICE")
                                    {
                                        child.InnerText = cell_values[11];
                                    }
                                    if (child.Name == "ORGWEIGHT")
                                    {
                                        child.InnerText = cell_values[13];
                                    }
                                    if (child.Name == "WEIGHT")
                                    {
                                        child.InnerText = cell_values[13];
                                    }
                                    if (child.Name == "QTY")
                                    {
                                        child.InnerText = cell_values[14];
                                    }
                                }
                            }
                            if (node.Name == "CURRENCY")
                            {
                                node.InnerText = cell_values[12];
                            }

                            if (node.Name == "IDENTITYCARDNUMBER")
                            {
                                node.InnerText = cell_values[17];
                            }
                            if (node.Name == "CONSIGNOR_IDENTITYCARD_ORGANIZATIONNAME")
                            {
                                node.InnerText = cell_values[18];
                            }
                            if (node.Name == "CONSIGNOR_RFORGANIZATIONFEATURES_INN")
                            {
                                node.InnerText = cell_values[20];
                            }
                            if (node.Name == "IDENTITYCARDSERIES")
                            {
                                node.InnerText = cell_values[16];
                            }


                        }
                        ProcessDisplay.AppendText(cell_values[0] + "\r\n");
                        xDoc.Save(folderBrowserDialog1.SelectedPath + "\\"+ cell_values[0] + ".xml");
                        cell_values.Clear();
                        
                    }
                    ProcessDisplay.AppendText("Обработано записей: " + rw + "\r\n" + "Обработка завершена.");
                }
            }
        }

   

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }
    }

}
