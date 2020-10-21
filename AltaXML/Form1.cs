﻿using Microsoft.CSharp.RuntimeBinder;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace AltaXML
{
    public partial class Form1 : Form
    {
        private string file_name;

        private string template_file_name;
        private string directory_root_name;

        private XmlElement root;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;
        private Excel.Range range;
        private Excel.Application stApp;
        private Excel.Workbook stWorkBook;
        private Excel.Worksheet stWorkSheet;
        private Excel.Range stRange;
        private int counter = 0;
        private int error_counter = 0;
        private bool full = false;
        private bool additive_data = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_name = openFileDialog1.FileName;
                
                FileDisplay.Text = file_name;

                if (!file_name.Contains(".xls"))
                {
                    MessageBox.Show("Ошибка: выбранный файл не является файлом Excel.", "Формат файла");
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

        private void button3_Click(object sender, EventArgs e)//import
        {
            DialogResult dialogResult = MessageBox.Show("Загрузить файл с дополнительными данными?", "Дополнительные данные", MessageBoxButtons.YesNo);

            List<string> missed_values = new List<string>();
            Excel.Workbooks stworkbooks = null;
            Excel.Sheets stworksheet = null;

            if (dialogResult == DialogResult.Yes)
            {
                if(openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    additive_data = true;
                    stApp = new Excel.Application();
                    stworkbooks = stApp.Workbooks;
                    stWorkBook = stworkbooks.Open(openFileDialog2.FileName);
                    stworksheet = stWorkBook.Worksheets;
                    stWorkSheet = (Excel.Worksheet)stworksheet.get_Item(1);
                    stRange = stWorkSheet.UsedRange;

                    for (int i = 1; i <= stRange.Columns.Count; i++)
                    {
                        missed_values.Add(Convert.ToString((stRange.Cells[2, i] as Excel.Range).Value));
                    }

                }
                

            }
            else if (dialogResult == DialogResult.No)
            {
                additive_data = false;
            }


            DirectoryInfo directory_root = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName);//получаем корнейвую папку
            directory_root_name = directory_root.FullName;//получаем имя корневой папки
                                                            //Debug.WriteLine(directory_root );
            foreach (FileInfo fi in directory_root.GetFiles())
            {
                //Debug.WriteLine(fi.FullName);
                if (fi.FullName.Contains("import.xml")) { template_file_name = fi.FullName; }//получаем путь к шаблону
            }
            //
            xlApp = new Excel.Application();

            try
            {
                
                var workbooks = xlApp.Workbooks;
                xlWorkBook = workbooks.Open(file_name);
                
                var worksheet = xlWorkBook.Worksheets;
                
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet = (Excel.Worksheet)worksheet.get_Item(1);
                
                FileDisplay.Text = file_name;

                int rw = 0;
                int cl = 0;

                range = xlWorkSheet.UsedRange;
                
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                List<string> fnames = new List<string>();
                List<string> cell_names = new List<string>();
                


                for (int i = 1; i <= cl; i++)
                {
                    cell_names.Add((string)(range.Cells[1, i] as Excel.Range).Value);
                }

                if (cell_names.Contains("inn"))
                {
                    full = true;
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


                        if (Convert.ToString((range.Cells[j, 1] as Excel.Range).Value) != null)
                        {
                            counter += 1;
                            for (int i = 1; i <= cl; i++)
                            {
                                cell_values.Add((Convert.ToString((range.Cells[j, i] as Excel.Range).Value))); //.Replace(" ", String.Empty));
                            }

                            foreach (XmlNode node in root)
                            {
                                if (full)
                                {
                                    if (node.Name == "IDENTITYCARDSERIES")
                                    {
                                        node.InnerText = cell_values[16];
                                    }
                                    if (node.Name == "IDENTITYCARDNUMBER")
                                    {
                                        node.InnerText = cell_values[17];
                                    }
                                    if (node.Name == "ORGANIZATIONNAME")
                                    {
                                        node.InnerText = cell_values[18];
                                    }
                                    if (node.Name == "IDENTITYCARDDATE")
                                    {
                                        node.InnerText = cell_values[19];
                                    }
                                }

                                if (node.Name == "NUM")
                                {
                                    node.InnerText = cell_values[0];
                                }
                                if (node.Name == "INVNUM")
                                {
                                    node.InnerText = cell_values[0];
                                }
                                if (node.Name == "SENDER")
                                {
                                    node.InnerText = cell_values[1];
                                }
                                if (node.Name == "PERSONSURNAME")
                                {
                                    node.InnerText = cell_values[2];
                                }
                                if (node.Name == "PERSONNAME")
                                {
                                    node.InnerText = cell_values[3];
                                }
                                if (node.Name == "PERSONMIDDLENAME")
                                {
                                    node.InnerText = cell_values[4];
                                }
                                if (node.Name == "PHONE")
                                {
                                    node.InnerText = cell_values[5];
                                }
                                if (node.Name == "PHONEMOB")
                                {
                                    node.InnerText = cell_values[5];
                                }
                                if (node.Name == "EMAIL")
                                {
                                    node.InnerText = cell_values[6];
                                }
                                if (node.Name == "CITY")
                                {
                                    node.InnerText = cell_values[8];
                                }
                                if (node.Name == "City")
                                {
                                    node.InnerText = cell_values[8];
                                }
                                if (node.Name == "POSTALCODE")
                                {
                                    node.InnerText = cell_values[7];
                                }
                                if (node.Name == "STREETHOUSE")
                                {
                                    node.InnerText = cell_values[9];
                                }
                                if (node.Name == "StreetHouse")
                                {
                                    node.InnerText = cell_values[9];
                                }
                                if (node.Name == "GOODS")
                                {
                                    foreach (XmlNode child in node.ChildNodes)
                                    {
                                        if (child.Name == "DESCR")
                                        {
                                            child.InnerText = cell_values[10];
                                        }
                                        if (child.Name == "TNVED")
                                        {
                                            if (full)
                                            {
                                                child.InnerText = cell_values[21];
                                            }
                                            else
                                            {
                                                child.InnerText = cell_values[16];
                                            }

                                        }
                                        if (child.Name == "PRICE")
                                        {
                                            child.InnerText = cell_values[11];
                                        }
                                        if (child.Name == "ORGWEIGHT")
                                        {
                                            child.InnerText = cell_values[13];
                                        }
                                        if (child.Name == "NETTO")
                                        {
                                            child.InnerText = cell_values[13];
                                        }
                                        if (child.Name == "WEIGHT")
                                        {
                                            child.InnerText = Convert.ToString(float.Parse(cell_values[13]) * float.Parse(cell_values[14]));
                                        }
                                        if (child.Name == "QTY")
                                        {
                                            child.InnerText = cell_values[14];
                                        }
                                        if (child.Name == "URL")
                                        {
                                            child.InnerText = cell_values[15];

                                        }
                                        if (child.Name == "COST")
                                        {
                                            child.InnerText = Convert.ToString(float.Parse(cell_values[11]) * float.Parse(cell_values[14]));
                                        }
                                        if (additive_data)
                                        {
                                            if (child.Name == "MEASUREUNITQUALIFIERCODE")
                                            {
                                                child.InnerText = missed_values[15];
                                            }
                                            if (child.Name == "MEASUREUNITQUALIFIERNAME")
                                            {
                                                child.InnerText = missed_values[16];
                                            }
                                        }
                                        
                                    }
                                }
                                if (node.Name == "CURRENCY")
                                {
                                    node.InnerText = cell_values[12];
                                }

                                if (node.Name == "CONSIGNOR_IDENTITYCARD_ORGANIZATIONNAME")
                                {
                                    node.InnerText = cell_values[1];
                                }
                                if (node.Name == "CONSIGNOR_RFORGANIZATIONFEATURES_INN")
                                {
                                    node.InnerText = cell_values[20];
                                }
                                if (node.Name == "RFORGANIZATIONFEATURES_INN")
                                {
                                    node.InnerText = cell_values[20];
                                }
                                //-----------------------------

                                if (additive_data)
                                {
                                    if (node.Name == "AVIANUM")
                                    {
                                        node.InnerText = missed_values[0];
                                    }
                                    if (node.Name == "IDENTITYCARDNAME")
                                    {
                                        node.InnerText = missed_values[1];
                                    }
                                    if (node.Name == "IDENTITYCARDCODE")
                                    {
                                        node.InnerText = missed_values[2];
                                    }
                                    if (node.Name == "CONSIGNEE_ADDRESS_COUNTRYCODE")
                                    {
                                        node.InnerText = missed_values[3];
                                    }
                                    if (node.Name == "CONSIGNEE_ADDRESS_COUNTRYNAME")
                                    {
                                        node.InnerText = missed_values[4];
                                    }
                                    if (node.Name == "ORGCOUNTRY")
                                    {
                                        node.InnerText = missed_values[5];
                                    }
                                    if (node.Name == "DELIVERYTERMS_TRADINGCOUNTRYCODE")
                                    {
                                        node.InnerText = missed_values[6];
                                    }
                                    if (node.Name == "DELIVERYTERMS_DISPATCHCOUNTRYCODE")
                                    {
                                        node.InnerText = missed_values[7];
                                    }
                                    if (node.Name == "DELIVERYTERMS_DELIVERYTERMSSTRINGCODE")
                                    {
                                        node.InnerText = missed_values[8];
                                    }
                                    if (node.Name == "DEPARTUREPOINT_IATACODE")
                                    {
                                        node.InnerText = missed_values[9];
                                    }
                                    if (node.Name == "DELIVERYPOINT_IATACODE")
                                    {
                                        node.InnerText = missed_values[10];
                                    }
                                    if (node.Name == "ARRIVEDATE")
                                    {
                                        node.InnerText = missed_values[11];
                                    }
                                    if (node.Name == "ALLCOST")
                                    {
                                        node.InnerText = missed_values[12];
                                    }
                                    if (node.Name == "ALLWEIGHT")
                                    {
                                        node.InnerText = missed_values[13];
                                    }
                                    if (node.Name == "PLACES")
                                    {
                                        node.InnerText = missed_values[14];
                                    }
                                    
                                }

                            }

                            if (fnames.Contains(cell_values[0]))
                            {

                                xDoc.Save(folderBrowserDialog1.SelectedPath + "\\" + cell_values[0] + "-" + fnames.Count<string>(p => p == cell_values[0]) + ".xml");
                            }
                            else
                            {
                                xDoc.Save(folderBrowserDialog1.SelectedPath + "\\" + cell_values[0] + ".xml");
                            }
                            if (fnames.Contains(cell_values[0]))
                            {
                                ProcessDisplay.AppendText(cell_values[0] + "-" + fnames.Count<string>(p => p == cell_values[0]) + "\r\n");
                            }
                            else
                            {
                                ProcessDisplay.AppendText(cell_values[0] + "\r\n");
                            }
                            fnames.Add(cell_values[0]);
                            cell_values.Clear();

                        }
                    }
                    ProcessDisplay.AppendText("Обработано записей: " + counter + "\r\n" + "Обработка завершена.\n");
                    full = false;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlWorkBook.Close(0);
                    //xlApp.Quit();
                    if (additive_data)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stworksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stworkbooks);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stApp);
                        stWorkBook.Close(0);
                    }
                    
                    
                    // stApp.Quit();
                }
            }
            catch (COMException comex)
            {
                MessageBox.Show("Пожалуйста, выберите файл Excel", "Выбор файла");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            DirectoryInfo directory_root = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName);//получаем корнейвую папку
            directory_root_name = directory_root.FullName;//получаем имя корневой папки

            DialogResult dialogResult = MessageBox.Show("Загрузить файл с дополнительными данными?", "Дополнительные данные", MessageBoxButtons.YesNo);

            List<string> missed_values = new List<string>();
            Excel.Workbooks stworkbooks = null;
            Excel.Sheets stworksheet = null;

            if (dialogResult == DialogResult.Yes)
            {
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    additive_data = true;
                    stApp = new Excel.Application();
                    stworkbooks = stApp.Workbooks;
                    stWorkBook = stworkbooks.Open(openFileDialog2.FileName);
                    stworksheet = stWorkBook.Worksheets;
                    stWorkSheet = (Excel.Worksheet)stworksheet.get_Item(1);
                    stRange = stWorkSheet.UsedRange;

                    for (int i = 1; i <= stRange.Columns.Count; i++)
                    {
                        missed_values.Add(Convert.ToString((stRange.Cells[2, i] as Excel.Range).Value));
                    }

                }


            }
            else if (dialogResult == DialogResult.No)
            {
                additive_data = false;
            }


            foreach (FileInfo fi in directory_root.GetFiles())
            {
                //Debug.WriteLine(fi.FullName);
                if (fi.FullName.Contains("export.xml")) { template_file_name = fi.FullName; }//получаем путь к шаблону
            }
            xlApp = new Excel.Application();

            try
            {
                var workbooks = xlApp.Workbooks;
                //xlWorkBook = xlApp.Workbooks.Open(file_name);
                xlWorkBook = workbooks.Open(file_name);
                var worksheet = xlWorkBook.Worksheets;
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet = (Excel.Worksheet)worksheet.get_Item(1);
                FileDisplay.Text = file_name;

                int rw = 0;
                int cl = 0;

                range = xlWorkSheet.UsedRange;

                rw = range.Rows.Count;
                cl = range.Columns.Count;

                List<string> fnames = new List<string>();
                List<string> cell_names = new List<string>();

                for (int i = 1; i <= cl; i++)
                {
                    cell_names.Add((string)(range.Cells[1, i] as Excel.Range).Value);
                }

                if (cell_names.Contains("inn"))
                {
                    full = true;
                    Debug.WriteLine("KAK");
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

                        try
                        {
                            if (Convert.ToString((range.Cells[j, 1] as Excel.Range).Value) != null)
                            {
                                counter += 1;
                                for (int i = 1; i <= cl; i++)
                                {
                                    cell_values.Add((Convert.ToString((range.Cells[j, i] as Excel.Range).Value)));  //.Replace(" ", String.Empty));
                                }

                                foreach (XmlNode node in root)
                                {
                                    
                                        if (node.Name == "IDENTITYCARDSERIES")
                                        {
                                            node.InnerText = cell_values[16];
                                        }
                                        if (node.Name == "IDENTITYCARDNUMBER")
                                        {
                                            node.InnerText = cell_values[17];
                                        }
                                        if (node.Name == "ORGANIZATIONNAME")
                                        {
                                            node.InnerText = cell_values[18];
                                        }
                                        if (node.Name == "IDENTITYCARDDATE")
                                        {
                                            node.InnerText = cell_values[19];
                                        }
                                        
                                    if (node.Name == "NUM")
                                    {
                                        node.InnerText = cell_values[0];
                                    }
                                    if (node.Name == "INVNUM")
                                    {
                                        node.InnerText = cell_values[0];
                                    }
                                    if (node.Name == "SENDER")
                                    {
                                        node.InnerText = cell_values[1];
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
                                    if (node.Name == "PHONE")
                                    {
                                        node.InnerText = cell_values[4];
                                    }
                                    if (node.Name == "PHONEMOB")
                                    {
                                        node.InnerText = cell_values[4];
                                    }
                                    if (node.Name == "EMAIL")
                                    {
                                        node.InnerText = cell_values[5];
                                    }
                                    if (node.Name == "CITY")
                                    {
                                        node.InnerText = cell_values[7];
                                    }
                                    if (node.Name == "City")
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
                                    if (node.Name == "StreetHouse")
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
                                                child.InnerText = cell_names[10];
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
                                                child.InnerText = Convert.ToString(float.Parse(cell_values[13]) * float.Parse(cell_values[14]));
                                            }
                                            if (child.Name == "QTY")
                                            {
                                                child.InnerText = cell_values[14];
                                            }
                                            if (child.Name == "URL")
                                            {
                                                child.InnerText = cell_values[15];
                                            }
                                            if (child.Name == "COST")
                                            {
                                                child.InnerText = Convert.ToString(float.Parse(cell_values[11]) * float.Parse(cell_values[14]));
                                            }
                                        }
                                    }
                                    if (node.Name == "CURRENCY")
                                    {
                                        node.InnerText = cell_values[12];
                                    }

                                    if (node.Name == "CONSIGNOR_IDENTITYCARD_ORGANIZATIONNAME")
                                    {
                                        node.InnerText = cell_values[18];
                                    }
                                    if (node.Name == "CONSIGNOR_RFORGANIZATIONFEATURES_INN")
                                    {
                                        node.InnerText = cell_values[20];
                                    }
                                    if (node.Name == "RFORGANIZATIONFEATURES_INN")
                                    {
                                        node.InnerText = cell_values[20];
                                    }
                                    //---------------------------------

                                    if (additive_data)
                                    {
                                        if (node.Name == "INVDATE")
                                        {
                                            node.InnerText = missed_values[0];
                                        }
                                        if (node.Name == "AVIANUM")
                                        {
                                            node.InnerText = missed_values[1];
                                        }
                                        if (node.Name == "AVIADATE")
                                        {
                                            node.InnerText = missed_values[2];
                                        }
                                        if (node.Name == "TYPE")
                                        {
                                            node.InnerText = missed_values[3];
                                        }
                                        if (node.Name == "CONSIGNEE_CHOICE")
                                        {
                                            node.InnerText = missed_values[4];
                                        }
                                        if (node.Name == "CONSIGNEE_ADDRESS_COUNTRYCODE")
                                        {
                                            node.InnerText = missed_values[5];
                                        }
                                        if (node.Name == "CONSIGNEE_ADDRESS_COUNRYNAME")
                                        {
                                            node.InnerText = missed_values[6];
                                        }
                                        if (node.Name == "IDENTITYCARDNAME")
                                        {
                                            node.InnerText = missed_values[7];
                                        }
                                        if (node.Name == "IDENTITYCARDCODE")
                                        {
                                            node.InnerText = missed_values[8];
                                        }
                                        if (node.Name == "CONSIGNOR_CHOICE")
                                        {
                                            node.InnerText = missed_values[9];
                                        }
                                        if (node.Name == "CONSIGNOR_ADDRESS_POSTALCODE")
                                        {
                                            node.InnerText = missed_values[10];
                                        }
                                        if (node.Name == "CONSIGNOR_ADDRESS_CITY")
                                        {
                                            node.InnerText = missed_values[11];
                                        }
                                        if (node.Name == "CONSIGNOR_ADDRESS_STREETHOUSE")
                                        {
                                            node.InnerText = missed_values[12];
                                        }
                                        if (node.Name == "COUNTRYCODE")
                                        {
                                            node.InnerText = missed_values[13];
                                        }
                                        if (node.Name == "COUNTRYNAME")
                                        {
                                            node.InnerText = missed_values[14];
                                        }
                                        if (node.Name == "ORGCOUNTRY")
                                        {
                                            node.InnerText = missed_values[15];
                                        }
                                        if (node.Name == "DELIVERYTERMS_TRADINGCOUNTRYCODE")
                                        {
                                            node.InnerText = missed_values[16];
                                        }
                                        if (node.Name == "DELIVERYTERMS_DISPATCHCOUNTRYCODE")
                                        {
                                            node.InnerText = missed_values[17];
                                        }
                                        if (node.Name == "DELIVERYTERMS_DELIVERYTERMSSTRINGCODE")
                                        {
                                            node.InnerText = missed_values[18];
                                        }
                                        if (node.Name == "PLACES")
                                        {
                                            node.InnerText = missed_values[19];
                                        }
                                        if (node.Name == "ALLCOST")
                                        {
                                            node.InnerText = Convert.ToString(float.Parse(cell_values[11]) * float.Parse(cell_values[14])).Replace('.', ',');
                                        }
                                        if (node.Name == "ALLWEIGHT")
                                        {
                                            node.InnerText = Convert.ToString(float.Parse(cell_values[13]) * float.Parse(cell_values[14])).Replace('.', ',');
                                        }
                                    }
                                }

                                if (fnames.Contains(cell_values[0]))
                                {

                                    xDoc.Save(folderBrowserDialog1.SelectedPath + "\\" + cell_values[0] + "-" + fnames.Count<string>(p => p == cell_values[0]) + ".xml");
                                }
                                else
                                {
                                    xDoc.Save(folderBrowserDialog1.SelectedPath + "\\" + cell_values[0] + ".xml");
                                }

                                if (fnames.Contains(cell_values[0]))
                                {
                                    ProcessDisplay.AppendText(cell_values[0] + "-" + fnames.Count<string>(p => p == cell_values[0]) + "\r\n");
                                }
                                else
                                {
                                    ProcessDisplay.AppendText(cell_values[0] + "\r\n");
                                }
                                fnames.Add(cell_values[0]);
                                cell_values.Clear();

                            }
                        }
                        catch (RuntimeBinderException ex)
                        {
                            ProcessDisplay.AppendText("Ошибка чтения строки \r\n");
                            error_counter += 1;
                        }
                    }
                    ProcessDisplay.AppendText("Обработано записей: " + counter + "\r\n" + "Ошибок чтения: " + error_counter + "\r\n" + "Обработка завершена. \r\n");
                    counter = 0;
                    error_counter = 0;
                    full = false;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlWorkBook.Close(0);
                    if (additive_data)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stworksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stworkbooks);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(stApp);
                        stWorkBook.Close(0);
                    }
                        
                    //xlApp.Quit();
                    //stApp.Quit();
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show("Пожалуйста, выберите файл Excel", "Выбор файла");
                    
            }

        }
    }
}