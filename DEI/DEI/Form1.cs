using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Win32;
using Microsoft.CSharp;
using excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace DEI
{
    public partial class Form1 : Form
    {
         DataTable dt2 = new DataTable();
         DataTable dt1 = new DataTable();
         DataTable dt3 = new DataTable();
         DataColumn dc = new DataColumn();

        public int first_Time = 0;

        public Form1()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog DialogA = new OpenFileDialog();
            DialogA.CheckFileExists = true;
            DialogA.Title = "Select a File";

            DialogA.ShowDialog();
            textBox1.Text = DialogA.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog DialogB = new OpenFileDialog();
            DialogB.CheckFileExists = true;
            DialogB.Title = "Select a File";

            DialogB.ShowDialog();
            textBox3.Text = DialogB.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(stringconn);
            if (textBox1.Text != "" || textBox2.Text != "")
            {
                    try
                {
                    OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                    //  DataTable dt1 = new DataTable();

                    da.Fill(dt1);

                    dataGridView1.DataSource = dt1;
                    load1.Text = "Number of Loaded Objects = " + dt1.Rows.Count.ToString();
                }
           catch(Exception ex)
                {
                    if (true)
                        MessageBox.Show("ER" + ex);
                }

            }
            else
                MessageBox.Show("ER : You must Enter file path and Correct Sheet name");
            }

        private void button4_Click(object sender, EventArgs e)
        {
            
               string stringconn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox3.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
                OleDbConnection conn2 = new OleDbConnection(stringconn2);


            if (textBox3.Text != "" || textBox4.Text != "")
                    {

                try
                {
                    OleDbDataAdapter da2 = new OleDbDataAdapter("Select * from [" + textBox4.Text + "$]", conn2);


                    da2.Fill(dt2);



                    dataGridView2.DataSource = dt2;
                    test123.Text = "Number of Loaded Objects = " + dt2.Rows.Count.ToString();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("ER : You must Enter file path and Correct Sheet name" + ex);
                }
              
          }
             else
                  MessageBox.Show("ER : You must Enter file path and Correct Sheet name");
      
                }

        private void Same_As_Click(object sender, EventArgs e)
        {
            


                if (first_Time == 0)
                {
                    dc = new DataColumn("Sub1", typeof(String));
                    dt3.Columns.Add(dc);
                    dc = new DataColumn("Lat1", typeof(String));
                    dt3.Columns.Add(dc);
                    dc = new DataColumn("Long1", typeof(String));
                    dt3.Columns.Add(dc);
                    dc = new DataColumn("Sub2", typeof(String));
                    dt3.Columns.Add(dc);
                    dc = new DataColumn("Lat2", typeof(String));
                    dt3.Columns.Add(dc);
                    dc = new DataColumn("Long2", typeof(String));
                    dt3.Columns.Add(dc);

                }
                first_Time = 1;

                DataRow dr = dt3.NewRow();
            float TBV;

            if (float.TryParse(distance.Text, out TBV))
            {
                // success! Use f here
            }
            else
            {
                TBV = 1000;
                MessageBox.Show("ER : you Must Enter Number in Distance Text Box");

            }
            
                int obj1_Col = Convert.ToInt32(ob1.Text);
                int obj2_Col = Convert.ToInt32(ob2.Text);
                int lat1_Col = Convert.ToInt32(lat1.Text);
                int lat2_Col = Convert.ToInt32(lat2.Text);
                int long1_Col = Convert.ToInt32(long1.Text);
                int long2_Col = Convert.ToInt32(long2.Text);

                float dist_L = (TBV / 11) / 10000;


                test.Text = "Max Displacement Distance = " + TBV + "  Max Displacement Degrees  = " + dist_L;

                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    for (int c = 0; c < dt2.Rows.Count; c++)
                    {

                        string obj1_Value = dt1.Rows[i][obj1_Col].ToString();
                        float lat1_Value = float.Parse(dt1.Rows[i][lat1_Col].ToString());
                        float long1_Value = float.Parse(dt1.Rows[i][long1_Col].ToString());

                        string obj2_Value = dt2.Rows[c][obj2_Col].ToString();
                        float lat2_Value = float.Parse(dt2.Rows[c][lat2_Col].ToString());
                        float long2_Value = float.Parse(dt2.Rows[c][long2_Col].ToString());



                        if (Math.Abs(lat1_Value - lat2_Value) <= dist_L && Math.Abs(long1_Value - long2_Value) <= dist_L)
                        {

                            dt3.Rows.Add(new Object[]{obj1_Value , lat1_Value , long1_Value , obj2_Value , lat2_Value , long2_Value

                     });

                        }

                    }
                    dataGridView3.DataSource = dt3;
                lastresult.Text = "Number Of Same-as Subjects =" + dt3.Rows.Count.ToString();
                }
           
          
        }

    
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
         //   label2.Text = trackBar1.Value.ToString();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                ((DataTable)dataGridView3.DataSource).Rows.Clear();
            }
           catch (Exception ex)
            {

            }         
        }

        private void button6_Click(object sender, EventArgs e)
        {
         
            string file = @"D:\SameAsOutput.txt";
            int ds_r = dt3.Rows.Count;
            String[] lines = new String[ds_r];
           

            for (int i = 0; i < dt3.Rows.Count; i++)

            {
                lines[i] =   "<rdf:Description rdf:about=\"" + dt3.Rows[i][0].ToString()  + "\"> \n <owl:sameAs rdf:resource =\"" + dt3.Rows[i][3].ToString() + "\"/> \n </rdf:Description>"  ;
            }
              
            File.WriteAllLines(file, lines);
          
        }

        private void lat2_SelectedItemChanged(object sender, EventArgs e)
        {

        }
    }
    }

