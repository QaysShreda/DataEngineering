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
            try
            {
                ((DataTable)dataGridView1.DataSource).Rows.Clear();
            }
            catch (Exception ex)
            {

            }



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
                catch (Exception ex)
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
            try
            {
                ((DataTable)dataGridView2.DataSource).Rows.Clear();
            }
            catch (Exception ex)
            {

            }
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
                catch (Exception ex)
                {
                    MessageBox.Show("ER : You must Enter file path and Correct Sheet name" + ex);
                }

            }
            else
                MessageBox.Show("ER : You must Enter file path and Correct Sheet name");

        }

        private void Same_As_Click(object sender, EventArgs e)
        {



            //    test_Lable.Text = "";
            try
            {
                ((DataTable)dataGridView3.DataSource).Rows.Clear();
            }
            catch (Exception ex)
            {

            }


            if (first_Time == 0)
            {
                dc = new DataColumn("Sub1", typeof(String));
                dt3.Columns.Add(dc);
                dc = new DataColumn("Lat1", typeof(String));
                dt3.Columns.Add(dc);
                dc = new DataColumn("Long1", typeof(String));
                dt3.Columns.Add(dc);
                dc = new DataColumn("Object1", typeof(String));
                dt3.Columns.Add(dc);
                dc = new DataColumn("Object2", typeof(String));
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
                MessageBox.Show("ER : You Must Enter Number in Distance Text Box");

            }

            int sub1_Col = Convert.ToInt32(sub1.Text);
            int sub2_Col = Convert.ToInt32(sub2.Text);

            int lat1_Col = Convert.ToInt32(lat1.Text);
            int lat2_Col = Convert.ToInt32(lat2.Text);

            int long1_Col = Convert.ToInt32(long1.Text);
            int long2_Col = Convert.ToInt32(long2.Text);

            int string1_Col = Convert.ToInt32(str1.Text);
            int string2_Col = Convert.ToInt32(str2.Text);


            int req_Distance;
            int actual_distance;
            int req_Similarity;
            int actual_Similarity;
            //double actual_JaroSimilarity;
            double percent;




            float dist_L = (TBV / 11) / 10000;


            test.Text = "Max Displacement Distance = " + TBV + "  Max Displacement Degrees  = " + dist_L;

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                for (int c = 0; c < dt2.Rows.Count; c++)
                {

                    string sub1_Value = dt1.Rows[i][sub1_Col].ToString();
                    string string1_Value = dt1.Rows[i][string1_Col].ToString();
                    //   float long1_Value = float.Parse(dt1.Rows[i][long1_Col].ToString());

                    string sub2_Value = dt2.Rows[c][sub2_Col].ToString();
                    string string2_Value = dt2.Rows[c][string2_Col].ToString();
                    float lat1_Value;
                    float long1_Value;
                    if (LC.Checked)
                    {
                        string1_Value.ToLower();
                        string2_Value.ToLower();

                    }

                    if (float.TryParse(dt1.Rows[i][lat1_Col].ToString(), out lat1_Value))
                    {
                        // success! Use f here
                    }
                    if (float.TryParse(dt1.Rows[i][long1_Col].ToString(), out long1_Value))
                    {
                        // success! Use f here
                    }


                    float lat2_Value;
                    float long2_Value;


                    if (float.TryParse(dt2.Rows[c][lat2_Col].ToString(), out lat2_Value))
                    {
                        // success! Use f here
                    }
                    if (float.TryParse(dt2.Rows[c][long2_Col].ToString(), out long2_Value))
                    {

                    }
                    else
                    {
                        long2_Value = 0;
                    }


                    if (radioButton1.Checked == true)
                    {
                        req_Distance = Convert.ToInt32(trackBar1.Value.ToString());
                        distance_Algorithm dis_Al = new distance_Algorithm();
                        actual_distance = dis_Al.Compute_distance(string1_Value, string2_Value);

                        if (Math.Abs(lat1_Value - lat2_Value) <= dist_L && Math.Abs(long1_Value - long2_Value) <= dist_L && actual_distance <= req_Distance)
                        {

                            dt3.Rows.Add(new Object[] { sub1_Value, lat1_Value, long1_Value, string1_Value, string2_Value, sub2_Value, lat2_Value, long2_Value }

                   );
                        }
                    }

                    else if (radioButton2.Checked == true)
                    {
                        req_Similarity = Convert.ToInt32(trackBar2.Value.ToString());
                        SimilarString sim_string = new SimilarString();
                        actual_Similarity = sim_string.SimilarText(string1_Value, string2_Value, out percent);
                        //       test_Lable.Text += " Num === " + percent.ToString();
                        if (Math.Abs(lat1_Value - lat2_Value) <= dist_L && Math.Abs(long1_Value - long2_Value) <= dist_L && req_Similarity <= percent)
                        {

                            dt3.Rows.Add(new Object[] { sub1_Value, lat1_Value, long1_Value, string1_Value, string2_Value, sub2_Value, lat2_Value, long2_Value }

                   );
                        }
                    }

                    else if (radioButton3.Checked == true)
                    {
                        req_Similarity = Convert.ToInt32(trackBar2.Value.ToString());
                        JaroWinklerDistance sim2_string = new JaroWinklerDistance();
                        percent = sim2_string.proximity(string1_Value, string2_Value);
                        //       test_Lable.Text    sim2_string.proximity = sim2_string.SimilarText(string1_Value, string2_Value, out percent);
                        //       test_Lable.Text += " Num === " + percent.ToString();
                        // label12.Text = percent.ToString();
                        if (Math.Abs(lat1_Value - lat2_Value) <= dist_L)
                        {
                            if (Math.Abs(long1_Value - long2_Value) <= dist_L)
                            {
                                if (req_Similarity <= (percent * 100))
                                {

                                    dt3.Rows.Add(new Object[] { sub1_Value, lat1_Value, long1_Value, string1_Value, string2_Value, sub2_Value, lat2_Value, long2_Value });
                                }
                            }
                        }
                    }

                    /************/


                    /*
                    if (Math.Abs(lat1_Value - lat2_Value) <= dist_L && Math.Abs(long1_Value - long2_Value) <= dist_L)
                        {

                            dt3.Rows.Add(new Object[]{sub1_Value , lat1_Value , long1_Value , sub2_Value , lat2_Value , long2_Value

                     });

                        }

                    */

                }
                dataGridView3.DataSource = dt3;
             
                result.Text =  dt3.Rows.Count.ToString();

    
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
            string filePath = @"D:\SameAs.txt";
            //  label12.Text = filePath;
            int ds_r = dt3.Rows.Count;
            String[] lines = new String[ds_r];
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";

            saveFileDialog1.Title = "Save an Text File";
            saveFileDialog1.FileName = "SameAs.txt";
            // set filters - this can be done in properties as well
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                //   using (Stream s = File.Open(saveFileDialog1.FileName, FileMode.Create)) ;


                filePath = saveFileDialog1.FileName.ToString();


            }


            // l1.Text = filePath.ToString();


            for (int i = 0; i < dt3.Rows.Count; i++)

            {
                lines[i] = "<rdf:Description rdf:about=\"" + dt3.Rows[i][0].ToString() + "\"> \n <owl:sameAs rdf:resource =\"" + dt3.Rows[i][5].ToString() + "\"/> \n </rdf:Description>";
            }

            File.WriteAllLines(filePath, lines);

        }

        private void lat2_SelectedItemChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            label_Sim.Text = trackBar2.Value.ToString() + " %";
        }

        private void trackBar1_Scroll_1(object sender, EventArgs e)
        {
            label_Distance.Text = trackBar1.Value.ToString();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            groupBox3.Visible = true;

            trackBar_Similarity.Visible = false;



        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            groupBox3.Visible = false;
            trackBar_Similarity.Visible = true;

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            // saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";
            // saveFileDialog1.Title = "Save an Image File";
            saveFileDialog1.Title = "Save an Image File";
            saveFileDialog1.FileName = "unknown.txt";
            // set filters - this can be done in properties as well
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                //   using (Stream s = File.Open(saveFileDialog1.FileName, FileMode.Create)) ;


                //   l1.Text = saveFileDialog1.FileName.ToString();


            }

            //  saveFileDialog1.ShowDialog();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            trackBar_Similarity.Visible = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Algorithms frm2 = new Algorithms();
            frm2.Show();
        }

  
    }
    }

