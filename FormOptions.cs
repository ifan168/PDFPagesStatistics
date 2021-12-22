using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Collections;
using ifan.PDF;

namespace PDF页面统计
{
    public partial class FormOptions : Form
    {
        OptionsHelper optionsHelper = new OptionsHelper();
        MyOptions myOptions = new MyOptions();
        public FormOptions()
        {
            InitializeComponent();
            ReadOptions();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否要恢复默认设置", "提示框", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ResetOptions();
            }

        }

        private void ResetOptions()
        {
            textBox1.Text = "1189";
            textBox2.Text = "841";
            textBox3.Text = "841";
            textBox4.Text = "594";
            textBox5.Text = "594";
            textBox6.Text = "420";
            textBox7.Text = "420";
            textBox8.Text = "297";
            textBox9.Text = "297";
            textBox10.Text = "210";
            textBox11.Text = "5";
            WriteOptions();
            myOptions = optionsHelper.GetOptions();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WriteOptions();
            this.Close();
        }

        private void WriteOptions()
        {
            List<Paper> lstPaper = new List<Paper>();
            lstPaper.Add(new Paper(PaperSize.A0, double.Parse(textBox1.Text), double.Parse(textBox2.Text)));
            lstPaper.Add(new Paper(PaperSize.A1, double.Parse(textBox3.Text), double.Parse(textBox4.Text)));
            lstPaper.Add(new Paper(PaperSize.A2, double.Parse(textBox5.Text), double.Parse(textBox6.Text)));
            lstPaper.Add(new Paper(PaperSize.A3, double.Parse(textBox7.Text), double.Parse(textBox8.Text)));
            lstPaper.Add(new Paper(PaperSize.A4, double.Parse(textBox9.Text), double.Parse(textBox10.Text)));

            myOptions.PaperSizes = lstPaper;
            myOptions.tolerance = int.Parse(textBox11.Text);
            optionsHelper.SetOptions(myOptions);
        }



        private void ReadOptions()
        {
            if (optionsHelper.GetOptions()!= null)
            {
                myOptions = optionsHelper.GetOptions();
                textBox1.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A0).First().width.ToString();
                textBox2.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A0).First().height.ToString();
                textBox3.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A1).First().width.ToString();
                textBox4.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A1).First().height.ToString();
                textBox5.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A2).First().width.ToString();
                textBox6.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A2).First().height.ToString();
                textBox7.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A3).First().width.ToString();
                textBox8.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A3).First().height.ToString();
                textBox9.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A4).First().width.ToString();
                textBox10.Text = myOptions.PaperSizes.Where(x => x.Size == PaperSize.A4).First().height.ToString();
                textBox11.Text = myOptions.tolerance.ToString();




            }
            else
            {
                ResetOptions();
            }


        }

        private void FormOptions_Load(object sender, EventArgs e)
        {

        }

        private void FormOptions_FormClosed(object sender, FormClosedEventArgs e)
        {
            ifanPdfPaper.m_papers = myOptions.PaperSizes;
            ifanPdfPaper.intTolerance = myOptions.tolerance;
        }
    }
}
