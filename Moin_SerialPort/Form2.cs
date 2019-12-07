using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Moin_SerialPort
{
    public partial class Form2 : MetroFramework.Forms.MetroForm
    {
        public Form2()
        {
            InitializeComponent();
        }
        private bool Flag_text = false;
        const string Filename = "公式.txt";
        private void Form2_Load(object sender, EventArgs e)
        {
            
            try
            {
                StreamReader sr = new StreamReader(Filename, System.Text.Encoding.Default);
                richTextBox1.Text = sr.ReadToEnd();
                sr.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("无法打开公式 请检查文件（公式.txt）");
            }
            Flag_text = false;
        }
        private void write()
        {
            try
            {
                StreamWriter sw = new StreamWriter(Filename, false);
                sw.Write(richTextBox1.Text);
                sw.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("无法保存");
            }
            Flag_text = false;
            
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            write();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            if (Flag_text==true)
            { 
                if (MessageBox.Show("是否保存当前文件？", "温馨提示", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information)
                        == DialogResult.Yes)
                {

                    write();
                }
              }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            Flag_text = true;
        }
    }
}
