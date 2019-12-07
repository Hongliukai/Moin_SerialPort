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
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        [DllImport("User32.dll", EntryPoint = "SetParent")]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "ShowWindow")]
        private static extern int ShowWindow(IntPtr hwnd, int nCmdShow);
        string winds_handle="";
        private bool rev_flag = true;
        public Form1()
        {
            InitializeComponent();

        }
        void Application_Idle(object sender, EventArgs e)
        {
            if (appbox.IsStarted)
            {
                if (!appbox.AppProcess.HasExited)
                {
                    try
                    {
                        winds_handle=(string.Format("Main Window Handle:{0}|Original Parent Window Handle:{1}",
                            appbox.AppProcess.MainWindowHandle,appbox.embedResult) );
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            //throw new NotImplementedException();
        }
        private SerialPort ComDevice = new SerialPort();
        SmileWei.EmbeddedApp.AppContainer appbox = new SmileWei.EmbeddedApp.AppContainer(false);
        private void Form1_Load(object sender, EventArgs e)
        {
            rev_clear_button.Enabled = false;
            send_clear_button.Enabled = false;
            send_data_button.Enabled = false;
            cbbComList.Items.AddRange(SerialPort.GetPortNames());
            if (cbbComList.Items.Count > 0)
            {
                cbbComList.SelectedIndex = 0;
            }
            cbbBaudRate.SelectedIndex = 4;
            cbbDataBits.SelectedIndex = 2;
            cbbParity.SelectedIndex = 0;
            cbbStopBits.SelectedIndex = 0;
            pictureBox1.BackgroundImage = Properties.Resources.red;
            rev_text_box.Text = "";
            send_textbox.Text = "";
            ComDevice.DataReceived += new SerialDataReceivedEventHandler(Com_DataReceived);//绑定事件
            send_hex_button.Checked = true;
            rev_HEX_button.Checked = true;
            metroCheckBox4.Checked = false;
            sendxunhuan_timer1.Stop();
            
            appbox.AppFilename = "";
            appbox.AppProcess = null;
            appbox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            appbox.VerticalScroll.Enabled=true;
            appbox.Location = new System.Drawing.Point(-3, 6);
            appbox.Name = "appbox";
            appbox.Size = new System.Drawing.Size(295, 341);
            appbox.TabIndex = 2;
            appbox.ShowEmbedResult = true;
            Application.Idle += Application_Idle;
            this.metroPanel4.Controls.Add(appbox);
        }
        //打开串口
        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (cbbComList.Items.Count == 0)
            {
                MessageBox.Show("没有发现串口,请检查设备！");
                return;
            }

            if (ComDevice.IsOpen == false)
            {
                ComDevice.PortName = cbbComList.SelectedItem.ToString();
                ComDevice.BaudRate = Convert.ToInt32(cbbBaudRate.SelectedItem.ToString());
                ComDevice.Parity = (Parity)Convert.ToInt32(cbbParity.SelectedIndex.ToString());
                ComDevice.DataBits = Convert.ToInt32(cbbDataBits.SelectedItem.ToString());
                ComDevice.StopBits = (StopBits)Convert.ToInt32(cbbStopBits.SelectedItem.ToString());
                try
                {
                    ComDevice.Open();
                    rev_clear_button.Enabled = true;
                    send_clear_button.Enabled = true;
                    send_data_button.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                btnOpen.Text = "关闭串口";
                pictureBox1.BackgroundImage = Properties.Resources.green;
            }
            else
            {
                try
                {
                    ComDevice.Close();
                    rev_clear_button.Enabled = false;
                    send_clear_button.Enabled = false;
                    send_data_button.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                btnOpen.Text = "打开串口";
                pictureBox1.BackgroundImage = Properties.Resources.red;
            }

            cbbComList.Enabled = !ComDevice.IsOpen;
            cbbBaudRate.Enabled = !ComDevice.IsOpen;
            cbbParity.Enabled = !ComDevice.IsOpen;
            cbbDataBits.Enabled = !ComDevice.IsOpen;
            cbbStopBits.Enabled = !ComDevice.IsOpen;
        }
        private void cbbComList_Click(object sender, EventArgs e)
        {
            cbbComList.Items.Clear();
            cbbComList.Items.AddRange(SerialPort.GetPortNames());
            if (cbbComList.Items.Count > 0)
            {
                cbbComList.SelectedIndex = 0;
            }
        }
        //发送数据
        public bool SendData(byte[] data)
        {
            if (ComDevice.IsOpen)
            {
                try
                {
                    ComDevice.Write(data, 0, data.Length);//发送数据
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("串口未打开", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return false;
        }
        //HEX-String相互转换
        private byte[] strToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0) hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2).Replace(" ", ""), 16);
            return returnBytes;
        }

        //发送数据
        private void send_data_button_Click(object sender, EventArgs e)
        {
            byte[] sendData = null;

            if (send_hex_button.Checked)
            {
                sendData = strToHexByte(send_textbox.Text.Trim());
            }
            else if (send_ascll_button.Checked)
            {
                sendData = Encoding.ASCII.GetBytes(send_textbox.Text.Trim());
            }
            else if (send_utf8_button.Checked)
            {
                sendData = Encoding.UTF8.GetBytes(send_textbox.Text.Trim());
            }
            else if (send_unicode_button.Checked)
            {
                sendData = Encoding.Unicode.GetBytes(send_textbox.Text.Trim());
            }
            else
            {
                sendData = Encoding.ASCII.GetBytes(send_textbox.Text.Trim());
            }

            if (this.SendData(sendData))//发送数据成功计数
            {
                send_Bytes.Invoke(new MethodInvoker(delegate
                {
                    send_Bytes.Text = (int.Parse(send_Bytes.Text) + sendData.Length).ToString();
                    if(show_send_cB.Checked)
                    {
                        AddData(sendData, false);
                    }
                }));
            }
            else
            {
                MessageBox.Show("发送数据失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private string fenxi(string source)
        {
            Regex regex_1 = new Regex(@",\d+\}", RegexOptions.IgnoreCase);
            Regex regex_num = new Regex(@"\d+", RegexOptions.IgnoreCase);
            Regex regex_2 = new Regex(@"\w{1,},\d+", RegexOptions.IgnoreCase);
            for (int i = 0; i <= gongshi_list.Length; i++)
            {
                if (i == gongshi_list.Length)
                {
                    break;
                }
                else
                {
                    try
                    {
                        string[] head_info = new string[2];
                        string[] info_item = new string[2];
                        head_info = gongshi_list[i].Split(':');
                        Match col = regex_1.Match(head_info[0]);
                        Match col_real = regex_num.Match(col.Value);
                        // Match[] marr = matches.OfType<Match>().ToArray();
                        if ((source.Replace(" ", "").Length) / 2 == int.Parse(col_real.Value))
                        {
                            //MessageBox.Show(source.Replace(" ", "").Length.ToString() + "   " + col_real.Value);

                            string result = "分析：";
                            //MessageBox.Show(source.Replace(" ", "").Length.ToString() + "   " + col_real.Value);
                            MatchCollection col_2 = regex_2.Matches(head_info[1]);
                            source = source.Replace(" ", "");
                            char[] source_arrary = source.ToCharArray();
                            int num = 0;
                            foreach (Match m in col_2)
                            {
                                info_item = m.Value.Split(',');
                                result = result + " " + info_item[0] + ":";
                                for (int q = 0; q < int.Parse(info_item[1]); q++)
                                {
                                    result = result + source_arrary[num].ToString() + source_arrary[num + 1].ToString();
                                    num += 2;
                                }
                            }
                            result += "|";
                            return result;
                        }
                    }
                    catch(Exception ex)
                    {
                        return "";
                    }
                    

                }
            }
            return "";
        }
        //接收数据转换-显示
        private void AddContent(string content)
        {
            if (fenxi_flag == false)
            {
                    this.BeginInvoke(new MethodInvoker(delegate
                {
                    if (auto_line_cB.Checked && rev_text_box.Text.Length > 0)
                    {
                        rev_text_box.AppendText("\r\n");
                    }
                    if (show_time_cB.Checked)
                    {
                        rev_text_box.AppendText("[" + DateTime.Now.TimeOfDay.ToString() + "]");
                    }
                    rev_text_box.AppendText(content);
                }));
            }
            else
            {
                this.BeginInvoke(new MethodInvoker(delegate
                {
                    if (auto_line_cB.Checked && rev_text_box.Text.Length > 0)
                    {
                        rev_text_box.AppendText("\r\n");
                    }
                    if (show_time_cB.Checked)
                    {
                        rev_text_box.AppendText("[" + DateTime.Now.TimeOfDay.ToString() + "]");
                    }
                    rev_text_box.AppendText(content);
                    string fengxi_rev = "|";
                    fengxi_rev += fenxi(content);
                    rev_text_box.AppendText(fengxi_rev);
                }));
            }
        }
        public void AddData(byte[] data,bool Is_Rev)
        {
            if (rev_HEX_button.Checked)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < data.Length; i++)
                {
                    sb.AppendFormat("{0:x2}" + " ", data[i]);
                }
                AddContent(sb.ToString().ToUpper());
            }
            else if (rev_ascll_button.Checked)
            {
                AddContent(new ASCIIEncoding().GetString(data));
            }
            else if (rev_utf8_button.Checked)
            {
                AddContent(new UTF8Encoding().GetString(data));
            }
            else if (rev_unicode_button.Checked)
            {
                AddContent(new UnicodeEncoding().GetString(data));
            }
            else
            { }
            if (Is_Rev == true)
            {
                rev_bytes.Invoke(new MethodInvoker(delegate
                {
                    rev_bytes.Text = (int.Parse(rev_bytes.Text) + data.Length).ToString();
                }));
            }
        }
        //接收数据
        private void Com_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte[] ReDatas = new byte[ComDevice.BytesToRead];
            ComDevice.Read(ReDatas, 0, ReDatas.Length);//读取数据
            if(rev_flag==true)
            {
                this.AddData(ReDatas, true);//输出数据
            }
            
        }

        private void rev_clear_button_Click(object sender, EventArgs e)
        {
            rev_text_box.Clear();
            rev_bytes.Text = "0";
        }

        private void send_clear_button_Click(object sender, EventArgs e)
        {
            send_textbox.Clear();
            send_Bytes.Text= "0";
        }

        private void metroCheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            if(metroCheckBox4.Checked)
            {
                sendxunhuan_timer1.Interval = Convert.ToInt32(numericUpDown1.Value);
                sendxunhuan_timer1.Start();
            }
            else
            {
                sendxunhuan_timer1.Stop();
            }
            
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if(metroCheckBox4.Checked)
            {
                sendxunhuan_timer1.Stop();
                sendxunhuan_timer1.Interval = Convert.ToInt32(numericUpDown1.Value);
                sendxunhuan_timer1.Start();
            }
        }

        private void sendxunhuan_timer1_Tick(object sender, EventArgs e)
        {
            byte[] sendData1 = null;

            if (send_hex_button.Checked)
            {
                sendData1 = strToHexByte(send_textbox.Text.Trim());
            }
            else if (send_ascll_button.Checked)
            {
                sendData1 = Encoding.ASCII.GetBytes(send_textbox.Text.Trim());
            }
            else if (send_utf8_button.Checked)
            {
                sendData1 = Encoding.UTF8.GetBytes(send_textbox.Text.Trim());
            }
            else if (send_unicode_button.Checked)
            {
                sendData1 = Encoding.Unicode.GetBytes(send_textbox.Text.Trim());
            }
            else
            {
                sendData1 = Encoding.ASCII.GetBytes(send_textbox.Text.Trim());
            }

            if (this.SendData(sendData1))//发送数据成功计数
            {
                send_Bytes.Invoke(new MethodInvoker(delegate
                {
                    send_Bytes.Text = (int.Parse(send_Bytes.Text) + sendData1.Length).ToString();
                    if (show_send_cB.Checked)
                    {
                        AddData(sendData1, false);
                    }
                }));
            }
            else
            {
                MessageBox.Show("发送数据失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //校验算法
        //校验和
        public Byte XOR_cal(Byte[] source,int n)
        {
            Byte xor_result = source[0];
            for(int i=1;i<n;i++)
            {
                xor_result = (Byte)(xor_result ^ source[i]);
            }
            return xor_result;
        }
        //CRC16
        private static ushort[] crctab = new ushort[256]{
                    0x0000, 0x1021, 0x2042, 0x3063, 0x4084, 0x50a5, 0x60c6, 0x70e7,
                    0x8108, 0x9129, 0xa14a, 0xb16b, 0xc18c, 0xd1ad, 0xe1ce, 0xf1ef,
                    0x1231, 0x0210, 0x3273, 0x2252, 0x52b5, 0x4294, 0x72f7, 0x62d6,
                    0x9339, 0x8318, 0xb37b, 0xa35a, 0xd3bd, 0xc39c, 0xf3ff, 0xe3de,
                    0x2462, 0x3443, 0x0420, 0x1401, 0x64e6, 0x74c7, 0x44a4, 0x5485,
                    0xa56a, 0xb54b, 0x8528, 0x9509, 0xe5ee, 0xf5cf, 0xc5ac, 0xd58d,
                    0x3653, 0x2672, 0x1611, 0x0630, 0x76d7, 0x66f6, 0x5695, 0x46b4,
                    0xb75b, 0xa77a, 0x9719, 0x8738, 0xf7df, 0xe7fe, 0xd79d, 0xc7bc,
                    0x48c4, 0x58e5, 0x6886, 0x78a7, 0x0840, 0x1861, 0x2802, 0x3823,
                    0xc9cc, 0xd9ed, 0xe98e, 0xf9af, 0x8948, 0x9969, 0xa90a, 0xb92b,
                    0x5af5, 0x4ad4, 0x7ab7, 0x6a96, 0x1a71, 0x0a50, 0x3a33, 0x2a12,
                    0xdbfd, 0xcbdc, 0xfbbf, 0xeb9e, 0x9b79, 0x8b58, 0xbb3b, 0xab1a,
                    0x6ca6, 0x7c87, 0x4ce4, 0x5cc5, 0x2c22, 0x3c03, 0x0c60, 0x1c41,
                    0xedae, 0xfd8f, 0xcdec, 0xddcd, 0xad2a, 0xbd0b, 0x8d68, 0x9d49,
                    0x7e97, 0x6eb6, 0x5ed5, 0x4ef4, 0x3e13, 0x2e32, 0x1e51, 0x0e70,
                    0xff9f, 0xefbe, 0xdfdd, 0xcffc, 0xbf1b, 0xaf3a, 0x9f59, 0x8f78,
                    0x9188, 0x81a9, 0xb1ca, 0xa1eb, 0xd10c, 0xc12d, 0xf14e, 0xe16f,
                    0x1080, 0x00a1, 0x30c2, 0x20e3, 0x5004, 0x4025, 0x7046, 0x6067,
                    0x83b9, 0x9398, 0xa3fb, 0xb3da, 0xc33d, 0xd31c, 0xe37f, 0xf35e,
                    0x02b1, 0x1290, 0x22f3, 0x32d2, 0x4235, 0x5214, 0x6277, 0x7256,
                    0xb5ea, 0xa5cb, 0x95a8, 0x8589, 0xf56e, 0xe54f, 0xd52c, 0xc50d,
                    0x34e2, 0x24c3, 0x14a0, 0x0481, 0x7466, 0x6447, 0x5424, 0x4405,
                    0xa7db, 0xb7fa, 0x8799, 0x97b8, 0xe75f, 0xf77e, 0xc71d, 0xd73c,
                    0x26d3, 0x36f2, 0x0691, 0x16b0, 0x6657, 0x7676, 0x4615, 0x5634,
                    0xd94c, 0xc96d, 0xf90e, 0xe92f, 0x99c8, 0x89e9, 0xb98a, 0xa9ab,
                    0x5844, 0x4865, 0x7806, 0x6827, 0x18c0, 0x08e1, 0x3882, 0x28a3,
                    0xcb7d, 0xdb5c, 0xeb3f, 0xfb1e, 0x8bf9, 0x9bd8, 0xabbb, 0xbb9a,
                    0x4a75, 0x5a54, 0x6a37, 0x7a16, 0x0af1, 0x1ad0, 0x2ab3, 0x3a92,
                    0xfd2e, 0xed0f, 0xdd6c, 0xcd4d, 0xbdaa, 0xad8b, 0x9de8, 0x8dc9,
                    0x7c26, 0x6c07, 0x5c64, 0x4c45, 0x3ca2, 0x2c83, 0x1ce0, 0x0cc1,
                    0xef1f, 0xff3e, 0xcf5d, 0xdf7c, 0xaf9b, 0xbfba, 0x8fd9, 0x9ff8,
                    0x6e17, 0x7e36, 0x4e55, 0x5e74, 0x2e93, 0x3eb2, 0x0ed1, 0x1ef0
                    };
        private static ushort xcrc(ushort crc, byte cp)
        {
            ushort t1 = 0, t2 = 0, t3 = 0, t4 = 0, t5 = 0, t6 = 0;
            t1 = (ushort)(crc >> 8);
            t2 = (ushort)(t1 & 0xff);
            t3 = (ushort)(cp & 0xff);
            t4 = (ushort)(crc << 8);
            t5 = (ushort)(t2 ^ t3);
            t6 = (ushort)(crctab[t5] ^ t4);
            return t6;
        }
        public static ushort ConCRC(byte[] bufin, int n)
        {
            ushort crc16 = 0;
            byte i;
            //n个数据的CRC校验
            for (i = 0; i < n; i++)
            {
                crc16 = xcrc(crc16, bufin[i]);
            }
            return crc16;

        }
        //Modbus-CRC
        private static Byte[] aucCRCHi = new Byte[256]{
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40};

        private static Byte[] aucCRCLo =new Byte[256] {
                0x00, 0xC0, 0xC1, 0x01, 0xC3, 0x03, 0x02, 0xC2, 0xC6, 0x06, 0x07, 0xC7,
                0x05, 0xC5, 0xC4, 0x04, 0xCC, 0x0C, 0x0D, 0xCD, 0x0F, 0xCF, 0xCE, 0x0E,
                0x0A, 0xCA, 0xCB, 0x0B, 0xC9, 0x09, 0x08, 0xC8, 0xD8, 0x18, 0x19, 0xD9,
                0x1B, 0xDB, 0xDA, 0x1A, 0x1E, 0xDE, 0xDF, 0x1F, 0xDD, 0x1D, 0x1C, 0xDC,
                0x14, 0xD4, 0xD5, 0x15, 0xD7, 0x17, 0x16, 0xD6, 0xD2, 0x12, 0x13, 0xD3,
                0x11, 0xD1, 0xD0, 0x10, 0xF0, 0x30, 0x31, 0xF1, 0x33, 0xF3, 0xF2, 0x32,
                0x36, 0xF6, 0xF7, 0x37, 0xF5, 0x35, 0x34, 0xF4, 0x3C, 0xFC, 0xFD, 0x3D,
                0xFF, 0x3F, 0x3E, 0xFE, 0xFA, 0x3A, 0x3B, 0xFB, 0x39, 0xF9, 0xF8, 0x38,
                0x28, 0xE8, 0xE9, 0x29, 0xEB, 0x2B, 0x2A, 0xEA, 0xEE, 0x2E, 0x2F, 0xEF,
                0x2D, 0xED, 0xEC, 0x2C, 0xE4, 0x24, 0x25, 0xE5, 0x27, 0xE7, 0xE6, 0x26,
                0x22, 0xE2, 0xE3, 0x23, 0xE1, 0x21, 0x20, 0xE0, 0xA0, 0x60, 0x61, 0xA1,
                0x63, 0xA3, 0xA2, 0x62, 0x66, 0xA6, 0xA7, 0x67, 0xA5, 0x65, 0x64, 0xA4,
                0x6C, 0xAC, 0xAD, 0x6D, 0xAF, 0x6F, 0x6E, 0xAE, 0xAA, 0x6A, 0x6B, 0xAB,
                0x69, 0xA9, 0xA8, 0x68, 0x78, 0xB8, 0xB9, 0x79, 0xBB, 0x7B, 0x7A, 0xBA,
                0xBE, 0x7E, 0x7F, 0xBF, 0x7D, 0xBD, 0xBC, 0x7C, 0xB4, 0x74, 0x75, 0xB5,
                0x77, 0xB7, 0xB6, 0x76, 0x72, 0xB2, 0xB3, 0x73, 0xB1, 0x71, 0x70, 0xB0,
                0x50, 0x90, 0x91, 0x51, 0x93, 0x53, 0x52, 0x92, 0x96, 0x56, 0x57, 0x97,
                0x55, 0x95, 0x94, 0x54, 0x9C, 0x5C, 0x5D, 0x9D, 0x5F, 0x9F, 0x9E, 0x5E,
                0x5A, 0x9A, 0x9B, 0x5B, 0x99, 0x59, 0x58, 0x98, 0x88, 0x48, 0x49, 0x89,
                0x4B, 0x8B, 0x8A, 0x4A, 0x4E, 0x8E, 0x8F, 0x4F, 0x8D, 0x4D, 0x4C, 0x8C,
                0x44, 0x84, 0x85, 0x45, 0x87, 0x47, 0x46, 0x86, 0x82, 0x42, 0x43, 0x83,
                0x41, 0x81, 0x80, 0x40};

        public ushort usMBCRC16(Byte[] pucFrame, int usLen)
        {
            Byte ucCRCHi = 0xFF;
            Byte ucCRCLo = 0xFF;
            int iIndex;
            int n = 0;
            while ((usLen--)>0)
            {
                iIndex = ucCRCLo ^ pucFrame[n];
                n++;
                ucCRCLo = (Byte)(ucCRCHi ^ aucCRCHi[iIndex]);
                ucCRCHi = aucCRCLo[iIndex];
            }
            ushort result =(ushort)( Convert.ToUInt16(ucCRCHi)<<8);
            result = (ushort)(result | Convert.ToUInt16(ucCRCLo));
            return result;
        }
        //crc32
        private static UInt32[] crcTable =new UInt32[256]
        {
          0x00000000, 0x04c11db7, 0x09823b6e, 0x0d4326d9, 0x130476dc, 0x17c56b6b, 0x1a864db2, 0x1e475005,
          0x2608edb8, 0x22c9f00f, 0x2f8ad6d6, 0x2b4bcb61, 0x350c9b64, 0x31cd86d3, 0x3c8ea00a, 0x384fbdbd,
          0x4c11db70, 0x48d0c6c7, 0x4593e01e, 0x4152fda9, 0x5f15adac, 0x5bd4b01b, 0x569796c2, 0x52568b75,
          0x6a1936c8, 0x6ed82b7f, 0x639b0da6, 0x675a1011, 0x791d4014, 0x7ddc5da3, 0x709f7b7a, 0x745e66cd,
          0x9823b6e0, 0x9ce2ab57, 0x91a18d8e, 0x95609039, 0x8b27c03c, 0x8fe6dd8b, 0x82a5fb52, 0x8664e6e5,
          0xbe2b5b58, 0xbaea46ef, 0xb7a96036, 0xb3687d81, 0xad2f2d84, 0xa9ee3033, 0xa4ad16ea, 0xa06c0b5d,
          0xd4326d90, 0xd0f37027, 0xddb056fe, 0xd9714b49, 0xc7361b4c, 0xc3f706fb, 0xceb42022, 0xca753d95,
          0xf23a8028, 0xf6fb9d9f, 0xfbb8bb46, 0xff79a6f1, 0xe13ef6f4, 0xe5ffeb43, 0xe8bccd9a, 0xec7dd02d,
          0x34867077, 0x30476dc0, 0x3d044b19, 0x39c556ae, 0x278206ab, 0x23431b1c, 0x2e003dc5, 0x2ac12072,
          0x128e9dcf, 0x164f8078, 0x1b0ca6a1, 0x1fcdbb16, 0x018aeb13, 0x054bf6a4, 0x0808d07d, 0x0cc9cdca,
          0x7897ab07, 0x7c56b6b0, 0x71159069, 0x75d48dde, 0x6b93dddb, 0x6f52c06c, 0x6211e6b5, 0x66d0fb02,
          0x5e9f46bf, 0x5a5e5b08, 0x571d7dd1, 0x53dc6066, 0x4d9b3063, 0x495a2dd4, 0x44190b0d, 0x40d816ba,
          0xaca5c697, 0xa864db20, 0xa527fdf9, 0xa1e6e04e, 0xbfa1b04b, 0xbb60adfc, 0xb6238b25, 0xb2e29692,
          0x8aad2b2f, 0x8e6c3698, 0x832f1041, 0x87ee0df6, 0x99a95df3, 0x9d684044, 0x902b669d, 0x94ea7b2a,
          0xe0b41de7, 0xe4750050, 0xe9362689, 0xedf73b3e, 0xf3b06b3b, 0xf771768c, 0xfa325055, 0xfef34de2,
          0xc6bcf05f, 0xc27dede8, 0xcf3ecb31, 0xcbffd686, 0xd5b88683, 0xd1799b34, 0xdc3abded, 0xd8fba05a,
          0x690ce0ee, 0x6dcdfd59, 0x608edb80, 0x644fc637, 0x7a089632, 0x7ec98b85, 0x738aad5c, 0x774bb0eb,
          0x4f040d56, 0x4bc510e1, 0x46863638, 0x42472b8f, 0x5c007b8a, 0x58c1663d, 0x558240e4, 0x51435d53,
          0x251d3b9e, 0x21dc2629, 0x2c9f00f0, 0x285e1d47, 0x36194d42, 0x32d850f5, 0x3f9b762c, 0x3b5a6b9b,
          0x0315d626, 0x07d4cb91, 0x0a97ed48, 0x0e56f0ff, 0x1011a0fa, 0x14d0bd4d, 0x19939b94, 0x1d528623,
          0xf12f560e, 0xf5ee4bb9, 0xf8ad6d60, 0xfc6c70d7, 0xe22b20d2, 0xe6ea3d65, 0xeba91bbc, 0xef68060b,
          0xd727bbb6, 0xd3e6a601, 0xdea580d8, 0xda649d6f, 0xc423cd6a, 0xc0e2d0dd, 0xcda1f604, 0xc960ebb3,
          0xbd3e8d7e, 0xb9ff90c9, 0xb4bcb610, 0xb07daba7, 0xae3afba2, 0xaafbe615, 0xa7b8c0cc, 0xa379dd7b,
          0x9b3660c6, 0x9ff77d71, 0x92b45ba8, 0x9675461f, 0x8832161a, 0x8cf30bad, 0x81b02d74, 0x857130c3,
          0x5d8a9099, 0x594b8d2e, 0x5408abf7, 0x50c9b640, 0x4e8ee645, 0x4a4ffbf2, 0x470cdd2b, 0x43cdc09c,
          0x7b827d21, 0x7f436096, 0x7200464f, 0x76c15bf8, 0x68860bfd, 0x6c47164a, 0x61043093, 0x65c52d24,
          0x119b4be9, 0x155a565e, 0x18197087, 0x1cd86d30, 0x029f3d35, 0x065e2082, 0x0b1d065b, 0x0fdc1bec,
          0x3793a651, 0x3352bbe6, 0x3e119d3f, 0x3ad08088, 0x2497d08d, 0x2056cd3a, 0x2d15ebe3, 0x29d4f654,
          0xc5a92679, 0xc1683bce, 0xcc2b1d17, 0xc8ea00a0, 0xd6ad50a5, 0xd26c4d12, 0xdf2f6bcb, 0xdbee767c,
          0xe3a1cbc1, 0xe760d676, 0xea23f0af, 0xeee2ed18, 0xf0a5bd1d, 0xf464a0aa, 0xf9278673, 0xfde69bc4,
          0x89b8fd09, 0x8d79e0be, 0x803ac667, 0x84fbdbd0, 0x9abc8bd5, 0x9e7d9662, 0x933eb0bb, 0x97ffad0c,
          0xafb010b1, 0xab710d06, 0xa6322bdf, 0xa2f33668, 0xbcb4666d, 0xb8757bda, 0xb5365d03, 0xb1f740b4
        };

        public static uint GetCRC32(byte[] bytes,int n)
        {
            uint crc = 0xFFFFFFFF;

            for (uint i = 0; i < n; i++)
            {
                crc = (crc << 8) ^ crcTable[(crc >> 24) ^ bytes[i]];
            }

            return crc;
        }
        //校验算法结束
        private void metroButton1_Click(object sender, EventArgs e)
        {
            byte[] sendData2 = null;
            sendData2 = strToHexByte(jiaoyan_data.Text.Trim());
            if (sendData2.Length > 0)
            {
                if (jiaoyan_ComboBox1.SelectedItem != null)
                {
                    switch (jiaoyan_ComboBox1.SelectedItem.ToString())
                    {
                        case "校验和":
                            ushort result_xor = XOR_cal(sendData2, sendData2.Length);
                            jiaoyanjieguo_textB.Text = String.Format("{0:X2}", result_xor);
                            break;
                        case "CRC16计算":
                            ushort result_crc16 = ConCRC(sendData2, sendData2.Length);
                            jiaoyanjieguo_textB.Text = String.Format("{0:X4}", result_crc16);
                            break;
                        case "ModbusCRC计算":
                            ushort result_crcMod = usMBCRC16(sendData2, sendData2.Length);
                            jiaoyanjieguo_textB.Text = String.Format("{0:X4}", result_crcMod);
                            break;
                        case "CRC32计算":
                            uint result_crc32 = GetCRC32(sendData2, sendData2.Length);
                            jiaoyanjieguo_textB.Text = String.Format("{0:X8}", result_crc32);
                            break;
                        

                    }
                }
                else
                {
                    MessageBox.Show("请选择校验方式");

                }

            }
            
        }

        private void 快捷校验ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            byte[] sendData3 = null;
            sendData3 = strToHexByte(send_textbox.SelectedText.Trim());
            string insert="";
            if (sendData3.Length > 0)
            {
                if (quick_jiaoyan_cB.SelectedItem != null)
                {
                    switch (quick_jiaoyan_cB.SelectedItem.ToString())
                    {
                        case "校验和":
                            ushort result_xor = XOR_cal(sendData3, sendData3.Length);
                            insert = String.Format("{0:X2}", result_xor);
                            break;
                        case "CRC16计算":
                            ushort result_crc16 = ConCRC(sendData3, sendData3.Length);
                            insert = String.Format("{0:X4}", result_crc16);
                            break;
                        case "ModbusCRC计算":
                            ushort result_crcMod = usMBCRC16(sendData3, sendData3.Length);
                            insert = String.Format("{0:X4}", result_crcMod);
                            break;
                        case "CRC32计算":
                            uint result_crc32 = GetCRC32(sendData3, sendData3.Length);
                            insert = String.Format("{0:X8}", result_crc32);
                            break;


                    }
                    send_textbox.Text=send_textbox.Text.Insert(send_textbox.SelectionStart+send_textbox.SelectedText.Length,insert);
                }
                else
                {
                    MessageBox.Show("请选择校验方式");

                }
            }
            send_textbox.Text.Insert(send_textbox.SelectionStart+send_textbox.SelectedText.Length, insert);
        }
        //快捷校验
        private void 全选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            send_textbox.Focus();
            send_textbox.SelectAll();
        }
        string rev_tool = "";
        private void 剪切ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rev_tool = send_textbox.SelectedText;
            Clipboard.SetText(rev_tool);
            send_textbox.Text=send_textbox.Text.Remove(send_textbox.SelectionStart, send_textbox.SelectedText.Length);
        }

        private void 复制toolStripMenuItem_Click(object sender, EventArgs e)
        {
            rev_tool = send_textbox.SelectedText;
            Clipboard.SetText(rev_tool);
        }

        private void 粘贴toolStripMenuItem_Click(object sender, EventArgs e)
        {
            send_textbox.Text = send_textbox.Text.Insert(send_textbox.SelectionStart, Clipboard.GetText());
        }
        string openFile_path = "";
        private void metroButton2_Click(object sender, EventArgs e)
        {
            byte[] binchar = new byte[] { };
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "二进制文件(Bin,)|*bin;|bin 二进制文件(*.bin)|*.bin|所有文件(*.*)|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                openFile_path = openFile.FileName;
                bin_file_label_show.Text = openFile.FileName.Insert(40,"\n");
                
            }
        }

        //bin文件发送
        public byte[] buff;
        public byte[] data_buff;
        public byte[] crc_buff;
        private void metroButton3_Click(object sender, EventArgs e)
        {
            this.BeginInvoke(new MethodInvoker(delegate
            {
                
            
            char[] high=metroTextBox1.Text.Trim().ToCharArray();
            char[] low = metroTextBox1.Text.Trim().ToCharArray();
            Byte id_high =(Byte)((int)(Convert.ToByte(high[0])*256)+(int)Convert.ToByte(high[1]));
            Byte id_low = (Byte)((int)(Convert.ToByte(low[0]) * 256) + (int)Convert.ToByte(low[1]));
            crc_buff = new byte[2] { 0x00, 0x00 };
            FileStream Myfile = new FileStream(openFile_path.Trim(), FileMode.Open, FileAccess.Read);
            BinaryReader binreader = new BinaryReader(Myfile);
            int file_len = (int)Myfile.Length;//获取bin文件长度
            string Mytext = "";
            int file_len1 = file_len / 256;
            Byte file_len_high = Convert.ToByte(file_len1);
            Byte file_len_low = Convert.ToByte(file_len - file_len1 * 256);
            //头
            buff = new byte[6] { 0xFE, 0x08, 0x00, 0x00, file_len_high, file_len_low };
            buff = buff.Concat(crc_buff).ToArray();
            ConCRC(buff, 6);
            this.SendData(buff);
            foreach (byte j in buff)
            {
                Mytext += j.ToString("x2"); //大写 8位显示 增加前导0
                Mytext += " ";
            }
            metroTextBox2.Text = Mytext;
            System.Threading.Thread.Sleep(15000);
            ushort packet_head = 1;
            int packet_head1 = 0;
            Byte packet_head_high = 0;
            Byte packet_head_low = 0;
            //数据
            while (file_len > 0)
            {
                if (file_len <= 32)
                {
                    buff = new byte[file_len];
                    buff = binreader.ReadBytes(file_len);
                    packet_head1 = packet_head / 256;
                    packet_head_high = Convert.ToByte(packet_head1);
                    packet_head_low = Convert.ToByte(packet_head - packet_head1 * 256);
                    data_buff = new byte[6] { 0xFF, 0x08, 0x00, 0x00, packet_head_high, packet_head_low };
                    data_buff = data_buff.Concat(buff).ToArray();
                    data_buff = data_buff.Concat(crc_buff).ToArray();
                    ConCRC(data_buff, file_len + 6);
                    this.SendData(data_buff);
                    foreach (byte j in data_buff)
                    {
                        metroTextBox2.Text += j.ToString("x2"); //大写 8位显示 增加前导0
                        metroTextBox2.Text += " ";
                    }
                    System.Threading.Thread.Sleep(2000);
                    file_len = 0;
                }
                else
                {
                    buff = new byte[32];
                    buff = binreader.ReadBytes(32);
                    packet_head1 = packet_head / 256;
                    packet_head_high = Convert.ToByte(packet_head1);
                    packet_head_low = Convert.ToByte(packet_head - packet_head1 * 256);
                    data_buff = new byte[6] { 0xFF, 0x08, 0x00, 0x00, packet_head_high, packet_head_low };
                    data_buff = data_buff.Concat(buff).ToArray();
                    data_buff = data_buff.Concat(crc_buff).ToArray();
                    ConCRC(data_buff, 38);
                    this.SendData(data_buff);
                    foreach (byte j in data_buff)
                    {
                        metroTextBox2.Text += j.ToString("x2"); //大写 8位显示 增加前导0
                        metroTextBox2.Text += " ";
                    }

                    file_len = file_len - 32;
                    System.Threading.Thread.Sleep(2000);
                }
                packet_head++;
                //lunxunsend(byte1, byte2);
                //System.Threading.Thread.Sleep(1000);
                //lunxunsend(byte1, byte2);
            }

            //lunxunsend(byte1, byte2);
            System.Threading.Thread.Sleep(1000);

            byte[] buff_schlus = new byte[4] { 0xFA, 0x08, id_high, id_low };
            byte[] crc_buff_1 = new byte[2] { 0x00, 0x00 };
            buff_schlus = buff_schlus.Concat(crc_buff).ToArray();
            ConCRC(buff_schlus, 4);
            this.SendData(buff_schlus);
            //lunxunsend(byte1, byte2);
            //System.Threading.Thread.Sleep(5000);
            metroTextBox2.Text = "OK";
            }));
        }

        //内嵌应用打开//resmon.exe mspaint.exe control.exe calc.exe
        const string path_notepad = "C:/Windows/SysWOW64/notepad.exe";
        const string path_cal = "C:/Windows/SysWOW64/mspaint.exe";
        const string path_cmd = "C:/Windows/SysWOW64/cmd.exe";
        const string path_device = "C:/Windows/SysWOW64/calc.exe";
        Process p = new Process();
        private static int Process_flag=0;

        private void openexe_button_Click(object sender, EventArgs e)
        {
            
            string p_path = "";
            
            if (metroComboBox3.SelectedItem == null)
            {
                return;
            }
            else
            {
                switch(Process_flag)
                {
                    case 1:
                        appbox.Stop();
                        break;
                    case 2:
                        if (p.HasExited == false)
                        {
                            p.Kill();
                        }

                        break;
                }

                switch (metroComboBox3.SelectedItem.ToString())
                {
                    case "记事本":
                        p_path = path_notepad;
                        appbox.AppFilename = p_path;
                        appbox.Start();
                        Process_flag = 1;
                        break;
                    case "画图":
                        p_path = path_cal;
                        appbox.AppFilename = p_path;
                        appbox.Start();
                        Process_flag = 1;
                        break;
                    case "命令行":
                        p_path = path_cmd;
                        p = new Process();
                        p.StartInfo.FileName = p_path;
                        p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                        p.Start();
                        Process_flag = 2;

                        while (p.MainWindowHandle.ToInt32() == 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }
                        SetParent(p.MainWindowHandle, this.metroPanel4.Handle);
                        ShowWindow(p.MainWindowHandle, (int)ProcessWindowStyle.Maximized);
                        break;
                    case "计算器":
                        p_path = path_device;
                        appbox.AppFilename = p_path;
                        appbox.Start();
                        Process_flag = 1;
                        break;
                }
                
               
            }




        }

        private void addexe_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "可执行文件(exe,)|*exe;|exe 可执行文件(*.exe)|*.exe|所有文件(*.*)|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                appbox.AppFilename = openFile.FileName;
                appbox.Start();
            }
        }

        private void 命令行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("cmd.exe");
        }

        private void 记事本ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe");
        }

        private void 画图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("mspaint.exe");
        }

        private void 计算器ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("calc.exe");
        }

        private void 其他工具ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "文本文件(*.txt)|*.txt";

            try
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    filePath = ofd.FileName;
                    StreamReader sr = new StreamReader(ofd.FileName, System.Text.Encoding.Default);
                    this.rev_text_box.Clear();
                    this.rev_text_box.Text = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private string filePath = "";
        private void Save()
        {
            if (filePath.Equals("") == true)
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "文本文件(*.txt)|*.txt";
                if(saveFile.ShowDialog()==DialogResult.OK)
                {
                    filePath = saveFile.FileName;
                    StreamWriter sw = new StreamWriter(filePath, false);
                    sw.Write(rev_text_box.Text);
                    sw.Close();
                }
                
            }
            else
            {
                StreamWriter sw = new StreamWriter(filePath, false);
                sw.Write(rev_text_box.Text);
                sw.Close();
            }
            
        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rev_text_box.Text != string.Empty)
            {


                if (MessageBox.Show("是否保存当前文件？", "温馨提示", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {

                    Save();
                }
            }
            
        }

        private void 另存为ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "文本文件(*.txt)|*.txt";
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter sw = new StreamWriter(saveFile.FileName, false);
                    sw.Write(rev_text_box.Text);
                    sw.Close();
                }

        }

        private void 开发者ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("商务软件编程课程设计\n开发者：洪刘凯","开发者", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private static string gongshi_text = "";
        private static string[] gongshi_list;
        const string Filename = "公式.txt";
        private void 导入分析公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
            
        }
        private static bool fenxi_flag = false;
        private void 启用自动分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if(fenxi_flag==false)
            {
                try
                {

                    StreamReader sr = new StreamReader(Filename, System.Text.Encoding.Default);
                    gongshi_text = sr.ReadToEnd();
                    sr.Close();
                    gongshi_list = gongshi_text.Split('|');
                    if (gongshi_list.Length > 0)
                    {
                        fenxi_flag = true;
                        启用自动分析ToolStripMenuItem.Text = "关闭自动分析";
                    }
                    else
                    {
                        MessageBox.Show("请导入公式");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                
                
            }
            else
            {
                fenxi_flag = false;
                启用自动分析ToolStripMenuItem.Text = "启动自动分析";
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if(rev_flag==true)
            {
                rev_flag = false;
                this.metroButton4.Text = "继续";
            }
            else
            {
                rev_flag = true;
                this.metroButton4.Text = "暂停";
            }
        }
    }
}
