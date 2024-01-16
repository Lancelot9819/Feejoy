using feejoy_wojie.database;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.Text;
using System.Timers;
using System.Collections.Generic;

namespace feejoy_wojie.subform
{

    public partial class monitor : DevExpress.XtraEditors.XtraForm
    {
        public int cal_point = 0;
        public bool auto_cal = false;
        public string cal_file = "";
        
        int returncode;
        string[] h_pulse_st = new string[9];
        string[] l_pulse_st = new string[9];
        string[] h_pulse_1 = new string[9];
        string[] l_pulse_1 = new string[9];
        string[] h_pulse_2 = new string[9];
        string[] l_pulse_2 = new string[9];
        string[] h_pulse_3 = new string[9];
        string[] l_pulse_3 = new string[9];
        string[] h_pulse_4 = new string[9];
        string[] l_pulse_4 = new string[9];
        string[] h_pulse_5 = new string[9];
        string[] l_pulse_5 = new string[9];
        string[] h_pulse_6 = new string[9];
        string[] l_pulse_6 = new string[9];
        string[] h_pulse_7 = new string[9];
        string[] l_pulse_7 = new string[9];
        string[] h_pulse_8 = new string[9];
        string[] l_pulse_8 = new string[9];
        string[] h_pulse_9 = new string[9];
        string[] l_pulse_9 = new string[9];
        string[] h_pulse_10 = new string[9];
        string[] l_pulse_10 = new string[9];
        string[] h_pulse_11 = new string[9];
        string[] l_pulse_11 = new string[9];
        string[] h_pulse_12 = new string[9];
        string[] l_pulse_12 = new string[9];
        string[] h_pulse_13 = new string[9];
        string[] l_pulse_13 = new string[9];
        string[] h_pulse_14 = new string[9];
        string[] l_pulse_14 = new string[9];
        string[] h_pulse_15 = new string[9];
        string[] l_pulse_15 = new string[9];


        public monitor()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private int PLC_connect_DN25_1()
        {
            try
            {
                DN25_1.ActLogicalStationNumber = 25;
                DN25_1.ActPassword = "";
                returncode = DN25_1.Open();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private int PLC_connect_DN25_2()
        {
            try
            {
                DN25_2.ActLogicalStationNumber = 26;
                DN25_2.ActPassword = "";
                returncode = DN25_2.Open();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private int PLC_connect_DN25_3()
        {
            try
            {
                DN25_3.ActLogicalStationNumber = 27;
                DN25_3.ActPassword = "";
                returncode = DN25_3.Open();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private int PLC_disconnect_DN25_1()
        {
            try
            {
                returncode = DN25_1.Close();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private int PLC_disconnect_DN25_2()
        {
            try
            {
                returncode = DN25_2.Close();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private int PLC_disconnect_DN25_3()
        {
            try
            {
                returncode = DN25_3.Close();
            }
            catch
            {
            }
            if (returncode == 0)
            {
                return returncode;
            }
            else
            {
                return -1;
            }
        }

        private string Read_M1(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            int[] iData = new int[length];
            PLC_connect_DN25_1();
            iReturnCode = DN25_1.ReadDeviceRandom(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_1();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        private string Read_D1(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            short[] iData = new short[length];
            PLC_connect_DN25_1();
            iReturnCode = DN25_1.ReadDeviceBlock2(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_1();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        public string Write_M1(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            int[] iData = new int[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int32.Parse(strArray[i]);
            }
            PLC_connect_DN25_1();
            iReturnCode = DN25_1.WriteDeviceRandom(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_1();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Write_D1(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            short[] iData = new short[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int16.Parse(strArray[i]);
            }
            PLC_connect_DN25_1();
            iReturnCode = DN25_1.WriteDeviceBlock2(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_1();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Read_M2(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            int[] iData = new int[length];
            PLC_connect_DN25_2();
            iReturnCode = DN25_2.ReadDeviceRandom(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_2();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        private string Read_D2(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            short[] iData = new short[length];
            PLC_connect_DN25_2();
            iReturnCode = DN25_2.ReadDeviceBlock2(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_2();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        public string Write_M2(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            int[] iData = new int[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int32.Parse(strArray[i]);
            }
            PLC_connect_DN25_2();
            iReturnCode = DN25_2.WriteDeviceRandom(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_2();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Write_D2(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            short[] iData = new short[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int16.Parse(strArray[i]);
            }
            PLC_connect_DN25_2();
            iReturnCode = DN25_2.WriteDeviceBlock2(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_2();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Read_M3(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            int[] iData = new int[length];
            PLC_connect_DN25_3();
            iReturnCode = DN25_3.ReadDeviceRandom(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_3();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        private string Read_D3(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            short[] iData = new short[length];
            PLC_connect_DN25_3();
            iReturnCode = DN25_3.ReadDeviceBlock2(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN25_3();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        public string Write_M3(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            int[] iData = new int[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int32.Parse(strArray[i]);
            }
            PLC_connect_DN25_3();
            iReturnCode = DN25_3.WriteDeviceRandom(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_3();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Write_D3(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            short[] iData = new short[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int16.Parse(strArray[i]);
            }
            PLC_connect_DN25_3();
            iReturnCode = DN25_3.WriteDeviceBlock2(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN25_3();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private void btn_cleardata_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否清空" + tabControl1.SelectedTab.Text + "标定数据?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (tabControl1.SelectedIndex == 0)
                {
                    listView1.Items.Clear();
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    listView2.Items.Clear();
                }
            }
        }

        private void btn_testdata_Click(object sender, EventArgs e)
        {
            //获取执行的方案在整体序列中的索引值
            string[] lv_time = new string[15];
            int[] exec_index = new int[15];
            bool[] execing = { plan_data.exec1, plan_data.exec2, plan_data.exec3, plan_data.exec4, plan_data.exec5,
                               plan_data.exec6, plan_data.exec7, plan_data.exec8, plan_data.exec9, plan_data.exec10,
                               plan_data.exec11,plan_data.exec12,plan_data.exec13,plan_data.exec14,plan_data.exec15 };

            for (int j = 0; j < execing.Length; j++)
            {
                if (execing[j] == true)
                {
                    exec_index[j] = j;
                }
                else
                {
                    exec_index[j] = -1;
                }
            }

            int numtoremove = -1;
            exec_index = exec_index.Where(Val => Val != numtoremove).ToArray();


            double[] b_time = new double[5] { Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10 };
            string[] s_time = new string[9] { b_time[1].ToString(), b_time[1].ToString(), b_time[1].ToString(), b_time[2].ToString(), b_time[2].ToString(), b_time[2].ToString(), b_time[3].ToString(), b_time[3].ToString(), b_time[3].ToString() };
            string[] stp = new string[9] { tb_stpulse1.Text, tb_stpulse2.Text, tb_stpulse3.Text, tb_stpulse4.Text, tb_stpulse5.Text, tb_stpulse6.Text, tb_stpulse7.Text, tb_stpulse8.Text, tb_stpulse9.Text };
            string[] b1p = new string[9] { tb_b1pulse1.Text, tb_b1pulse2.Text, tb_b1pulse3.Text, tb_b1pulse4.Text, tb_b1pulse5.Text, tb_b1pulse6.Text, tb_b1pulse7.Text, tb_b1pulse8.Text, tb_b1pulse9.Text };
            string[] b2p = new string[9] { tb_b2pulse1.Text, tb_b2pulse2.Text, tb_b2pulse3.Text, tb_b2pulse4.Text, tb_b2pulse5.Text, tb_b2pulse6.Text, tb_b2pulse7.Text, tb_b2pulse8.Text, tb_b2pulse9.Text };
            string[] b3p = new string[9] { tb_b3pulse1.Text, tb_b3pulse2.Text, tb_b3pulse3.Text, tb_b3pulse4.Text, tb_b3pulse5.Text, tb_b3pulse6.Text, tb_b3pulse7.Text, tb_b3pulse8.Text, tb_b3pulse9.Text };
            string[] b4p = new string[9] { tb_b4pulse1.Text, tb_b4pulse2.Text, tb_b4pulse3.Text, tb_b4pulse4.Text, tb_b4pulse5.Text, tb_b4pulse6.Text, tb_b4pulse7.Text, tb_b4pulse8.Text, tb_b4pulse9.Text };
            string[] b5p = new string[9] { tb_b5pulse1.Text, tb_b5pulse2.Text, tb_b5pulse3.Text, tb_b5pulse4.Text, tb_b5pulse5.Text, tb_b5pulse6.Text, tb_b5pulse7.Text, tb_b5pulse8.Text, tb_b5pulse9.Text };
            string[] b6p = new string[9] { tb_b6pulse1.Text, tb_b6pulse2.Text, tb_b6pulse3.Text, tb_b6pulse4.Text, tb_b6pulse5.Text, tb_b6pulse6.Text, tb_b6pulse7.Text, tb_b6pulse8.Text, tb_b6pulse9.Text };
            string[] b7p = new string[9] { tb_b7pulse1.Text, tb_b7pulse2.Text, tb_b7pulse3.Text, tb_b7pulse4.Text, tb_b7pulse5.Text, tb_b7pulse6.Text, tb_b7pulse7.Text, tb_b7pulse8.Text, tb_b7pulse9.Text };
            string[] b8p = new string[9] { tb_b8pulse1.Text, tb_b8pulse2.Text, tb_b8pulse3.Text, tb_b8pulse4.Text, tb_b8pulse5.Text, tb_b8pulse6.Text, tb_b8pulse7.Text, tb_b8pulse8.Text, tb_b8pulse9.Text };
            string[] b9p = new string[9] { tb_b9pulse1.Text, tb_b9pulse2.Text, tb_b9pulse3.Text, tb_b9pulse4.Text, tb_b9pulse5.Text, tb_b9pulse6.Text, tb_b9pulse7.Text, tb_b9pulse8.Text, tb_b9pulse9.Text };
            string[] b10p = new string[9] { tb_b10pulse1.Text, tb_b10pulse2.Text, tb_b10pulse3.Text, tb_b10pulse4.Text, tb_b10pulse5.Text, tb_b10pulse6.Text, tb_b10pulse7.Text, tb_b10pulse8.Text, tb_b10pulse9.Text };
            string[] b11p = new string[9] { tb_b11pulse1.Text, tb_b11pulse2.Text, tb_b11pulse3.Text, tb_b11pulse4.Text, tb_b11pulse5.Text, tb_b11pulse6.Text, tb_b11pulse7.Text, tb_b11pulse8.Text, tb_b11pulse9.Text };
            string[] b12p = new string[9] { tb_b12pulse1.Text, tb_b12pulse2.Text, tb_b12pulse3.Text, tb_b12pulse4.Text, tb_b12pulse5.Text, tb_b12pulse6.Text, tb_b12pulse7.Text, tb_b12pulse8.Text, tb_b12pulse9.Text };
            string[] b13p = new string[9] { tb_b13pulse1.Text, tb_b13pulse2.Text, tb_b13pulse3.Text, tb_b13pulse4.Text, tb_b13pulse5.Text, tb_b13pulse6.Text, tb_b13pulse7.Text, tb_b13pulse8.Text, tb_b13pulse9.Text };
            string[] b14p = new string[9] { tb_b14pulse1.Text, tb_b14pulse2.Text, tb_b14pulse3.Text, tb_b14pulse4.Text, tb_b14pulse5.Text, tb_b14pulse6.Text, tb_b14pulse7.Text, tb_b14pulse8.Text, tb_b14pulse9.Text };
            string[] b15p = new string[9] { tb_b15pulse1.Text, tb_b15pulse2.Text, tb_b15pulse3.Text, tb_b15pulse4.Text, tb_b15pulse5.Text, tb_b15pulse6.Text, tb_b15pulse7.Text, tb_b15pulse8.Text, tb_b15pulse9.Text };
            if (tabControl1.SelectedIndex == 0)      //标准表法
            {
                for (int i = 0; i < 9; i++)
                {
                    ListViewItem hang = new ListViewItem();
                    hang.Text = (i + 1).ToString();
                    hang.SubItems.Add(plan_data.operatorname);
                    hang.SubItems.Add(b_time[0].ToString());
                    hang.SubItems.Add(stp[i]);
                    hang.SubItems.Add(b1p[i]);
                    hang.SubItems.Add(b2p[i]);
                    hang.SubItems.Add(b3p[i]);
                    hang.SubItems.Add(b4p[i]);
                    hang.SubItems.Add(b5p[i]);
                    hang.SubItems.Add(b6p[i]);
                    hang.SubItems.Add(b7p[i]);
                    hang.SubItems.Add(b8p[i]);
                    hang.SubItems.Add(b9p[i]);
                    hang.SubItems.Add(b10p[i]);
                    hang.SubItems.Add(b11p[i]);
                    hang.SubItems.Add(b12p[i]);
                    hang.SubItems.Add(b13p[i]);
                    hang.SubItems.Add(b14p[i]);
                    hang.SubItems.Add(b15p[i]);
                    listView1.Items.Add(hang);
                }
            }

            if (tabControl1.SelectedIndex == 1)      //称重法
            {

            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.SelectedItems.Count; i++)
            {
                tb_id.Text = (listView1.SelectedItems[i].Index + 1).ToString();
            }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < listView2.SelectedItems.Count; i++)
            {
                tb_id.Text = (listView2.SelectedItems[i].Index + 1).ToString();
            }
        }

        private void monitor_Load(object sender, EventArgs e)
        {
            cal_point = 0;

            tb_settingfreq.Text = "10";
            tb_settingdoorvalue.Text = "100";

            tb_b1order.Text = "feejoy001";
            tb_b2order.Text = "feejoy002";
            tb_b3order.Text = "feejoy003";
            tb_b4order.Text = "feejoy004";
            tb_b5order.Text = "feejoy005";
            tb_b6order.Text = "feejoy006";
            tb_b7order.Text = "feejoy007";
            tb_b8order.Text = "feejoy008";
            tb_b9order.Text = "feejoy009";
            tb_b10order.Text = "feejoy010";
            tb_b11order.Text = "feejoy011";
            tb_b12order.Text = "feejoy012";
            tb_b13order.Text = "feejoy013";
            tb_b14order.Text = "feejoy014";
            tb_b15order.Text = "feejoy015";

            tb_operator.Text = "高启强";
        }

        private void btn_manualstart_Click(object sender, EventArgs e)
        {
            Write_M1("M23", "1");
            if (cal_point == 0)
            {
                timer1.Enabled = true;
            }
            if (cal_point == 3)
            {
                timer1.Enabled = false;
                timer3.Enabled = true;
            }
            if (cal_point == 6)
            {
                timer3.Enabled = false;
                timer4.Enabled = true;
            }
        }

        private void btn_calstop_Click(object sender, EventArgs e)
        {
            Write_M1("M12", "0");
            timer1.Enabled = false;
            timer2.Enabled = false;
            timer3.Enabled = false;
            timer4.Enabled = false;
        }

        private void btn_saveoperator_Click(object sender, EventArgs e)
        {
            plan_data.operatorname = tb_operator.Text;
            tb_operator.Enabled = false;
            btn_saveoperator.Enabled = false;
        }

        private void btn_resetoperatorname_Click(object sender, EventArgs e)
        {
            plan_data.operatorname = "";
            tb_operator.Enabled = true;
            btn_saveoperator.Enabled = true;
        }

        private void btn_saveorder_Click(object sender, EventArgs e)
        {
            plan_data.b1order = tb_b1order.Text;
            plan_data.b2order = tb_b2order.Text;
            plan_data.b3order = tb_b3order.Text;
            plan_data.b4order = tb_b4order.Text;
            plan_data.b5order = tb_b5order.Text;
            plan_data.b6order = tb_b6order.Text;
            plan_data.b7order = tb_b7order.Text;
            plan_data.b8order = tb_b8order.Text;
            plan_data.b9order = tb_b9order.Text;
            plan_data.b10order = tb_b10order.Text;
            plan_data.b11order = tb_b11order.Text;
            plan_data.b12order = tb_b12order.Text;
            plan_data.b13order = tb_b13order.Text;
            plan_data.b14order = tb_b14order.Text;
            plan_data.b15order = tb_b15order.Text;

            tb_b1order.Enabled = false;
            tb_b2order.Enabled = false;
            tb_b3order.Enabled = false;
            tb_b4order.Enabled = false;
            tb_b5order.Enabled = false;
            tb_b6order.Enabled = false;
            tb_b7order.Enabled = false;
            tb_b8order.Enabled = false;
            tb_b9order.Enabled = false;
            tb_b10order.Enabled = false;
            tb_b11order.Enabled = false;
            tb_b12order.Enabled = false;
            tb_b13order.Enabled = false;
            tb_b14order.Enabled = false;
            tb_b15order.Enabled = false;
            btn_saveorder.Enabled = false;

        }

        private void btn_resetorder_Click(object sender, EventArgs e)
        {
            tb_b1order.Enabled = true;
            tb_b2order.Enabled = true;
            tb_b3order.Enabled = true;
            tb_b4order.Enabled = true;
            tb_b5order.Enabled = true;
            tb_b6order.Enabled = true;
            tb_b7order.Enabled = true;
            tb_b8order.Enabled = true;
            tb_b9order.Enabled = true;
            tb_b10order.Enabled = true;
            tb_b11order.Enabled = true;
            tb_b12order.Enabled = true;
            tb_b13order.Enabled = true;
            tb_b14order.Enabled = true;
            tb_b15order.Enabled = true;
            btn_saveorder.Enabled = true;
        }

        private float cell_float(double cellorivalue, int valuelen)
        {
            //float value = Convert.ToSingle(cellorivalue.ToString().Substring(0, valuelen));
            float value = Convert.ToSingle(cellorivalue.ToString());
            return value;
        }

        private double get_pulse1(string DH, string DL)
        {
            double pulse;
            DH = Read_D1(DH);
            DL = Read_D1(DL);
            if (DL.Contains("-"))
            {
                pulse = (65536 - Convert.ToInt16(DL.Remove(0, 1))) + Convert.ToInt16(DH) * 65536;
            }
            else
            {
                pulse = Convert.ToInt16(DL) + Convert.ToInt16(DH) * 65536;
            }
            return pulse;
        }

        private double get_pulse2(string DH, string DL)
        {
            double pulse;
            DH = Read_D2(DH);
            DL = Read_D2(DL);
            if (DL.Contains("-"))
            {
                pulse = (65536 - Convert.ToInt16(DL.Remove(0, 1))) + Convert.ToInt16(DH) * 65536;
            }
            else
            {
                pulse = Convert.ToInt16(DL) + Convert.ToInt16(DH) * 65536;
            }
            return pulse;
        }

        private void btn_saveorigindata_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = null;
            string filename = @"C:\DN25\DN25模板文件.xlsx";
            string firsttime = System.DateTime.Now.ToString().Replace(":", "-");
            string savetime = firsttime.Replace("/", "-");
            string storename = @"C:\DN25\DN25 " + savetime + ".xlsx";
            cal_file = storename;
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(fs);
            }

            ISheet sheet1 = workbook.GetSheet("被检表1");
            ISheet sheet2 = workbook.GetSheet("被检表2");
            ISheet sheet3 = workbook.GetSheet("被检表3");
            ISheet sheet4 = workbook.GetSheet("被检表4");
            ISheet sheet5 = workbook.GetSheet("被检表5");
            ISheet sheet6 = workbook.GetSheet("被检表6");
            ISheet sheet7 = workbook.GetSheet("被检表7");
            ISheet sheet8 = workbook.GetSheet("被检表8");
            ISheet sheet9 = workbook.GetSheet("被检表9");
            ISheet sheet10 = workbook.GetSheet("被检表10");
            ISheet sheet11 = workbook.GetSheet("被检表11");
            ISheet sheet12 = workbook.GetSheet("被检表12");
            ISheet sheet13 = workbook.GetSheet("被检表13");
            ISheet sheet14 = workbook.GetSheet("被检表14");
            ISheet sheet15 = workbook.GetSheet("被检表15");

            double[] b_time = new double[5] { Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10 };


            //标定时间
            for (int y = 12; y < 21; y = y + 3)
            {
                if (y < 15)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet2.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet2.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet2.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet3.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet3.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet3.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet4.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet4.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet4.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet5.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet5.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet5.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet6.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet6.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet6.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet7.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet7.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet7.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet8.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet8.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet8.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet9.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet9.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet9.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet10.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet10.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet10.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet11.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet11.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet11.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet12.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet12.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet12.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet13.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet13.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet13.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet14.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet14.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet14.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);

                    sheet15.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet15.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet15.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);
                }

                if (y >= 15 && y < 18)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet2.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet2.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet2.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet3.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet3.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet3.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet4.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet4.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet4.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet5.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet5.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet5.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet6.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet6.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet6.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet7.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet7.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet7.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet8.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet8.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet8.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet9.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet9.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet9.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet10.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet10.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet10.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet11.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet11.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet11.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet12.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet12.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet12.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet13.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet13.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet13.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet14.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet14.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet14.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);

                    sheet15.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet15.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet15.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);
                }

                if (y >= 18 && y < 21)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet2.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet2.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet2.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet3.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet3.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet3.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet4.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet4.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet4.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet5.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet5.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet5.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet6.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet6.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet6.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet7.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet7.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet7.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet8.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet8.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet8.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet9.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet9.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet9.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet10.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet10.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet10.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet11.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet11.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet11.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet12.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet12.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet12.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet13.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet13.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet13.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet14.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet14.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet14.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

                    sheet15.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet15.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet15.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);
                }
            }

            //标准被检
            for (int y = 12; y < 21; y = y + 3)
            {
                if (y < 15)
                {
                    sheet1.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet1.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet1.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet2.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet2.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet2.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet3.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet3.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet3.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet4.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet4.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet4.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet5.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet5.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet5.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet6.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet6.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet6.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet7.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet7.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet7.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet8.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet8.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet8.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet9.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet9.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet9.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet10.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet10.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet10.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet11.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet11.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet11.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet12.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet12.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet12.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet13.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet13.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet13.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet14.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet14.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet14.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));
                    sheet15.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse1.Text));
                    sheet15.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse2.Text));
                    sheet15.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse3.Text));

                    sheet1.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse1.Text));
                    sheet1.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse2.Text));
                    sheet1.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse3.Text));
                    sheet2.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse1.Text));
                    sheet2.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse2.Text));
                    sheet2.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse3.Text));
                    sheet3.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse1.Text));
                    sheet3.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse2.Text));
                    sheet3.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse3.Text));
                    sheet4.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse1.Text));
                    sheet4.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse2.Text));
                    sheet4.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse3.Text));
                    sheet5.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse1.Text));
                    sheet5.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse2.Text));
                    sheet5.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse3.Text));
                    sheet6.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse1.Text));
                    sheet6.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse2.Text));
                    sheet6.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse3.Text));
                    sheet7.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse1.Text));
                    sheet7.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse2.Text));
                    sheet7.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse3.Text));
                    sheet8.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse1.Text));
                    sheet8.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse2.Text));
                    sheet8.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse3.Text));
                    sheet9.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse1.Text));
                    sheet9.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse2.Text));
                    sheet9.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse3.Text));
                    sheet10.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse1.Text));
                    sheet10.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse2.Text));
                    sheet10.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse3.Text));
                    sheet11.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse1.Text));
                    sheet11.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse2.Text));
                    sheet11.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse3.Text));
                    sheet12.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse1.Text));
                    sheet12.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse2.Text));
                    sheet12.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse3.Text));
                    sheet13.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse1.Text));
                    sheet13.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse2.Text));
                    sheet13.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse3.Text));
                    sheet14.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse1.Text));
                    sheet14.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse2.Text));
                    sheet14.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse3.Text));
                    sheet15.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse1.Text));
                    sheet15.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse2.Text));
                    sheet15.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse3.Text));
                }

                if (y >= 15 && y < 18)
                {
                    sheet1.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet1.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet1.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet2.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet2.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet2.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet3.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet3.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet3.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet4.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet4.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet4.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet5.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet5.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet5.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet6.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet6.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet6.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet7.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet7.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet7.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet8.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet8.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet8.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet9.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet9.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet9.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet10.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet10.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet10.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet11.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet11.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet11.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet12.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet12.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet12.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet13.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet13.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet13.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet14.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet14.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet14.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                    sheet15.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet15.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet15.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));

                    sheet1.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse4.Text));
                    sheet1.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse5.Text));
                    sheet1.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse6.Text));
                    sheet2.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse4.Text));
                    sheet2.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse5.Text));
                    sheet2.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse6.Text));
                    sheet3.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse4.Text));
                    sheet3.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse5.Text));
                    sheet3.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse6.Text));
                    sheet4.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse4.Text));
                    sheet4.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse5.Text));
                    sheet4.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse6.Text));
                    sheet5.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse4.Text));
                    sheet5.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse5.Text));
                    sheet5.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse6.Text));
                    sheet6.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse4.Text));
                    sheet6.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse5.Text));
                    sheet6.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse6.Text));
                    sheet7.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse4.Text));
                    sheet7.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse5.Text));
                    sheet7.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse6.Text));
                    sheet8.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse4.Text));
                    sheet8.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse5.Text));
                    sheet8.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse6.Text));
                    sheet9.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse4.Text));
                    sheet9.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse5.Text));
                    sheet9.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse6.Text));
                    sheet10.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse4.Text));
                    sheet10.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse5.Text));
                    sheet10.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse6.Text));
                    sheet11.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse4.Text));
                    sheet11.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse5.Text));
                    sheet11.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse6.Text));
                    sheet12.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse4.Text));
                    sheet12.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse5.Text));
                    sheet12.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse6.Text));
                    sheet13.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse4.Text));
                    sheet13.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse5.Text));
                    sheet13.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse6.Text));
                    sheet14.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse4.Text));
                    sheet14.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse5.Text));
                    sheet14.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse6.Text));
                    sheet15.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse4.Text));
                    sheet15.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse5.Text));
                    sheet15.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse6.Text));
                }

                if (y >= 18 && y < 21)
                {
                    sheet1.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet1.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet1.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet2.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet2.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet2.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet3.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet3.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet3.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet4.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet4.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet4.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet5.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet5.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet5.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet6.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet6.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet6.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet7.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet7.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet7.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet8.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet8.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet8.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet9.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet9.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet9.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet10.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet10.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet10.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet11.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet11.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet11.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet12.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet12.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet12.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet13.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet13.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet13.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet14.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet14.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet14.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                    sheet15.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet15.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet15.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));

                    sheet1.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse7.Text));
                    sheet1.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse8.Text));
                    sheet1.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b1pulse9.Text));
                    sheet2.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse7.Text));
                    sheet2.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse8.Text));
                    sheet2.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b2pulse9.Text));
                    sheet3.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse7.Text));
                    sheet3.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse8.Text));
                    sheet3.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b3pulse9.Text));
                    sheet4.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse7.Text));
                    sheet4.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse8.Text));
                    sheet4.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b4pulse9.Text));
                    sheet5.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse7.Text));
                    sheet5.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse8.Text));
                    sheet5.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b5pulse9.Text));
                    sheet6.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse7.Text));
                    sheet6.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse8.Text));
                    sheet6.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b6pulse9.Text));
                    sheet7.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse7.Text));
                    sheet7.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse8.Text));
                    sheet7.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b7pulse9.Text));
                    sheet8.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse7.Text));
                    sheet8.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse8.Text));
                    sheet8.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b8pulse9.Text));
                    sheet9.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse7.Text));
                    sheet9.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse8.Text));
                    sheet9.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b9pulse9.Text));
                    sheet10.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse7.Text));
                    sheet10.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse8.Text));
                    sheet10.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b10pulse9.Text));
                    sheet11.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse7.Text));
                    sheet11.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse8.Text));
                    sheet11.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b11pulse9.Text));
                    sheet12.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse7.Text));
                    sheet12.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse8.Text));
                    sheet12.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b12pulse9.Text));
                    sheet13.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse7.Text));
                    sheet13.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse8.Text));
                    sheet13.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b13pulse9.Text));
                    sheet14.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse7.Text));
                    sheet14.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse8.Text));
                    sheet14.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b14pulse9.Text));
                    sheet15.GetRow(y).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse7.Text));
                    sheet15.GetRow(y + 1).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse8.Text));
                    sheet15.GetRow(y + 2).GetCell(22).SetCellValue(Convert.ToDouble(tb_b15pulse9.Text));
                }
            }

            sheet1.ForceFormulaRecalculation = true;
            sheet2.ForceFormulaRecalculation = true;
            sheet3.ForceFormulaRecalculation = true;
            sheet4.ForceFormulaRecalculation = true;
            sheet5.ForceFormulaRecalculation = true;
            sheet6.ForceFormulaRecalculation = true;
            sheet7.ForceFormulaRecalculation = true;
            sheet8.ForceFormulaRecalculation = true;
            sheet9.ForceFormulaRecalculation = true;
            sheet10.ForceFormulaRecalculation = true;
            sheet11.ForceFormulaRecalculation = true;
            sheet12.ForceFormulaRecalculation = true;
            sheet13.ForceFormulaRecalculation = true;
            sheet14.ForceFormulaRecalculation = true;
            sheet15.ForceFormulaRecalculation = true;

            XSSFFormulaEvaluator dio = new XSSFFormulaEvaluator(workbook);
            plan_data.b1_flow1 = cell_float(dio.EvaluateInCell(sheet1.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b1_flow2 = cell_float(dio.EvaluateInCell(sheet1.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b1_flow3 = cell_float(dio.EvaluateInCell(sheet1.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b1_k1 = cell_float(dio.EvaluateInCell(sheet1.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b1_k2 = cell_float(dio.EvaluateInCell(sheet1.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b1_k3 = cell_float(dio.EvaluateInCell(sheet1.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b2_flow1 = cell_float(dio.EvaluateInCell(sheet2.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b2_flow2 = cell_float(dio.EvaluateInCell(sheet2.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b2_flow3 = cell_float(dio.EvaluateInCell(sheet2.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b2_k1 = cell_float(dio.EvaluateInCell(sheet2.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b2_k2 = cell_float(dio.EvaluateInCell(sheet2.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b2_k3 = cell_float(dio.EvaluateInCell(sheet2.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b3_flow1 = cell_float(dio.EvaluateInCell(sheet3.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b3_flow2 = cell_float(dio.EvaluateInCell(sheet3.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b3_flow3 = cell_float(dio.EvaluateInCell(sheet3.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b3_k1 = cell_float(dio.EvaluateInCell(sheet3.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b3_k2 = cell_float(dio.EvaluateInCell(sheet3.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b3_k3 = cell_float(dio.EvaluateInCell(sheet3.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b4_flow1 = cell_float(dio.EvaluateInCell(sheet4.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b4_flow2 = cell_float(dio.EvaluateInCell(sheet4.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b4_flow3 = cell_float(dio.EvaluateInCell(sheet4.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b4_k1 = cell_float(dio.EvaluateInCell(sheet4.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b4_k2 = cell_float(dio.EvaluateInCell(sheet4.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b4_k3 = cell_float(dio.EvaluateInCell(sheet4.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b5_flow1 = cell_float(dio.EvaluateInCell(sheet5.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b5_flow2 = cell_float(dio.EvaluateInCell(sheet5.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b5_flow3 = cell_float(dio.EvaluateInCell(sheet5.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b5_k1 = cell_float(dio.EvaluateInCell(sheet5.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b5_k2 = cell_float(dio.EvaluateInCell(sheet5.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b5_k3 = cell_float(dio.EvaluateInCell(sheet5.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b6_flow1 = cell_float(dio.EvaluateInCell(sheet6.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b6_flow2 = cell_float(dio.EvaluateInCell(sheet6.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b6_flow3 = cell_float(dio.EvaluateInCell(sheet6.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b6_k1 = cell_float(dio.EvaluateInCell(sheet6.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b6_k2 = cell_float(dio.EvaluateInCell(sheet6.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b6_k3 = cell_float(dio.EvaluateInCell(sheet6.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b7_flow1 = cell_float(dio.EvaluateInCell(sheet7.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b7_flow2 = cell_float(dio.EvaluateInCell(sheet7.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b7_flow3 = cell_float(dio.EvaluateInCell(sheet7.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b7_k1 = cell_float(dio.EvaluateInCell(sheet7.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b7_k2 = cell_float(dio.EvaluateInCell(sheet7.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b7_k3 = cell_float(dio.EvaluateInCell(sheet7.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b8_flow1 = cell_float(dio.EvaluateInCell(sheet8.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b8_flow2 = cell_float(dio.EvaluateInCell(sheet8.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b8_flow3 = cell_float(dio.EvaluateInCell(sheet8.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b8_k1 = cell_float(dio.EvaluateInCell(sheet8.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b8_k2 = cell_float(dio.EvaluateInCell(sheet8.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b8_k3 = cell_float(dio.EvaluateInCell(sheet8.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b9_flow1 = cell_float(dio.EvaluateInCell(sheet9.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b9_flow2 = cell_float(dio.EvaluateInCell(sheet9.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b9_flow3 = cell_float(dio.EvaluateInCell(sheet9.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b9_k1 = cell_float(dio.EvaluateInCell(sheet9.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b9_k2 = cell_float(dio.EvaluateInCell(sheet9.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b9_k3 = cell_float(dio.EvaluateInCell(sheet9.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b10_flow1 = cell_float(dio.EvaluateInCell(sheet10.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b10_flow2 = cell_float(dio.EvaluateInCell(sheet10.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b10_flow3 = cell_float(dio.EvaluateInCell(sheet10.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b10_k1 = cell_float(dio.EvaluateInCell(sheet10.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b10_k2 = cell_float(dio.EvaluateInCell(sheet10.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b10_k3 = cell_float(dio.EvaluateInCell(sheet10.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b11_flow1 = cell_float(dio.EvaluateInCell(sheet11.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b11_flow2 = cell_float(dio.EvaluateInCell(sheet11.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b11_flow3 = cell_float(dio.EvaluateInCell(sheet11.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b11_k1 = cell_float(dio.EvaluateInCell(sheet11.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b11_k2 = cell_float(dio.EvaluateInCell(sheet11.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b11_k3 = cell_float(dio.EvaluateInCell(sheet11.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b12_flow1 = cell_float(dio.EvaluateInCell(sheet12.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b12_flow2 = cell_float(dio.EvaluateInCell(sheet12.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b12_flow3 = cell_float(dio.EvaluateInCell(sheet12.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b12_k1 = cell_float(dio.EvaluateInCell(sheet12.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b12_k2 = cell_float(dio.EvaluateInCell(sheet12.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b12_k3 = cell_float(dio.EvaluateInCell(sheet12.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b13_flow1 = cell_float(dio.EvaluateInCell(sheet13.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b13_flow2 = cell_float(dio.EvaluateInCell(sheet13.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b13_flow3 = cell_float(dio.EvaluateInCell(sheet13.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b13_k1 = cell_float(dio.EvaluateInCell(sheet13.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b13_k2 = cell_float(dio.EvaluateInCell(sheet13.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b13_k3 = cell_float(dio.EvaluateInCell(sheet13.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b14_flow1 = cell_float(dio.EvaluateInCell(sheet14.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b14_flow2 = cell_float(dio.EvaluateInCell(sheet14.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b14_flow3 = cell_float(dio.EvaluateInCell(sheet14.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b14_k1 = cell_float(dio.EvaluateInCell(sheet14.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b14_k2 = cell_float(dio.EvaluateInCell(sheet14.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b14_k3 = cell_float(dio.EvaluateInCell(sheet14.GetRow(24).GetCell(12)).NumericCellValue, 12);

            plan_data.b15_flow1 = cell_float(dio.EvaluateInCell(sheet15.GetRow(18).GetCell(1)).NumericCellValue, 4);
            plan_data.b15_flow2 = cell_float(dio.EvaluateInCell(sheet15.GetRow(21).GetCell(1)).NumericCellValue, 4);
            plan_data.b15_flow3 = cell_float(dio.EvaluateInCell(sheet15.GetRow(24).GetCell(1)).NumericCellValue, 4);

            plan_data.b15_k1 = cell_float(dio.EvaluateInCell(sheet15.GetRow(18).GetCell(12)).NumericCellValue, 12);
            plan_data.b15_k2 = cell_float(dio.EvaluateInCell(sheet15.GetRow(21).GetCell(12)).NumericCellValue, 12);
            plan_data.b15_k3 = cell_float(dio.EvaluateInCell(sheet15.GetRow(24).GetCell(12)).NumericCellValue, 12);

            using (FileStream fs_new = new FileStream(storename, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs_new);  
                workbook.Close();
            }
            MessageBox.Show("数据表格已导出。", "提示", MessageBoxButtons.OK);
        }

        private byte[] CRC(byte[] pByte)
        {
            byte[] crc = new byte[2];
            int nBit;
            ushort nShiftedBit;
            ushort pChecksum = 0xFFFF;
            int nNumberOfBytes = pByte.Length - 2;
            for (int nByte = 0; nByte < nNumberOfBytes; nByte++)
            {
                pChecksum ^= pByte[nByte];
                for (nBit = 0; nBit < 8; nBit++)
                {
                    if ((pChecksum & 0x1) == 1)
                    {
                        nShiftedBit = 1;
                    }
                    else
                    {
                        nShiftedBit = 0;
                    }
                    pChecksum >>= 1;
                    if (nShiftedBit != 0)
                    {
                        pChecksum ^= 0xA001;
                    }
                }
            }
            crc[1] = (byte)(pChecksum & 0xFF);
            crc[0] = (byte)((pChecksum & 0xFF00) >> 8);
            return crc;
        }

        private void calibrationstart_Click(object sender, EventArgs e)
        {
            subform.calibratedata calibratedata = new subform.calibratedata();
            calibratedata.ShowDialog();
        }


        private string tentime_D(string D_value)
        {
            string tentime_value = (Convert.ToSingle(D_value) * 10).ToString();
            return tentime_value;
        }


        private void store_pulse(int calpoint)
        {
            h_pulse_st[calpoint] = Read_D1("D52"); l_pulse_st[calpoint] = Read_D1("D51");
            h_pulse_1[calpoint] = Read_D1("D22"); l_pulse_1[calpoint] = Read_D1("D21");
            h_pulse_2[calpoint] = Read_D1("D24"); l_pulse_2[calpoint] = Read_D1("D23");
            h_pulse_3[calpoint] = Read_D1("D26"); l_pulse_3[calpoint] = Read_D1("D25");
            h_pulse_4[calpoint] = Read_D1("D28"); l_pulse_4[calpoint] = Read_D1("D27");
            h_pulse_5[calpoint] = Read_D1("D30"); l_pulse_5[calpoint] = Read_D1("D29");

            h_pulse_6[calpoint] = Read_D2("D32"); l_pulse_6[calpoint] = Read_D2("D31");
            h_pulse_7[calpoint] = Read_D2("D34"); l_pulse_7[calpoint] = Read_D2("D33");
            h_pulse_8[calpoint] = Read_D2("D36"); l_pulse_8[calpoint] = Read_D2("D35");
            h_pulse_9[calpoint] = Read_D2("D38"); l_pulse_9[calpoint] = Read_D2("D37");
            h_pulse_10[calpoint] = Read_D2("D40"); l_pulse_10[calpoint] = Read_D2("D39");
            h_pulse_11[calpoint] = Read_D2("D42"); l_pulse_11[calpoint] = Read_D2("D41");

            h_pulse_12[calpoint] = Read_D3("D44"); l_pulse_12[calpoint] = Read_D3("D43");
            h_pulse_13[calpoint] = Read_D3("D46"); l_pulse_13[calpoint] = Read_D3("D45");
            h_pulse_14[calpoint] = Read_D3("D48"); l_pulse_14[calpoint] = Read_D3("D47");
            h_pulse_15[calpoint] = Read_D3("D50"); l_pulse_15[calpoint] = Read_D3("D49");


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (cal_point == 0)
            {
                Write_D1("D17", plan_data.pl1);
                Write_D1("D16", plan_data.kd1);

                Thread.Sleep(plan_data.stable_time*1000);

                Write_M1("M12", "1");

                cal_point = 1;
            }
            if (Read_M1("M14") == "1")
            {
                switch (cal_point)
                {
                    case 1:
                        store_pulse(0);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 2;
                        Thread.Sleep(1000);

                        break;
                    case 2:
                        store_pulse(1);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 3;
                        Thread.Sleep(1000);

                        break;
                    case 3:
                        store_pulse(2);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        //cal_point = 3;
                        //Thread.Sleep(1000);                        
                        break;


                        //case 1:
                        //    store_pulse(0);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");


                        //    cal_point = 2;
                        //    Thread.Sleep(1000);

                        //    break;
                        //case 2:
                        //    store_pulse(1);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");

                        //    cal_point = 3;
                        //    Thread.Sleep(1000);

                        //    break;
                        //case 3:
                        //    store_pulse(2);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);                 

                        //    Write_D1("D17", "300");
                        //    Write_D1("D16", "700");

                        //    cal_point = 3;
                        //    Thread.Sleep(1000);
                        //    break;
                        //case 4:
                        //    store_pulse(3);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");

                        //    cal_point = 5;

                        //    Thread.Sleep(1000);
                        //    break;
                        //case 5:
                        //    store_pulse(4);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");

                        //    cal_point = 6;

                        //    Thread.Sleep(1000);

                        //    break;
                        //case 6:
                        //    store_pulse(5);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    cal_point = 7;
                        //    Thread.Sleep(1000);

                        //    Write_D1("D17", "150");
                        //    Write_D1("D16", "500");

                        //    Thread.Sleep(10000);
                        //    break;
                        //case 7:
                        //    store_pulse(6);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");

                        //    cal_point = 8;

                        //    Thread.Sleep(1000);

                        //    break;
                        //case 8:
                        //    store_pulse(7);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    Write_M1("M12", "1");

                        //    cal_point = 9;

                        //    Thread.Sleep(1000);


                        //    break;
                        //case 9:
                        //    store_pulse(8);

                        //    tb_D4.Text = cal_point.ToString();

                        //    Thread.Sleep(1000);

                        //    cal_point = 0;

                        //    stop_store();
                        //    break;

                }
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (cal_point == 3)
            {
                Write_D1("D17", plan_data.pl2);
                Write_D1("D16", plan_data.kd2);

                Thread.Sleep(plan_data.stable_time*1000);

                Write_M1("M12", "1");

                cal_point = 4;

            }
            if (Read_M1("M14") == "1")
            {
                switch (cal_point)
                {
                    case 4:
                        store_pulse(3);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 5;
                        Thread.Sleep(1000);

                        break;
                    case 5:
                        store_pulse(4);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 6;
                        Thread.Sleep(1000);

                        break;
                    case 6:
                        store_pulse(5);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        //cal_point = 7;
                        //Thread.Sleep(1000);
                        break;
                }
            }
        }


        private void timer4_Tick(object sender, EventArgs e)
        {
            if (cal_point == 6)
            {
                Write_D1("D17", plan_data.pl3);
                Write_D1("D16", plan_data.kd3);

                Thread.Sleep(plan_data.stable_time*1000);

                Write_M1("M12", "1");

                cal_point = 7;

            }
            if (Read_M1("M14") == "1")
            {
                switch (cal_point)
                {
                    case 7:
                        store_pulse(6);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 8;
                        Thread.Sleep(1000);

                        break;
                    case 8:
                        store_pulse(7);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 9;
                        Thread.Sleep(1000);

                        break;
                    case 9:
                        store_pulse(8);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        cal_point = 0;
                        Thread.Sleep(1000);

                        break;
                }
            }
        }

        private void btn_reset_Click(object sender, EventArgs e)
        {
            Write_M1("M23", "1");
            auto_cal = false;
            cal_point = 0;
            tb_D4.Clear();
            timer1.Enabled = false;
            timer2.Enabled = false;

            tb_stpulse1.Clear();

            tb_b1pulse1.Clear();
            tb_b2pulse1.Clear();
            tb_b3pulse1.Clear();
            tb_b4pulse1.Clear();
            tb_b5pulse1.Clear();

            tb_b6pulse1.Clear();
            tb_b7pulse1.Clear();
            tb_b8pulse1.Clear();
            tb_b9pulse1.Clear();
            tb_b10pulse1.Clear();
            tb_b11pulse1.Clear();

            tb_b12pulse1.Clear();
            tb_b13pulse1.Clear();
            tb_b14pulse1.Clear();
            tb_b15pulse1.Clear();

            tb_stpulse2.Clear();

            tb_b1pulse2.Clear();
            tb_b2pulse2.Clear();
            tb_b3pulse2.Clear();
            tb_b4pulse2.Clear();
            tb_b5pulse2.Clear();

            tb_b6pulse2.Clear();
            tb_b7pulse2.Clear();
            tb_b8pulse2.Clear();
            tb_b9pulse2.Clear();
            tb_b10pulse2.Clear();
            tb_b11pulse2.Clear();

            tb_b12pulse2.Clear();
            tb_b13pulse2.Clear();
            tb_b14pulse2.Clear();
            tb_b15pulse2.Clear();

            tb_stpulse3.Clear();

            tb_b1pulse3.Clear();
            tb_b2pulse3.Clear();
            tb_b3pulse3.Clear();
            tb_b4pulse3.Clear();
            tb_b5pulse3.Clear();

            tb_b6pulse3.Clear();
            tb_b7pulse3.Clear();
            tb_b8pulse3.Clear();
            tb_b9pulse3.Clear();
            tb_b10pulse3.Clear();
            tb_b11pulse3.Clear();

            tb_b12pulse3.Clear();
            tb_b13pulse3.Clear();
            tb_b14pulse3.Clear();
            tb_b15pulse3.Clear();

            tb_stpulse4.Clear();

            tb_b1pulse4.Clear();
            tb_b2pulse4.Clear();
            tb_b3pulse4.Clear();
            tb_b4pulse4.Clear();
            tb_b5pulse4.Clear();

            tb_b6pulse4.Clear();
            tb_b7pulse4.Clear();
            tb_b8pulse4.Clear();
            tb_b9pulse4.Clear();
            tb_b10pulse4.Clear();
            tb_b11pulse4.Clear();

            tb_b12pulse4.Clear();
            tb_b13pulse4.Clear();
            tb_b14pulse4.Clear();
            tb_b15pulse4.Clear();

            tb_stpulse5.Clear();

            tb_b1pulse5.Clear();
            tb_b2pulse5.Clear();
            tb_b3pulse5.Clear();
            tb_b4pulse5.Clear();
            tb_b5pulse5.Clear();

            tb_b6pulse5.Clear();
            tb_b7pulse5.Clear();
            tb_b8pulse5.Clear();
            tb_b9pulse5.Clear();
            tb_b10pulse5.Clear();
            tb_b11pulse5.Clear();

            tb_b12pulse5.Clear();
            tb_b13pulse5.Clear();
            tb_b14pulse5.Clear();
            tb_b15pulse5.Clear();

            tb_stpulse6.Clear();

            tb_b1pulse6.Clear();
            tb_b2pulse6.Clear();
            tb_b3pulse6.Clear();
            tb_b4pulse6.Clear();
            tb_b5pulse6.Clear();

            tb_b6pulse6.Clear();
            tb_b7pulse6.Clear();
            tb_b8pulse6.Clear();
            tb_b9pulse6.Clear();
            tb_b10pulse6.Clear();
            tb_b11pulse6.Clear();

            tb_b12pulse6.Clear();
            tb_b13pulse6.Clear();
            tb_b14pulse6.Clear();
            tb_b15pulse6.Clear();

            tb_stpulse7.Clear();

            tb_b1pulse7.Clear();
            tb_b2pulse7.Clear();
            tb_b3pulse7.Clear();
            tb_b4pulse7.Clear();
            tb_b5pulse7.Clear();

            tb_b6pulse7.Clear();
            tb_b7pulse7.Clear();
            tb_b8pulse7.Clear();
            tb_b9pulse7.Clear();
            tb_b10pulse7.Clear();
            tb_b11pulse7.Clear();

            tb_b12pulse7.Clear();
            tb_b13pulse7.Clear();
            tb_b14pulse7.Clear();
            tb_b15pulse7.Clear();

            tb_stpulse8.Clear();

            tb_b1pulse8.Clear();
            tb_b2pulse8.Clear();
            tb_b3pulse8.Clear();
            tb_b4pulse8.Clear();
            tb_b5pulse8.Clear();

            tb_b6pulse8.Clear();
            tb_b7pulse8.Clear();
            tb_b8pulse8.Clear();
            tb_b9pulse8.Clear();
            tb_b10pulse8.Clear();
            tb_b11pulse8.Clear();

            tb_b12pulse8.Clear();
            tb_b13pulse8.Clear();
            tb_b14pulse8.Clear();
            tb_b15pulse8.Clear();

            tb_stpulse9.Clear();

            tb_b1pulse9.Clear();
            tb_b2pulse9.Clear();
            tb_b3pulse9.Clear();
            tb_b4pulse9.Clear();
            tb_b5pulse9.Clear();

            tb_b6pulse9.Clear();
            tb_b7pulse9.Clear();
            tb_b8pulse9.Clear();
            tb_b9pulse9.Clear();
            tb_b10pulse9.Clear();
            tb_b11pulse9.Clear();

            tb_b12pulse9.Clear();
            tb_b13pulse9.Clear();
            tb_b14pulse9.Clear();
            tb_b15pulse9.Clear();
        }

        private void btn_selfauto_Click(object sender, EventArgs e)
        {
            Write_M1("M11", "0");
            Write_M1("M31", "0");
        }

        private void btn_auto_Click(object sender, EventArgs e)
        {
            Write_M1("M11", "0");
            Write_M1("M31", "0");
        }

        private void btn_pumprun_Click(object sender, EventArgs e)
        {
            Write_M1("M35", "1");
            Write_M1("M35", "0");
        }

        private void btn_pumpstop_Click(object sender, EventArgs e)
        {
            Write_M1("M36", "1");
            Write_M1("M36", "0");
        }

        private void btn_setfreq_Click(object sender, EventArgs e)
        {
            Write_D1("D17", tentime_D(tb_settingfreq.Text));
        }

        private void btn_setkd_Click(object sender, EventArgs e)
        {
            Write_D1("D16", tentime_D(tb_settingdoorvalue.Text));
        }

        private void btn_startauto_Click(object sender, EventArgs e)
        {
            Write_D1("D17", plan_data.pl1);
            Write_D1("D16", plan_data.kd1);

            Thread.Sleep(plan_data.stable_time*1000);

            Write_M1("M12", "1");
            cal_point = cal_point + 1;
            timer2.Enabled = true;
        }

        private double get_pulse(string DH, string DL)
        {
            double pulse;
            if (DL.Contains("-"))
            {
                pulse = (65536 - Convert.ToInt16(DL.Remove(0, 1))) + Convert.ToInt16(DH) * 65536;
            }
            else
            {
                pulse = Convert.ToInt16(DL) + Convert.ToInt16(DH) * 65536;
            }
            return pulse;
        }

        private void stop_store()
        {
            timer2.Enabled = false;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (Read_M1("M14") == "1")
            {
                switch (cal_point)
                {
                    case 1:
                        store_pulse(0);                       

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");


                        cal_point = 2;
                        Thread.Sleep(1000);

                        break;
                    case 2:                        
                        store_pulse(1);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 3;
                        Thread.Sleep(1000);

                        break;
                    case 3:
                        store_pulse(2);

                        tb_D4.Text = cal_point.ToString();

                        Write_D1("D17", plan_data.pl2);
                        Write_D1("D16", plan_data.kd2);

                        Thread.Sleep(plan_data.stable_time*1000);

                        Write_M1("M12", "1");

                        cal_point = 4;
                        Thread.Sleep(1000);
                        break;
                    case 4:
                        store_pulse(3);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 5;

                        Thread.Sleep(1000);
                        break;
                    case 5:
                        store_pulse(4);
                    
                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 6;

                        Thread.Sleep(1000);

                        break;
                    case 6:
                        store_pulse(5);

                        tb_D4.Text = cal_point.ToString();

                        Write_D1("D17", plan_data.pl3);
                        Write_D1("D16", plan_data.kd3);

                        Thread.Sleep(plan_data.stable_time*1000);

                        Write_M1("M12", "1");

                        cal_point = 7;

                        Thread.Sleep(1000);

                        break;
                    case 7:
                        store_pulse(6);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 8;

                        Thread.Sleep(1000);

                        break;
                    case 8:
                        store_pulse(7);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");

                        cal_point = 9;

                        Thread.Sleep(1000);


                        break;
                    case 9:
                        store_pulse(8);

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        cal_point = 0;

                        stop_store();
                        break;
                }
                
            }
        }

        private void btn_exportpulse_Click(object sender, EventArgs e)
        {
                //time1
            tb_stpulse1.Text = get_pulse(h_pulse_st[0], l_pulse_st[0]).ToString();
            tb_b1pulse1.Text = get_pulse(h_pulse_1[0], l_pulse_1[0]).ToString();
            tb_b2pulse1.Text = get_pulse(h_pulse_2[0], l_pulse_2[0]).ToString();
            tb_b3pulse1.Text = get_pulse(h_pulse_3[0], l_pulse_3[0]).ToString();
            tb_b4pulse1.Text = get_pulse(h_pulse_4[0], l_pulse_4[0]).ToString();
            tb_b5pulse1.Text = get_pulse(h_pulse_5[0], l_pulse_5[0]).ToString();

            tb_b6pulse1.Text = get_pulse(h_pulse_6[0], l_pulse_6[0]).ToString();
            tb_b7pulse1.Text = get_pulse(h_pulse_7[0], l_pulse_7[0]).ToString();
            tb_b8pulse1.Text = get_pulse(h_pulse_8[0], l_pulse_8[0]).ToString();
            tb_b9pulse1.Text = get_pulse(h_pulse_9[0], l_pulse_9[0]).ToString();
            tb_b10pulse1.Text = get_pulse(h_pulse_10[0], l_pulse_10[0]).ToString();
            tb_b11pulse1.Text = get_pulse(h_pulse_11[0], l_pulse_11[0]).ToString();

            tb_b12pulse1.Text = get_pulse(h_pulse_12[0], l_pulse_12[0]).ToString();
            tb_b13pulse1.Text = get_pulse(h_pulse_13[0], l_pulse_13[0]).ToString();
            tb_b14pulse1.Text = get_pulse(h_pulse_14[0], l_pulse_14[0]).ToString();
            tb_b15pulse1.Text = get_pulse(h_pulse_15[0], l_pulse_15[0]).ToString();
            //time2
            tb_stpulse2.Text = get_pulse(h_pulse_st[1], l_pulse_st[1]).ToString();
            tb_b1pulse2.Text = get_pulse(h_pulse_1[1], l_pulse_1[1]).ToString();
            tb_b2pulse2.Text = get_pulse(h_pulse_2[1], l_pulse_2[1]).ToString();
            tb_b3pulse2.Text = get_pulse(h_pulse_3[1], l_pulse_3[1]).ToString();
            tb_b4pulse2.Text = get_pulse(h_pulse_4[1], l_pulse_4[1]).ToString();
            tb_b5pulse2.Text = get_pulse(h_pulse_5[1], l_pulse_5[1]).ToString();

            tb_b6pulse2.Text = get_pulse(h_pulse_6[1], l_pulse_6[1]).ToString();
            tb_b7pulse2.Text = get_pulse(h_pulse_7[1], l_pulse_7[1]).ToString();
            tb_b8pulse2.Text = get_pulse(h_pulse_8[1], l_pulse_8[1]).ToString();
            tb_b9pulse2.Text = get_pulse(h_pulse_9[1], l_pulse_9[1]).ToString();
            tb_b10pulse2.Text = get_pulse(h_pulse_10[1], l_pulse_10[1]).ToString();
            tb_b11pulse2.Text = get_pulse(h_pulse_11[1], l_pulse_11[1]).ToString();

            tb_b12pulse2.Text = get_pulse(h_pulse_12[1], l_pulse_12[1]).ToString();
            tb_b13pulse2.Text = get_pulse(h_pulse_13[1], l_pulse_13[1]).ToString();
            tb_b14pulse2.Text = get_pulse(h_pulse_14[1], l_pulse_14[1]).ToString();
            tb_b15pulse2.Text = get_pulse(h_pulse_15[1], l_pulse_15[1]).ToString();
            //time3
            tb_stpulse3.Text = get_pulse(h_pulse_st[2], l_pulse_st[2]).ToString();
            tb_b1pulse3.Text = get_pulse(h_pulse_1[2], l_pulse_1[2]).ToString();
            tb_b2pulse3.Text = get_pulse(h_pulse_2[2], l_pulse_2[2]).ToString();
            tb_b3pulse3.Text = get_pulse(h_pulse_3[2], l_pulse_3[2]).ToString();
            tb_b4pulse3.Text = get_pulse(h_pulse_4[2], l_pulse_4[2]).ToString();
            tb_b5pulse3.Text = get_pulse(h_pulse_5[2], l_pulse_5[2]).ToString();

            tb_b6pulse3.Text = get_pulse(h_pulse_6[2], l_pulse_6[2]).ToString();
            tb_b7pulse3.Text = get_pulse(h_pulse_7[2], l_pulse_7[2]).ToString();
            tb_b8pulse3.Text = get_pulse(h_pulse_8[2], l_pulse_8[2]).ToString();
            tb_b9pulse3.Text = get_pulse(h_pulse_9[2], l_pulse_9[2]).ToString();
            tb_b10pulse3.Text = get_pulse(h_pulse_10[2], l_pulse_10[2]).ToString();
            tb_b11pulse3.Text = get_pulse(h_pulse_11[2], l_pulse_11[2]).ToString();

            tb_b12pulse3.Text = get_pulse(h_pulse_12[2], l_pulse_12[2]).ToString();
            tb_b13pulse3.Text = get_pulse(h_pulse_13[2], l_pulse_13[2]).ToString();
            tb_b14pulse3.Text = get_pulse(h_pulse_14[2], l_pulse_14[2]).ToString();
            tb_b15pulse3.Text = get_pulse(h_pulse_15[2], l_pulse_15[2]).ToString();
            //time4
            tb_stpulse4.Text = get_pulse(h_pulse_st[3], l_pulse_st[3]).ToString();
            tb_b1pulse4.Text = get_pulse(h_pulse_1[3], l_pulse_1[3]).ToString();
            tb_b2pulse4.Text = get_pulse(h_pulse_2[3], l_pulse_2[3]).ToString();
            tb_b3pulse4.Text = get_pulse(h_pulse_3[3], l_pulse_3[3]).ToString();
            tb_b4pulse4.Text = get_pulse(h_pulse_4[3], l_pulse_4[3]).ToString();
            tb_b5pulse4.Text = get_pulse(h_pulse_5[3], l_pulse_5[3]).ToString();

            tb_b6pulse4.Text = get_pulse(h_pulse_6[3], l_pulse_6[3]).ToString();
            tb_b7pulse4.Text = get_pulse(h_pulse_7[3], l_pulse_7[3]).ToString();
            tb_b8pulse4.Text = get_pulse(h_pulse_8[3], l_pulse_8[3]).ToString();
            tb_b9pulse4.Text = get_pulse(h_pulse_9[3], l_pulse_9[3]).ToString();
            tb_b10pulse4.Text = get_pulse(h_pulse_10[3], l_pulse_10[3]).ToString();
            tb_b11pulse4.Text = get_pulse(h_pulse_11[3], l_pulse_11[3]).ToString();

            tb_b12pulse4.Text = get_pulse(h_pulse_12[3], l_pulse_12[3]).ToString();
            tb_b13pulse4.Text = get_pulse(h_pulse_13[3], l_pulse_13[3]).ToString();
            tb_b14pulse4.Text = get_pulse(h_pulse_14[3], l_pulse_14[3]).ToString();
            tb_b15pulse4.Text = get_pulse(h_pulse_15[3], l_pulse_15[3]).ToString();
            //time5
            tb_stpulse5.Text = get_pulse(h_pulse_st[4], l_pulse_st[4]).ToString();
            tb_b1pulse5.Text = get_pulse(h_pulse_1[4], l_pulse_1[4]).ToString();
            tb_b2pulse5.Text = get_pulse(h_pulse_2[4], l_pulse_2[4]).ToString();
            tb_b3pulse5.Text = get_pulse(h_pulse_3[4], l_pulse_3[4]).ToString();
            tb_b4pulse5.Text = get_pulse(h_pulse_4[4], l_pulse_4[4]).ToString();
            tb_b5pulse5.Text = get_pulse(h_pulse_5[4], l_pulse_5[4]).ToString();

            tb_b6pulse5.Text = get_pulse(h_pulse_6[4], l_pulse_6[4]).ToString();
            tb_b7pulse5.Text = get_pulse(h_pulse_7[4], l_pulse_7[4]).ToString();
            tb_b8pulse5.Text = get_pulse(h_pulse_8[4], l_pulse_8[4]).ToString();
            tb_b9pulse5.Text = get_pulse(h_pulse_9[4], l_pulse_9[4]).ToString();
            tb_b10pulse5.Text = get_pulse(h_pulse_10[4], l_pulse_10[4]).ToString();
            tb_b11pulse5.Text = get_pulse(h_pulse_11[4], l_pulse_11[4]).ToString();

            tb_b12pulse5.Text = get_pulse(h_pulse_12[4], l_pulse_12[4]).ToString();
            tb_b13pulse5.Text = get_pulse(h_pulse_13[4], l_pulse_13[4]).ToString();
            tb_b14pulse5.Text = get_pulse(h_pulse_14[4], l_pulse_14[4]).ToString();
            tb_b15pulse5.Text = get_pulse(h_pulse_15[4], l_pulse_15[4]).ToString();
            //time6
            tb_stpulse6.Text = get_pulse(h_pulse_st[5], l_pulse_st[5]).ToString();
            tb_b1pulse6.Text = get_pulse(h_pulse_1[5], l_pulse_1[5]).ToString();
            tb_b2pulse6.Text = get_pulse(h_pulse_2[5], l_pulse_2[5]).ToString();
            tb_b3pulse6.Text = get_pulse(h_pulse_3[5], l_pulse_3[5]).ToString();
            tb_b4pulse6.Text = get_pulse(h_pulse_4[5], l_pulse_4[5]).ToString();
            tb_b5pulse6.Text = get_pulse(h_pulse_5[5], l_pulse_5[5]).ToString();

            tb_b6pulse6.Text = get_pulse(h_pulse_6[5], l_pulse_6[5]).ToString();
            tb_b7pulse6.Text = get_pulse(h_pulse_7[5], l_pulse_7[5]).ToString();
            tb_b8pulse6.Text = get_pulse(h_pulse_8[5], l_pulse_8[5]).ToString();
            tb_b9pulse6.Text = get_pulse(h_pulse_9[5], l_pulse_9[5]).ToString();
            tb_b10pulse6.Text = get_pulse(h_pulse_10[5], l_pulse_10[5]).ToString();
            tb_b11pulse6.Text = get_pulse(h_pulse_11[5], l_pulse_11[5]).ToString();

            tb_b12pulse6.Text = get_pulse(h_pulse_12[5], l_pulse_12[5]).ToString();
            tb_b13pulse6.Text = get_pulse(h_pulse_13[5], l_pulse_13[5]).ToString();
            tb_b14pulse6.Text = get_pulse(h_pulse_14[5], l_pulse_14[5]).ToString();
            tb_b15pulse6.Text = get_pulse(h_pulse_15[5], l_pulse_15[5]).ToString();
            //time7
            tb_stpulse7.Text = get_pulse(h_pulse_st[6], l_pulse_st[6]).ToString();
            tb_b1pulse7.Text = get_pulse(h_pulse_1[6], l_pulse_1[6]).ToString();
            tb_b2pulse7.Text = get_pulse(h_pulse_2[6], l_pulse_2[6]).ToString();
            tb_b3pulse7.Text = get_pulse(h_pulse_3[6], l_pulse_3[6]).ToString();
            tb_b4pulse7.Text = get_pulse(h_pulse_4[6], l_pulse_4[6]).ToString();
            tb_b5pulse7.Text = get_pulse(h_pulse_5[6], l_pulse_5[6]).ToString();

            tb_b6pulse7.Text = get_pulse(h_pulse_6[6], l_pulse_6[6]).ToString();
            tb_b7pulse7.Text = get_pulse(h_pulse_7[6], l_pulse_7[6]).ToString();
            tb_b8pulse7.Text = get_pulse(h_pulse_8[6], l_pulse_8[6]).ToString();
            tb_b9pulse7.Text = get_pulse(h_pulse_9[6], l_pulse_9[6]).ToString();
            tb_b10pulse7.Text = get_pulse(h_pulse_10[6], l_pulse_10[6]).ToString();
            tb_b11pulse7.Text = get_pulse(h_pulse_11[6], l_pulse_11[6]).ToString();

            tb_b12pulse7.Text = get_pulse(h_pulse_12[6], l_pulse_12[6]).ToString();
            tb_b13pulse7.Text = get_pulse(h_pulse_13[6], l_pulse_13[6]).ToString();
            tb_b14pulse7.Text = get_pulse(h_pulse_14[6], l_pulse_14[6]).ToString();
            tb_b15pulse7.Text = get_pulse(h_pulse_15[6], l_pulse_15[6]).ToString();
            //time8
            tb_stpulse8.Text = get_pulse(h_pulse_st[7], l_pulse_st[7]).ToString();
            tb_b1pulse8.Text = get_pulse(h_pulse_1[7], l_pulse_1[7]).ToString();
            tb_b2pulse8.Text = get_pulse(h_pulse_2[7], l_pulse_2[7]).ToString();
            tb_b3pulse8.Text = get_pulse(h_pulse_3[7], l_pulse_3[7]).ToString();
            tb_b4pulse8.Text = get_pulse(h_pulse_4[7], l_pulse_4[7]).ToString();
            tb_b5pulse8.Text = get_pulse(h_pulse_5[7], l_pulse_5[7]).ToString();

            tb_b6pulse8.Text = get_pulse(h_pulse_6[7], l_pulse_6[7]).ToString();
            tb_b7pulse8.Text = get_pulse(h_pulse_7[7], l_pulse_7[7]).ToString();
            tb_b8pulse8.Text = get_pulse(h_pulse_8[7], l_pulse_8[7]).ToString();
            tb_b9pulse8.Text = get_pulse(h_pulse_9[7], l_pulse_9[7]).ToString();
            tb_b10pulse8.Text = get_pulse(h_pulse_10[7], l_pulse_10[7]).ToString();
            tb_b11pulse8.Text = get_pulse(h_pulse_11[7], l_pulse_11[7]).ToString();

            tb_b12pulse8.Text = get_pulse(h_pulse_12[7], l_pulse_12[7]).ToString();
            tb_b13pulse8.Text = get_pulse(h_pulse_13[7], l_pulse_13[7]).ToString();
            tb_b14pulse8.Text = get_pulse(h_pulse_14[7], l_pulse_14[7]).ToString();
            tb_b15pulse8.Text = get_pulse(h_pulse_15[7], l_pulse_15[7]).ToString();
            //time9
            tb_stpulse9.Text = get_pulse(h_pulse_st[8], l_pulse_st[8]).ToString();
            tb_b1pulse9.Text = get_pulse(h_pulse_1[8], l_pulse_1[8]).ToString();
            tb_b2pulse9.Text = get_pulse(h_pulse_2[8], l_pulse_2[8]).ToString();
            tb_b3pulse9.Text = get_pulse(h_pulse_3[8], l_pulse_3[8]).ToString();
            tb_b4pulse9.Text = get_pulse(h_pulse_4[8], l_pulse_4[8]).ToString();
            tb_b5pulse9.Text = get_pulse(h_pulse_5[8], l_pulse_5[8]).ToString();

            tb_b6pulse9.Text = get_pulse(h_pulse_6[8], l_pulse_6[8]).ToString();
            tb_b7pulse9.Text = get_pulse(h_pulse_7[8], l_pulse_7[8]).ToString();
            tb_b8pulse9.Text = get_pulse(h_pulse_8[8], l_pulse_8[8]).ToString();
            tb_b9pulse9.Text = get_pulse(h_pulse_9[8], l_pulse_9[8]).ToString();
            tb_b10pulse9.Text = get_pulse(h_pulse_10[8], l_pulse_10[8]).ToString();
            tb_b11pulse9.Text = get_pulse(h_pulse_11[8], l_pulse_11[8]).ToString();

            tb_b12pulse9.Text = get_pulse(h_pulse_12[8], l_pulse_12[8]).ToString();
            tb_b13pulse9.Text = get_pulse(h_pulse_13[8], l_pulse_13[8]).ToString();
            tb_b14pulse9.Text = get_pulse(h_pulse_14[8], l_pulse_14[8]).ToString();
            tb_b15pulse9.Text = get_pulse(h_pulse_15[8], l_pulse_15[8]).ToString();
        }
    }
}
   

