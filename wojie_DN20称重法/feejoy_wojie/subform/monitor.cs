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
        weight_comm weight_comm = new weight_comm();
        System.Timers.Timer t_weight = new System.Timers.Timer(300);
        Thread t_store = null;

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
        string[] weight = new string[9];


        public monitor()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private int PLC_connect_DN20_1()
        {
            try
            {
                DN20_1.ActLogicalStationNumber = 20;
                DN20_1.ActPassword = "";
                returncode = DN20_1.Open();
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


        private int PLC_disconnect_DN20_1()
        {
            try
            {
                returncode = DN20_1.Close();
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
            PLC_connect_DN20_1();
            iReturnCode = DN20_1.ReadDeviceRandom(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN20_1();
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
            PLC_connect_DN20_1();
            iReturnCode = DN20_1.ReadDeviceBlock2(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect_DN20_1();
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
            PLC_connect_DN20_1();
            iReturnCode = DN20_1.WriteDeviceRandom(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN20_1();
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
            PLC_connect_DN20_1();
            iReturnCode = DN20_1.WriteDeviceBlock2(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect_DN20_1();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private void monitor_Load(object sender, EventArgs e)
        {
            cal_point = 0;

            tb_settingfreq.Text = "10";
            tb_settingdoorvalue.Text = "100";

            tb_freq1.Text = "48";
            tb_freq2.Text = "29.7";
            tb_freq3.Text = "28.2";

            tb_doorvalue1.Text = "100";
            tb_doorvalue2.Text = "31";
            tb_doorvalue3.Text = "21.5";

            tb_flowdelaytime.Text = "180";
            tb_calpoints.Text = "3";
            tb_manualtime.Text = "30";

            tb_weightstabletime.Text = "30";
            tb_pourwatertime.Text = "180";

            weight_comm.sp_open();
            Write_M1("M121", "0");
            Write_M1("M120", "1");

            t_store = new Thread(new ThreadStart(weight_comm.read_weight));
            t_store.Start();
        }

        private void btn_calstop_Click(object sender, EventArgs e)
        {
            Write_M1("M12", "0");
            timer_weight.Enabled = false;
        }

        private void btn_saveorigindata_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = null;
            string filename = @"C:\DN20\DN20模板文件.xlsx";
            string firsttime = System.DateTime.Now.ToString().Replace(":", "-");
            string savetime = firsttime.Replace("/", "-");
            string storename = @"C:\DN20\DN20 " + savetime + ".xlsx";
            cal_file = storename;
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(fs);
            }

            ISheet sheet1 = workbook.GetSheet("被检表1");

            double[] b_time = new double[5] { Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10, Convert.ToDouble(Read_D1("D10")) / 10 };


            //标定时间
            for (int y = 12; y < 21; y = y + 3)
            {
                if (y < 15)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[0]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[0]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[0]);
                }

                if (y >= 15 && y < 18)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[1]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[1]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[1]);
                }

                if (y >= 18 && y < 21)
                {
                    sheet1.GetRow(y).GetCell(20).SetCellValue(b_time[2]);
                    sheet1.GetRow(y + 1).GetCell(20).SetCellValue(b_time[2]);
                    sheet1.GetRow(y + 2).GetCell(20).SetCellValue(b_time[2]);

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
                }

                if (y >= 15 && y < 18)
                {
                    sheet1.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse4.Text));
                    sheet1.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse5.Text));
                    sheet1.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse6.Text));
                }

                if (y >= 18 && y < 21)
                {
                    sheet1.GetRow(y).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse7.Text));
                    sheet1.GetRow(y + 1).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse8.Text));
                    sheet1.GetRow(y + 2).GetCell(21).SetCellValue(Convert.ToDouble(tb_stpulse9.Text));
                }
            }

            sheet1.ForceFormulaRecalculation = true;


            XSSFFormulaEvaluator dio = new XSSFFormulaEvaluator(workbook);


            using (FileStream fs_new = new FileStream(storename, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs_new);
                workbook.Close();
            }
            MessageBox.Show("数据表格已导出。", "提示", MessageBoxButtons.OK);
        }

        private string tentime_D(string D_value)
        {
            string tentime_value = (Convert.ToSingle(D_value) * 10).ToString();
            return tentime_value;
        }

        private void store_pulse(int calpoint)
        {
            h_pulse_st[calpoint] = Read_D1("D52"); l_pulse_st[calpoint] = Read_D1("D51");    
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
                        break;
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
            Write_M1("M121", "0");
            Write_M1("M120", "1");
            auto_cal = false;
            cal_point = 0;
            tb_D4.Clear();

            timer_weight.Enabled = false;

            tb_currentweight.Clear();
            tb_stpulse1.Clear();
            tb_stpulse2.Clear();
            tb_stpulse3.Clear();
            tb_stpulse4.Clear();
            tb_stpulse5.Clear();
            tb_stpulse6.Clear();
            tb_stpulse7.Clear();
            tb_stpulse8.Clear();
            tb_stpulse9.Clear();
            tb_stweight1.Clear();
            tb_stweight2.Clear();
            tb_stweight3.Clear();
            tb_stweight4.Clear();
            tb_stweight5.Clear();
            tb_stweight6.Clear();
            tb_stweight7.Clear();
            tb_stweight8.Clear();
            tb_stweight9.Clear();
            tb_tcweight1.Clear();
            tb_tcweight2.Clear();
            tb_tcweight3.Clear();
            tb_tcweight4.Clear();
            tb_tcweight5.Clear();
            tb_tcweight6.Clear();
            tb_tcweight7.Clear();
            tb_tcweight8.Clear();
            tb_tcweight9.Clear();
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
            Write_M1("M121", "1");

            cal_point = cal_point + 1;
            timer_weight.Enabled = true;
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

        private void btn_exportpulse_Click(object sender, EventArgs e)
        {
            //标准表脉冲
            tb_stpulse1.Text = get_pulse(h_pulse_st[0], l_pulse_st[0]).ToString();
            tb_stpulse2.Text = get_pulse(h_pulse_st[1], l_pulse_st[1]).ToString();
            tb_stpulse3.Text = get_pulse(h_pulse_st[2], l_pulse_st[2]).ToString();
            tb_stpulse4.Text = get_pulse(h_pulse_st[3], l_pulse_st[3]).ToString();
            tb_stpulse5.Text = get_pulse(h_pulse_st[4], l_pulse_st[4]).ToString();
            tb_stpulse6.Text = get_pulse(h_pulse_st[5], l_pulse_st[5]).ToString();
            tb_stpulse7.Text = get_pulse(h_pulse_st[6], l_pulse_st[6]).ToString();
            tb_stpulse8.Text = get_pulse(h_pulse_st[7], l_pulse_st[7]).ToString();
            tb_stpulse9.Text = get_pulse(h_pulse_st[8], l_pulse_st[8]).ToString();

            //脉冲转化质量

        }

        private void btn_dirtobox_Click(object sender, EventArgs e)
        {
            Write_M1("M121", "1");
        }

        private void btn_dirtocal_Click(object sender, EventArgs e)
        {
            Write_M1("M121", "0");
        }

        private void btn_pourwater_Click(object sender, EventArgs e)
        {
            Write_M1("M120", "0");
        }

        private void btn_savewater_Click(object sender, EventArgs e)
        {
            Write_M1("M120", "1");
        }

        private void btn_savesetting_Click(object sender, EventArgs e)
        {
            //time method
            Write_M1("M31", "0");
            Write_M1("M32", "0");
            Write_M1("M33", "0");

            if (tb_freq1.Text != "" && tb_freq2.Text != "" && tb_freq3.Text != "")
            {

                Write_D1("D241", tentime_D(tb_freq1.Text));
                plan_data.pl1 = tentime_D(tb_freq1.Text);
                Write_D1("D242", tentime_D(tb_freq2.Text));
                plan_data.pl2 = tentime_D(tb_freq2.Text);
                Write_D1("D243", tentime_D(tb_freq3.Text));
                plan_data.pl3 = tentime_D(tb_freq3.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——频率不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_doorvalue1.Text != "" && tb_doorvalue2.Text != "" && tb_doorvalue3.Text != "")
            {
                Write_D1("D246", tentime_D(tb_doorvalue1.Text));
                plan_data.kd1 = tentime_D(tb_doorvalue1.Text);
                Write_D1("D247", tentime_D(tb_doorvalue2.Text));
                plan_data.kd2 = tentime_D(tb_doorvalue2.Text);
                Write_D1("D248", tentime_D(tb_doorvalue3.Text));
                plan_data.kd3 = tentime_D(tb_doorvalue3.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——开度不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_flowdelaytime.Text != "")
            {
                plan_data.stable_time = Convert.ToInt16(tb_flowdelaytime.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——流量稳定延时不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_calpoints.Text != "")
            {
                Write_D1("D259", tb_calpoints.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——标定点数", "错误", MessageBoxButtons.OK);
            }

            if (tb_manualtime.Text != "")
            {
                Write_D1("D227", tentime_D(tb_manualtime.Text));
            }
            else
            {
                MessageBox.Show("流量控制参数——时间不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_weightstabletime.Text != ""&&int.Parse(tb_weightstabletime.Text)>=10)
            {
                plan_data.weightstabletime = int.Parse(tb_weightstabletime.Text);
            }
            else
            {
                MessageBox.Show("称重控制参数——质量稳定时间不能为空或小于30秒", "错误", MessageBoxButtons.OK);
            }
            if (tb_pourwatertime.Text != "" && int.Parse(tb_pourwatertime.Text) >= 30)
            {
                plan_data.pourwatertime= int.Parse(tb_pourwatertime.Text);
            }
            else
            {
                MessageBox.Show("称重控制参数——水箱排水时间不能为空或小于180秒", "错误", MessageBoxButtons.OK);
            }
            MessageBox.Show("参数配置设置完成。", "提示", MessageBoxButtons.OK);
        }

        private void btn_readweight_Click(object sender, EventArgs e)
        {
            t_weight.AutoReset = true;
            t_weight.Elapsed += new ElapsedEventHandler(read_weight);
            t_weight.Start();
        }

        private void read_weight(object o,ElapsedEventArgs e)
        {
            weight_comm.read_weight();
            tb_currentweight.Text = plan_data.weight;
        }

        private void btn_stopread_Click(object sender, EventArgs e)
        {
            t_weight.Stop();
            tb_currentweight.Clear();
        }

        private void timer_weight_Tick(object sender, EventArgs e)
        {
            if (Read_M1("M14") == "1")
            {
                Write_M1("M121", "0");
                switch (cal_point)
                {
                    case 1:
                        store_pulse(0);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight1.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");


                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 2;
                        Thread.Sleep(1000);
                        break;
                    case 2:
                        store_pulse(1);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight2.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 3;
                        Thread.Sleep(1000);
                        break;
                    case 3:
                        store_pulse(2);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight3.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Write_D1("D17", plan_data.pl2);
                        Write_D1("D16", plan_data.kd2);

                        Thread.Sleep(plan_data.stable_time * 1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 4;
                        Thread.Sleep(1000);
                        break;
                    case 4:
                        store_pulse(3);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight4.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 5;
                        Thread.Sleep(1000);
                        break;
                    case 5:
                        store_pulse(4);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight5.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 6;
                        Thread.Sleep(1000);
                        break;
                    case 6:
                        store_pulse(5);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight6.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Write_D1("D17", plan_data.pl3);
                        Write_D1("D16", plan_data.kd3);

                        Thread.Sleep(plan_data.stable_time * 1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 7;
                        Thread.Sleep(1000);
                        break;
                    case 7:
                        store_pulse(6);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight7.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 8;
                        Thread.Sleep(1000);
                        break;
                    case 8:
                        store_pulse(7);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight8.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);

                        Write_M1("M12", "1");
                        Write_M1("M121", "1");

                        cal_point = 9;
                        Thread.Sleep(1000);
                        break;
                    case 9:
                        store_pulse(8);

                        Thread.Sleep(plan_data.weightstabletime * 1000);

                        reset_store();
                        //tb_tcweight9.Text = plan_data.weight;

                        Write_M1("M120", "0");
                        Thread.Sleep(plan_data.pourwatertime * 1000);
                        Write_M1("M120", "1");

                        tb_D4.Text = cal_point.ToString();

                        Thread.Sleep(1000);
                        cal_point = 0;
                        stop_store();
                        break;
                }
            }
        }

        private void stop_store()
        {
            timer_weight.Enabled=false;
        }

        private void reset_store()
        {
            timer_weight.Enabled = false;
            switch (cal_point)
            {
                case 1:
                    tb_tcweight1.Text = plan_data.weight;
                    break;
                case 2:
                    tb_tcweight2.Text = plan_data.weight;
                    break;
                case 3:
                    tb_tcweight3.Text = plan_data.weight;
                    break;
                case 4:
                    tb_tcweight4.Text = plan_data.weight;
                    break;
                case 5:
                    tb_tcweight5.Text = plan_data.weight;
                    break;
                case 6:
                    tb_tcweight6.Text = plan_data.weight;
                    break;
                case 7:
                    tb_tcweight7.Text = plan_data.weight;
                    break;
                case 8:
                    tb_tcweight8.Text = plan_data.weight;
                    break;
                case 9:
                    tb_tcweight9.Text = plan_data.weight;
                    break;
            }
            timer_weight.Enabled = true;
        }

        private void btn_setzero_Click(object sender, EventArgs e)
        {
            weight_comm.set_zero();
        }

        private void btn_showrealvalue_Click(object sender, EventArgs e)
        {
           
        }

        private void tb_stpulse1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight1.Text = (Single.Parse(tb_stpulse1.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight2.Text = (Single.Parse(tb_stpulse2.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight3.Text = (Single.Parse(tb_stpulse3.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight4.Text = (Single.Parse(tb_stpulse4.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight5.Text = (Single.Parse(tb_stpulse5.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse6_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight6.Text = (Single.Parse(tb_stpulse6.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse7_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight7.Text = (Single.Parse(tb_stpulse7.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse8_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight8.Text = (Single.Parse(tb_stpulse8.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }

        private void tb_stpulse9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                tb_stweight9.Text = (Single.Parse(tb_stpulse9.Text) / 360000 * 1000).ToString();
            }
            catch
            {
            }
        }
    }
}
   

