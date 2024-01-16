using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using feejoy_wojie.database;
using AxActUtlTypeLib;

namespace feejoy_wojie.subform
{
    public partial class    plan : DevExpress.XtraEditors.XtraForm
    {
        int returncode;
        public plan()
        {
            InitializeComponent();
        }

        private int PLC_connect()
        {
            try
            {
                FX3U.ActLogicalStationNumber = 10;
                FX3U.ActPassword = "";
                returncode = FX3U.Open();
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

        private int PLC_disconnect()
        {
            try
            {
                returncode = FX3U.Close();
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

        private string Read_M(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            int[] iData = new int[length];
            PLC_connect();
            iReturnCode = FX3U.ReadDeviceRandom(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        private string Read_D(string data_label)
        {
            int iReturnCode = -1;
            int length = System.Text.RegularExpressions.Regex.Split(data_label, "\n").Length;
            int iNumberOfData = length;
            short[] iData = new short[length];
            PLC_connect();
            iReturnCode = FX3U.ReadDeviceBlock2(data_label, iNumberOfData, out iData[0]);
            PLC_disconnect();
            if (iReturnCode == 0)
            {
                return iData[0].ToString();
            }
            else
            {
                return "-1";
            }
        }

        public string Write_M(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            int[] iData = new int[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int32.Parse(strArray[i]);
            }
            PLC_connect();
            iReturnCode = FX3U.WriteDeviceRandom(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string Write_D(string data_address, string value)
        {
            int iReturnCode = -1;
            string[] strArray = System.Text.RegularExpressions.Regex.Split(value, "\n");
            int iNumberOfData = strArray.Length;
            short[] iData = new short[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                iData[i] = Int16.Parse(strArray[i]);
            }
            PLC_connect();
            iReturnCode = FX3U.WriteDeviceBlock2(data_address, iNumberOfData, ref iData[0]);
            PLC_disconnect();
            if (iReturnCode == 0)
            {
                return String.Format("0x{0:x8} [HEX]", iReturnCode);
            }
            else
            {
                return "-1";
            }
        }

        private string tentime_D(string D_value)
        {
            string tentime_value = (Convert.ToSingle(D_value) * 10).ToString();
            return tentime_value;
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            if (tb_freq1.Text != "" && tb_freq2.Text != "" && tb_freq3.Text != "")
            {
                //Write_D("D241", tentime_D(tb_freq1.Text));
                plan_data.pl1 = tentime_D(tb_freq1.Text);
                //Write_D("D242", tentime_D(tb_freq2.Text));
                plan_data.pl2 = tentime_D(tb_freq2.Text);
                //Write_D("D243", tentime_D(tb_freq3.Text));
                plan_data.pl3 = tentime_D(tb_freq3.Text);
      
            }
            else
            {
                MessageBox.Show("流量控制参数——频率不能为空", "错误", MessageBoxButtons.OK);
            }
           
            if (tb_doorvalue1.Text != "" && tb_doorvalue2.Text != "" && tb_doorvalue3.Text != "")
            {
                //Write_D("D246", tentime_D(tb_doorvalue1.Text));
                plan_data.kd1 = tentime_D(tb_doorvalue1.Text);
                //Write_D("D247", tentime_D(tb_doorvalue2.Text));
                plan_data.kd2 = tentime_D(tb_doorvalue2.Text);
                //Write_D("D248", tentime_D(tb_doorvalue3.Text));
                plan_data.kd3 = tentime_D(tb_doorvalue3.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——开度不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_flowdelaytime.Text != "")
            {
                //Write_D("D228", tentime_D(tb_flowdelaytime.Text));
                plan_data.stable_time = Convert.ToInt16(tb_flowdelaytime.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——流量稳定延时不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_calpoints.Text != "")
            {
                Write_D("D259", tb_calpoints.Text);
            }
            else
            {
                MessageBox.Show("流量控制参数——标定点数", "错误", MessageBoxButtons.OK);
            }

            if (tb_manualtime.Text != "")
            {
                Write_D("D227", tentime_D(tb_manualtime.Text));
            }
            else
            {
                MessageBox.Show("手动标定控制参数——时间不能为空", "错误", MessageBoxButtons.OK);
            }

            if (tb_autotime1.Text != "" && tb_autotime2.Text != "" && tb_autotime3.Text != "" && tb_autotime4.Text != "" && tb_autotime5.Text != "")
            {
                //Write_D("D231", tentime_D(tb_autotime1.Text));
                //Write_D("D232", tentime_D(tb_autotime2.Text));
                //Write_D("D233", tentime_D(tb_autotime3.Text));
                //Write_D("D234", tentime_D(tb_autotime4.Text));
                //Write_D("D235", tentime_D(tb_autotime5.Text));
            }
            else
            {
                MessageBox.Show("自动标定控制参数——时间不能为空", "错误", MessageBoxButtons.OK);
            }

            Write_M("M32", "0");
            Write_M("M33", "0");
      
            MessageBox.Show("参数配置设置完成。","提示",MessageBoxButtons.OK);
        }

        private void plan_Load(object sender, EventArgs e)
        {
            tb_freq1.Text = "30";
            tb_freq2.Text = "37.3";
            tb_freq3.Text = "31";
            //tb_freq4.Text = "30 ";
            //tb_freq5.Text = "30";

            tb_doorvalue1.Text = "73";
            tb_doorvalue2.Text = "34";
            tb_doorvalue3.Text = "31";
            //tb_doorvalue4.Text = "100";
            //tb_doorvalue5.Text = "100";

            tb_autotime1.Text = "30";
            tb_autotime2.Text = "30";
            tb_autotime3.Text = "30";
            tb_autotime4.Text = "30";
            tb_autotime5.Text = "30";

            tb_flowdelaytime.Text = "180";
            tb_calpoints.Text = "3";

            tb_manualtime.Text = "30";
        }

    }
}