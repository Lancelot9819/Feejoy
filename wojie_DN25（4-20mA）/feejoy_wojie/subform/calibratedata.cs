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
using System.IO.Ports;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using feejoy_wojie.database;

namespace feejoy_wojie.subform
{
    public partial class calibratedata : DevExpress.XtraEditors.XtraForm
    {

        public byte[] open_cal_old = new byte[25] { 0x0A, 0x10, 0x00, 0x4E, 0x00, 0x08, 0x10, 0x3F, 0x80, 0x00, 0x00, 0x49, 0x6C, 0xC6, 0x30, 0x00, 0x00, 0x00, 0x00, 0x40, 0x40, 0x00, 0x00, 0xE5, 0xAA };

        public byte[] open_cal = new byte[17] { 0x0A, 0x10, 0x00, 0x4E, 0x00, 0x04, 0x08, 0x3F, 0x80, 0x00, 0x00, 0x49, 0x6C, 0xC6, 0x30, 0x52, 0xFC };
        public byte[] close_cal = new byte[17] { 0x0A, 0x10, 0x00, 0x4E, 0x00, 0x04, 0x08, 0x40, 0x40, 0x00, 0x00, 0x49, 0x6C, 0xC6, 0x30, 0xD5, 0x94 };
        public byte[] CAL_INIT = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x42, 0xAA, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xBC, 0x5A };

        public byte[] b1_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b2_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b3_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b4_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b5_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b6_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b7_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b8_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b9_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b10_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b11_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b12_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b13_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b14_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] b15_cal = new byte[37] { 0x0A, 0x10, 0x00, 0x09, 0x00, 0x0E, 0x1C, 0x43, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

        public byte[] changeDN25 = new byte[65] { 0x0A, 0x10, 0x00, 0x78, 0x00, 0x1C, 0x38, 0x43, 0x7F, 0x00, 0x00, 0x3F, 0x80, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x41, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3C, 0xEF };
        public byte[] sum_init = new byte[21] { 0x0A, 0x10, 0x00, 0x4E, 0x00, 0x06, 0x0C, 0x43, 0x7F, 0x00, 0x00, 0x49, 0x6C, 0xC6, 0x30, 0x00, 0x00, 0x00, 0x00, 0xBF, 0xB6 };

        public SerialPort sp1 = new SerialPort("COM4", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp2 = new SerialPort("COM5", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp3 = new SerialPort("COM11", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp4 = new SerialPort("COM13", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp5 = new SerialPort("COM10", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp6 = new SerialPort("COM12", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp7 = new SerialPort("COM7", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp8 = new SerialPort("COM8", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp9 = new SerialPort("COM14", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp10 = new SerialPort("COM15", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp11 = new SerialPort("COM16", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp12 = new SerialPort("COM19", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp13 = new SerialPort("COM17", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp14 = new SerialPort("COM20", 9600, Parity.None, 8, StopBits.One);
        public SerialPort sp15 = new SerialPort("COM18", 9600, Parity.None, 8, StopBits.One);
        public calibratedata()
        {
            InitializeComponent();
        }

        private void data1_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b1_cal[11] = b_b_flow1[3];
            b1_cal[12] = b_b_flow1[2];
            b1_cal[13] = b_b_flow1[1];
            b1_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b1_cal[15] = b_b_k1[3];
            b1_cal[16] = b_b_k1[2];
            b1_cal[17] = b_b_k1[1];
            b1_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b1_cal[19] = b_b_flow2[3];
            b1_cal[20] = b_b_flow2[2];
            b1_cal[21] = b_b_flow2[1];
            b1_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b1_cal[23] = b_b_k2[3];
            b1_cal[24] = b_b_k2[2];
            b1_cal[25] = b_b_k2[1];
            b1_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b1_cal[27] = b_b_flow3[3];
            b1_cal[28] = b_b_flow3[2];
            b1_cal[29] = b_b_flow3[1];
            b1_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b1_cal[31] = b_b_k3[3];
            b1_cal[32] = b_b_k3[2];
            b1_cal[33] = b_b_k3[1];
            b1_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b1_cal);
            b1_cal[35] = crc[1];
            b1_cal[36] = crc[0];
        }

        private void data2_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b2_cal[11] = b_b_flow1[3];
            b2_cal[12] = b_b_flow1[2];
            b2_cal[13] = b_b_flow1[1];
            b2_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b2_cal[15] = b_b_k1[3];
            b2_cal[16] = b_b_k1[2];
            b2_cal[17] = b_b_k1[1];
            b2_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b2_cal[19] = b_b_flow2[3];
            b2_cal[20] = b_b_flow2[2];
            b2_cal[21] = b_b_flow2[1];
            b2_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b2_cal[23] = b_b_k2[3];
            b2_cal[24] = b_b_k2[2];
            b2_cal[25] = b_b_k2[1];
            b2_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b2_cal[27] = b_b_flow3[3];
            b2_cal[28] = b_b_flow3[2];
            b2_cal[29] = b_b_flow3[1];
            b2_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b2_cal[31] = b_b_k3[3];
            b2_cal[32] = b_b_k3[2];
            b2_cal[33] = b_b_k3[1];
            b2_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b2_cal);
            b2_cal[35] = crc[1];
            b2_cal[36] = crc[0];
        }

        private void data3_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b3_cal[11] = b_b_flow1[3];
            b3_cal[12] = b_b_flow1[2];
            b3_cal[13] = b_b_flow1[1];
            b3_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b3_cal[15] = b_b_k1[3];
            b3_cal[16] = b_b_k1[2];
            b3_cal[17] = b_b_k1[1];
            b3_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b3_cal[19] = b_b_flow2[3];
            b3_cal[20] = b_b_flow2[2];
            b3_cal[21] = b_b_flow2[1];
            b3_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b3_cal[23] = b_b_k2[3];
            b3_cal[24] = b_b_k2[2];
            b3_cal[25] = b_b_k2[1];
            b3_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b3_cal[27] = b_b_flow3[3];
            b3_cal[28] = b_b_flow3[2];
            b3_cal[29] = b_b_flow3[1];
            b3_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b3_cal[31] = b_b_k3[3];
            b3_cal[32] = b_b_k3[2];
            b3_cal[33] = b_b_k3[1];
            b3_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b3_cal);
            b3_cal[35] = crc[1];
            b3_cal[36] = crc[0];
        }

        private void data4_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b4_cal[11] = b_b_flow1[3];
            b4_cal[12] = b_b_flow1[2];
            b4_cal[13] = b_b_flow1[1];
            b4_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b4_cal[15] = b_b_k1[3];
            b4_cal[16] = b_b_k1[2];
            b4_cal[17] = b_b_k1[1];
            b4_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b4_cal[19] = b_b_flow2[3];
            b4_cal[20] = b_b_flow2[2];
            b4_cal[21] = b_b_flow2[1];
            b4_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b4_cal[23] = b_b_k2[3];
            b4_cal[24] = b_b_k2[2];
            b4_cal[25] = b_b_k2[1];
            b4_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b4_cal[27] = b_b_flow3[3];
            b4_cal[28] = b_b_flow3[2];
            b4_cal[29] = b_b_flow3[1];
            b4_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b4_cal[31] = b_b_k3[3];
            b4_cal[32] = b_b_k3[2];
            b4_cal[33] = b_b_k3[1];
            b4_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b4_cal);
            b4_cal[35] = crc[1];
            b4_cal[36] = crc[0];
        }

        private void data5_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b5_cal[11] = b_b_flow1[3];
            b5_cal[12] = b_b_flow1[2];
            b5_cal[13] = b_b_flow1[1];
            b5_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b5_cal[15] = b_b_k1[3];
            b5_cal[16] = b_b_k1[2];
            b5_cal[17] = b_b_k1[1];
            b5_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b5_cal[19] = b_b_flow2[3];
            b5_cal[20] = b_b_flow2[2];
            b5_cal[21] = b_b_flow2[1];
            b5_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b5_cal[23] = b_b_k2[3];
            b5_cal[24] = b_b_k2[2];
            b5_cal[25] = b_b_k2[1];
            b5_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b5_cal[27] = b_b_flow3[3];
            b5_cal[28] = b_b_flow3[2];
            b5_cal[29] = b_b_flow3[1];
            b5_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b5_cal[31] = b_b_k3[3];
            b5_cal[32] = b_b_k3[2];
            b5_cal[33] = b_b_k3[1];
            b5_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b5_cal);
            b5_cal[35] = crc[1];
            b5_cal[36] = crc[0];
        }
        private void data6_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b6_cal[11] = b_b_flow1[3];
            b6_cal[12] = b_b_flow1[2];
            b6_cal[13] = b_b_flow1[1];
            b6_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b6_cal[15] = b_b_k1[3];
            b6_cal[16] = b_b_k1[2];
            b6_cal[17] = b_b_k1[1];
            b6_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b6_cal[19] = b_b_flow2[3];
            b6_cal[20] = b_b_flow2[2];
            b6_cal[21] = b_b_flow2[1];
            b6_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b6_cal[23] = b_b_k2[3];
            b6_cal[24] = b_b_k2[2];
            b6_cal[25] = b_b_k2[1];
            b6_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b6_cal[27] = b_b_flow3[3];
            b6_cal[28] = b_b_flow3[2];
            b6_cal[29] = b_b_flow3[1];
            b6_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b6_cal[31] = b_b_k3[3];
            b6_cal[32] = b_b_k3[2];
            b6_cal[33] = b_b_k3[1];
            b6_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b6_cal);
            b6_cal[35] = crc[1];
            b6_cal[36] = crc[0];
        }

        private void data7_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b7_cal[11] = b_b_flow1[3];
            b7_cal[12] = b_b_flow1[2];
            b7_cal[13] = b_b_flow1[1];
            b7_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b7_cal[15] = b_b_k1[3];
            b7_cal[16] = b_b_k1[2];
            b7_cal[17] = b_b_k1[1];
            b7_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b7_cal[19] = b_b_flow2[3];
            b7_cal[20] = b_b_flow2[2];
            b7_cal[21] = b_b_flow2[1];
            b7_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b7_cal[23] = b_b_k2[3];
            b7_cal[24] = b_b_k2[2];
            b7_cal[25] = b_b_k2[1];
            b7_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b7_cal[27] = b_b_flow3[3];
            b7_cal[28] = b_b_flow3[2];
            b7_cal[29] = b_b_flow3[1];
            b7_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b7_cal[31] = b_b_k3[3];
            b7_cal[32] = b_b_k3[2];
            b7_cal[33] = b_b_k3[1];
            b7_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b7_cal);
            b7_cal[35] = crc[1];
            b7_cal[36] = crc[0];
        }

        private void data8_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b8_cal[11] = b_b_flow1[3];
            b8_cal[12] = b_b_flow1[2];
            b8_cal[13] = b_b_flow1[1];
            b8_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b8_cal[15] = b_b_k1[3];
            b8_cal[16] = b_b_k1[2];
            b8_cal[17] = b_b_k1[1];
            b8_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b8_cal[19] = b_b_flow2[3];
            b8_cal[20] = b_b_flow2[2];
            b8_cal[21] = b_b_flow2[1];
            b8_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b8_cal[23] = b_b_k2[3];
            b8_cal[24] = b_b_k2[2];
            b8_cal[25] = b_b_k2[1];
            b8_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b8_cal[27] = b_b_flow3[3];
            b8_cal[28] = b_b_flow3[2];
            b8_cal[29] = b_b_flow3[1];
            b8_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b8_cal[31] = b_b_k3[3];
            b8_cal[32] = b_b_k3[2];
            b8_cal[33] = b_b_k3[1];
            b8_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b8_cal);
            b8_cal[35] = crc[1];
            b8_cal[36] = crc[0];
        }

        private void data9_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b9_cal[11] = b_b_flow1[3];
            b9_cal[12] = b_b_flow1[2];
            b9_cal[13] = b_b_flow1[1];
            b9_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b9_cal[15] = b_b_k1[3];
            b9_cal[16] = b_b_k1[2];
            b9_cal[17] = b_b_k1[1];
            b9_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b9_cal[19] = b_b_flow2[3];
            b9_cal[20] = b_b_flow2[2];
            b9_cal[21] = b_b_flow2[1];
            b9_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b9_cal[23] = b_b_k2[3];
            b9_cal[24] = b_b_k2[2];
            b9_cal[25] = b_b_k2[1];
            b9_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b9_cal[27] = b_b_flow3[3];
            b9_cal[28] = b_b_flow3[2];
            b9_cal[29] = b_b_flow3[1];
            b9_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b9_cal[31] = b_b_k3[3];
            b9_cal[32] = b_b_k3[2];
            b9_cal[33] = b_b_k3[1];
            b9_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b9_cal);
            b9_cal[35] = crc[1];
            b9_cal[36] = crc[0];
        }

        private void data10_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b10_cal[11] = b_b_flow1[3];
            b10_cal[12] = b_b_flow1[2];
            b10_cal[13] = b_b_flow1[1];
            b10_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b10_cal[15] = b_b_k1[3];
            b10_cal[16] = b_b_k1[2];
            b10_cal[17] = b_b_k1[1];
            b10_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b10_cal[19] = b_b_flow2[3];
            b10_cal[20] = b_b_flow2[2];
            b10_cal[21] = b_b_flow2[1];
            b10_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b10_cal[23] = b_b_k2[3];
            b10_cal[24] = b_b_k2[2];
            b10_cal[25] = b_b_k2[1];
            b10_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b10_cal[27] = b_b_flow3[3];
            b10_cal[28] = b_b_flow3[2];
            b10_cal[29] = b_b_flow3[1];
            b10_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b10_cal[31] = b_b_k3[3];
            b10_cal[32] = b_b_k3[2];
            b10_cal[33] = b_b_k3[1];
            b10_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b10_cal);
            b10_cal[35] = crc[1];
            b10_cal[36] = crc[0];
        }

        private void data11_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b11_cal[11] = b_b_flow1[3];
            b11_cal[12] = b_b_flow1[2];
            b11_cal[13] = b_b_flow1[1];
            b11_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b11_cal[15] = b_b_k1[3];
            b11_cal[16] = b_b_k1[2];
            b11_cal[17] = b_b_k1[1];
            b11_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b11_cal[19] = b_b_flow2[3];
            b11_cal[20] = b_b_flow2[2];
            b11_cal[21] = b_b_flow2[1];
            b11_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b11_cal[23] = b_b_k2[3];
            b11_cal[24] = b_b_k2[2];
            b11_cal[25] = b_b_k2[1];
            b11_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b11_cal[27] = b_b_flow3[3];
            b11_cal[28] = b_b_flow3[2];
            b11_cal[29] = b_b_flow3[1];
            b11_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b11_cal[31] = b_b_k3[3];
            b11_cal[32] = b_b_k3[2];
            b11_cal[33] = b_b_k3[1];
            b11_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b11_cal);
            b11_cal[35] = crc[1];
            b11_cal[36] = crc[0];
        }

        private void data12_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b12_cal[11] = b_b_flow1[3];
            b12_cal[12] = b_b_flow1[2];
            b12_cal[13] = b_b_flow1[1];
            b12_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b12_cal[15] = b_b_k1[3];
            b12_cal[16] = b_b_k1[2];
            b12_cal[17] = b_b_k1[1];
            b12_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b12_cal[19] = b_b_flow2[3];
            b12_cal[20] = b_b_flow2[2];
            b12_cal[21] = b_b_flow2[1];
            b12_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b12_cal[23] = b_b_k2[3];
            b12_cal[24] = b_b_k2[2];
            b12_cal[25] = b_b_k2[1];
            b12_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b12_cal[27] = b_b_flow3[3];
            b12_cal[28] = b_b_flow3[2];
            b12_cal[29] = b_b_flow3[1];
            b12_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b12_cal[31] = b_b_k3[3];
            b12_cal[32] = b_b_k3[2];
            b12_cal[33] = b_b_k3[1];
            b12_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b12_cal);
            b12_cal[35] = crc[1];
            b12_cal[36] = crc[0];
        }

        private void data13_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b13_cal[11] = b_b_flow1[3];
            b13_cal[12] = b_b_flow1[2];
            b13_cal[13] = b_b_flow1[1];
            b13_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b13_cal[15] = b_b_k1[3];
            b13_cal[16] = b_b_k1[2];
            b13_cal[17] = b_b_k1[1];
            b13_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b13_cal[19] = b_b_flow2[3];
            b13_cal[20] = b_b_flow2[2];
            b13_cal[21] = b_b_flow2[1];
            b13_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b13_cal[23] = b_b_k2[3];
            b13_cal[24] = b_b_k2[2];
            b13_cal[25] = b_b_k2[1];
            b13_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b13_cal[27] = b_b_flow3[3];
            b13_cal[28] = b_b_flow3[2];
            b13_cal[29] = b_b_flow3[1];
            b13_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b13_cal[31] = b_b_k3[3];
            b13_cal[32] = b_b_k3[2];
            b13_cal[33] = b_b_k3[1];
            b13_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b13_cal);
            b13_cal[35] = crc[1];
            b13_cal[36] = crc[0];
        }

        private void data14_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b14_cal[11] = b_b_flow1[3];
            b14_cal[12] = b_b_flow1[2];
            b14_cal[13] = b_b_flow1[1];
            b14_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b14_cal[15] = b_b_k1[3];
            b14_cal[16] = b_b_k1[2];
            b14_cal[17] = b_b_k1[1];
            b14_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b14_cal[19] = b_b_flow2[3];
            b14_cal[20] = b_b_flow2[2];
            b14_cal[21] = b_b_flow2[1];
            b14_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b14_cal[23] = b_b_k2[3];
            b14_cal[24] = b_b_k2[2];
            b14_cal[25] = b_b_k2[1];
            b14_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b14_cal[27] = b_b_flow3[3];
            b14_cal[28] = b_b_flow3[2];
            b14_cal[29] = b_b_flow3[1];
            b14_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b14_cal[31] = b_b_k3[3];
            b14_cal[32] = b_b_k3[2];
            b14_cal[33] = b_b_k3[1];
            b14_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b14_cal);
            b14_cal[35] = crc[1];
            b14_cal[36] = crc[0];
        }

        private void data15_floattobyte(double b_flow1, double b_k1, double b_flow2, double b_k2, double b_flow3, double b_k3)
        {
            byte[] b_b_flow1 = BitConverter.GetBytes(Convert.ToSingle(b_flow1));
            b15_cal[11] = b_b_flow1[3];
            b15_cal[12] = b_b_flow1[2];
            b15_cal[13] = b_b_flow1[1];
            b15_cal[14] = b_b_flow1[0];
            byte[] b_b_k1 = BitConverter.GetBytes(Convert.ToSingle(b_k1));
            b15_cal[15] = b_b_k1[3];
            b15_cal[16] = b_b_k1[2];
            b15_cal[17] = b_b_k1[1];
            b15_cal[18] = b_b_k1[0];

            byte[] b_b_flow2 = BitConverter.GetBytes(Convert.ToSingle(b_flow2));
            b15_cal[19] = b_b_flow2[3];
            b15_cal[20] = b_b_flow2[2];
            b15_cal[21] = b_b_flow2[1];
            b15_cal[22] = b_b_flow2[0];
            byte[] b_b_k2 = BitConverter.GetBytes(Convert.ToSingle(b_k2));
            b15_cal[23] = b_b_k2[3];
            b15_cal[24] = b_b_k2[2];
            b15_cal[25] = b_b_k2[1];
            b15_cal[26] = b_b_k2[0];

            byte[] b_b_flow3 = BitConverter.GetBytes(Convert.ToSingle(b_flow3));
            b15_cal[27] = b_b_flow3[3];
            b15_cal[28] = b_b_flow3[2];
            b15_cal[29] = b_b_flow3[1];
            b15_cal[30] = b_b_flow3[0];
            byte[] b_b_k3 = BitConverter.GetBytes(Convert.ToSingle(b_k3));
            b15_cal[31] = b_b_k3[3];
            b15_cal[32] = b_b_k3[2];
            b15_cal[33] = b_b_k3[1];
            b15_cal[34] = b_b_k3[0];

            byte[] crc = CRC(b15_cal);
            b15_cal[35] = crc[1];
            b15_cal[36] = crc[0];
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

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            tb_info.Clear();
        }

        private void btn_refreshdata_Click(object sender, EventArgs e)
        {
            //b1
            tb_b1flow1.Text = plan_data.b1_flow1.ToString();
            tb_b1k1.Text = plan_data.b1_k1.ToString();
            tb_b1flow2.Text = plan_data.b1_flow2.ToString();
            tb_b1k2.Text = plan_data.b1_k2.ToString();
            tb_b1flow3.Text = plan_data.b1_flow3.ToString();
            tb_b1k3.Text = plan_data.b1_k3.ToString();

            data1_floattobyte(plan_data.b1_flow1, plan_data.b1_k1, plan_data.b1_flow2, plan_data.b1_k2, plan_data.b1_flow3, plan_data.b1_k3);
            //b2
            tb_b2flow1.Text = plan_data.b2_flow1.ToString();
            tb_b2k1.Text = plan_data.b2_k1.ToString();
            tb_b2flow2.Text = plan_data.b2_flow2.ToString();
            tb_b2k2.Text = plan_data.b2_k2.ToString();
            tb_b2flow3.Text = plan_data.b2_flow3.ToString();
            tb_b2k3.Text = plan_data.b2_k3.ToString();

            data2_floattobyte(plan_data.b2_flow1, plan_data.b2_k1, plan_data.b2_flow2, plan_data.b2_k2, plan_data.b2_flow3, plan_data.b2_k3);
            //b3
            tb_b3flow1.Text = plan_data.b3_flow1.ToString();
            tb_b3k1.Text = plan_data.b3_k1.ToString();
            tb_b3flow2.Text = plan_data.b3_flow2.ToString();
            tb_b3k2.Text = plan_data.b3_k2.ToString();
            tb_b3flow3.Text = plan_data.b3_flow3.ToString();
            tb_b3k3.Text = plan_data.b3_k3.ToString();

            data3_floattobyte(plan_data.b3_flow1, plan_data.b3_k1, plan_data.b3_flow2, plan_data.b3_k2, plan_data.b3_flow3, plan_data.b3_k3);
            //b4
            tb_b4flow1.Text = plan_data.b4_flow1.ToString();
            tb_b4k1.Text = plan_data.b4_k1.ToString();
            tb_b4flow2.Text = plan_data.b4_flow2.ToString();
            tb_b4k2.Text = plan_data.b4_k2.ToString();
            tb_b4flow3.Text = plan_data.b4_flow3.ToString();
            tb_b4k3.Text = plan_data.b4_k3.ToString();

            data4_floattobyte(plan_data.b4_flow1, plan_data.b4_k1, plan_data.b4_flow2, plan_data.b4_k2, plan_data.b4_flow3, plan_data.b4_k3);
            //b5
            tb_b5flow1.Text = plan_data.b5_flow1.ToString();
            tb_b5k1.Text = plan_data.b5_k1.ToString();
            tb_b5flow2.Text = plan_data.b5_flow2.ToString();
            tb_b5k2.Text = plan_data.b5_k2.ToString();
            tb_b5flow3.Text = plan_data.b5_flow3.ToString();
            tb_b5k3.Text = plan_data.b5_k3.ToString();

            data5_floattobyte(plan_data.b5_flow1, plan_data.b5_k1, plan_data.b5_flow2, plan_data.b5_k2, plan_data.b5_flow3, plan_data.b5_k3);
            //b6
            tb_b6flow1.Text = plan_data.b6_flow1.ToString();
            tb_b6k1.Text = plan_data.b6_k1.ToString();
            tb_b6flow2.Text = plan_data.b6_flow2.ToString();
            tb_b6k2.Text = plan_data.b6_k2.ToString();
            tb_b6flow3.Text = plan_data.b6_flow3.ToString();
            tb_b6k3.Text = plan_data.b6_k3.ToString();

            data6_floattobyte(plan_data.b6_flow1, plan_data.b6_k1, plan_data.b6_flow2, plan_data.b6_k2, plan_data.b6_flow3, plan_data.b6_k3);
            //b7
            tb_b7flow1.Text = plan_data.b7_flow1.ToString();
            tb_b7k1.Text = plan_data.b7_k1.ToString();
            tb_b7flow2.Text = plan_data.b7_flow2.ToString();
            tb_b7k2.Text = plan_data.b7_k2.ToString();
            tb_b7flow3.Text = plan_data.b7_flow3.ToString();
            tb_b7k3.Text = plan_data.b7_k3.ToString();

            data7_floattobyte(plan_data.b7_flow1, plan_data.b7_k1, plan_data.b7_flow2, plan_data.b7_k2, plan_data.b7_flow3, plan_data.b7_k3);
            //b8
            tb_b8flow1.Text = plan_data.b8_flow1.ToString();
            tb_b8k1.Text = plan_data.b8_k1.ToString();
            tb_b8flow2.Text = plan_data.b8_flow2.ToString();
            tb_b8k2.Text = plan_data.b8_k2.ToString();
            tb_b8flow3.Text = plan_data.b8_flow3.ToString();
            tb_b8k3.Text = plan_data.b8_k3.ToString();

            data8_floattobyte(plan_data.b8_flow1, plan_data.b8_k1, plan_data.b8_flow2, plan_data.b8_k2, plan_data.b8_flow3, plan_data.b8_k3);
            //b9
            tb_b9flow1.Text = plan_data.b9_flow1.ToString();
            tb_b9k1.Text = plan_data.b9_k1.ToString();
            tb_b9flow2.Text = plan_data.b9_flow2.ToString();
            tb_b9k2.Text = plan_data.b9_k2.ToString();
            tb_b9flow3.Text = plan_data.b9_flow3.ToString();
            tb_b9k3.Text = plan_data.b9_k3.ToString();

            data9_floattobyte(plan_data.b9_flow1, plan_data.b9_k1, plan_data.b9_flow2, plan_data.b9_k2, plan_data.b9_flow3, plan_data.b9_k3);
            //b10
            tb_b10flow1.Text = plan_data.b10_flow1.ToString();
            tb_b10k1.Text = plan_data.b10_k1.ToString();
            tb_b10flow2.Text = plan_data.b10_flow2.ToString();
            tb_b10k2.Text = plan_data.b10_k2.ToString();
            tb_b10flow3.Text = plan_data.b10_flow3.ToString();
            tb_b10k3.Text = plan_data.b10_k3.ToString();

            data10_floattobyte(plan_data.b10_flow1, plan_data.b10_k1, plan_data.b10_flow2, plan_data.b10_k2, plan_data.b10_flow3, plan_data.b10_k3);
            //b11
            tb_b11flow1.Text = plan_data.b11_flow1.ToString();
            tb_b11k1.Text = plan_data.b11_k1.ToString();
            tb_b11flow2.Text = plan_data.b11_flow2.ToString();
            tb_b11k2.Text = plan_data.b11_k2.ToString();
            tb_b11flow3.Text = plan_data.b11_flow3.ToString();
            tb_b11k3.Text = plan_data.b11_k3.ToString();

            data11_floattobyte(plan_data.b11_flow1, plan_data.b11_k1, plan_data.b11_flow2, plan_data.b11_k2, plan_data.b11_flow3, plan_data.b11_k3);
            //b12
            tb_b12flow1.Text = plan_data.b12_flow1.ToString();
            tb_b12k1.Text = plan_data.b12_k1.ToString();
            tb_b12flow2.Text = plan_data.b12_flow2.ToString();
            tb_b12k2.Text = plan_data.b12_k2.ToString();
            tb_b12flow3.Text = plan_data.b12_flow3.ToString();
            tb_b12k3.Text = plan_data.b12_k3.ToString();

            data12_floattobyte(plan_data.b12_flow1, plan_data.b12_k1, plan_data.b12_flow2, plan_data.b12_k2, plan_data.b12_flow3, plan_data.b12_k3);
            //b13
            tb_b13flow1.Text = plan_data.b13_flow1.ToString();
            tb_b13k1.Text = plan_data.b13_k1.ToString();
            tb_b13flow2.Text = plan_data.b13_flow2.ToString();
            tb_b13k2.Text = plan_data.b13_k2.ToString();
            tb_b13flow3.Text = plan_data.b13_flow3.ToString();
            tb_b13k3.Text = plan_data.b13_k3.ToString();

            data13_floattobyte(plan_data.b13_flow1, plan_data.b13_k1, plan_data.b13_flow2, plan_data.b13_k2, plan_data.b13_flow3, plan_data.b13_k3);
            //b14
            tb_b14flow1.Text = plan_data.b14_flow1.ToString();
            tb_b14k1.Text = plan_data.b14_k1.ToString();
            tb_b14flow2.Text = plan_data.b14_flow2.ToString();
            tb_b14k2.Text = plan_data.b14_k2.ToString();
            tb_b14flow3.Text = plan_data.b14_flow3.ToString();
            tb_b14k3.Text = plan_data.b14_k3.ToString();

            data14_floattobyte(plan_data.b14_flow1, plan_data.b14_k1, plan_data.b14_flow2, plan_data.b14_k2, plan_data.b14_flow3, plan_data.b14_k3);
            //b15
            tb_b15flow1.Text = plan_data.b15_flow1.ToString();
            tb_b15k1.Text = plan_data.b15_k1.ToString();
            tb_b15flow2.Text = plan_data.b15_flow2.ToString();
            tb_b15k2.Text = plan_data.b15_k2.ToString();
            tb_b15flow3.Text = plan_data.b15_flow3.ToString();
            tb_b15k3.Text = plan_data.b15_k3.ToString();

            data15_floattobyte(plan_data.b15_flow1, plan_data.b15_k1, plan_data.b15_flow2, plan_data.b15_k2, plan_data.b15_flow3, plan_data.b15_k3);
        }

        private void write1()
        {
            sp1.Open();
            if (sp1.IsOpen)
            {
                try
                {
                    sp1.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp1.Write(b1_cal, 0, b1_cal.Length);

                    Thread.Sleep(3000);

                    sp8.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp8.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Close();
                }
                catch { }
            }
        }

        private void write2()
        {
            sp2.Open();
            if (sp2.IsOpen)
            {
                try
                {
                    sp2.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp2.Write(b2_cal, 0, b2_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp2.Close();
                }
                catch { }
            }
        }

        private void write3()
        {
            sp3.Open();
            if (sp3.IsOpen)
            {
                try
                {
                    sp3.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp3.Write(b3_cal, 0, b3_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp3.Close();
                }
                catch { }
            }
        }

        private void write4()
        {
            sp4.Open();
            if (sp4.IsOpen)
            {
                try
                {
                    sp4.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp4.Write(b4_cal, 0, b4_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp4.Close();
                }
                catch { }
            }
        }

        private void write5()
        {
            sp5.Open();
            if (sp5.IsOpen)
            {
                try
                {
                    sp5.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp5.Write(b5_cal, 0, b5_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp5.Close();
                }
                catch { }
            }
        }

        private void write6()
        {
            sp6.Open();
            if (sp6.IsOpen)
            {
                try
                {
                    sp6.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp6.Write(b6_cal, 0, b6_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp6.Close();
                }
                catch { }
            }
        }

        private void write7()
        {
            sp7.Open();
            if (sp7.IsOpen)
            {
                try
                {
                    sp7.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp7.Write(b7_cal, 0, b7_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp7.Close();
                }
                catch { }
            }
        }

        private void write8()
        {
            sp8.Open();
            if (sp8.IsOpen)
            {
                try
                {
                    sp8.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp8.Write(b8_cal, 0, b8_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp8.Close();
                }
                catch { }
            }
        }

        private void write9()
        {
            sp9.Open();
            if (sp9.IsOpen)
            {
                try
                {
                    sp9.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp9.Write(b9_cal, 0, b9_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp9.Close();
                }
                catch { }
            }
        }

        private void write10()
        {
            sp10.Open();
            if (sp10.IsOpen)
            {
                try
                {
                    sp10.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp10.Write(b10_cal, 0, b10_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp10.Close();
                }
                catch { }
            }
        }

        private void write11()
        {
            sp11.Open();
            if (sp11.IsOpen)
            {
                try
                {
                    sp11.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp11.Write(b11_cal, 0, b11_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp11.Close();
                }
                catch { }
            }
        }

        private void write12()
        {

            sp12.Open();
            if (sp12.IsOpen)
            {
                try
                {
                    sp12.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp12.Write(b12_cal, 0, b12_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp12.Close();
                }
                catch { }
            }
        }

        private void write13()
        {
            sp13.Open();
            if (sp13.IsOpen)
            {
                try
                {
                    sp13.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp13.Write(b13_cal, 0, b13_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp13.Close();
                }
                catch { }
            }
        }

        private void write14()
        {
            sp14.Open();
            if (sp14.IsOpen)
            {
                try
                {
                    sp14.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp14.Write(b14_cal, 0, b14_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp14.Close();
                }
                catch { }
            }
        }

        private void write15()
        {
            sp15.Open();
            if (sp15.IsOpen)
            {
                try
                {
                    sp15.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp15.Write(b15_cal, 0, b15_cal.Length);

                    Thread.Sleep(3000);

                    sp1.Write(sum_init, 0, sum_init.Length);

                    Thread.Sleep(3000);

                    sp1.Write(close_cal, 0, close_cal.Length);

                    Thread.Sleep(3000);

                    sp15.Close();
                }
                catch { }
            }
        }



        private void init1()
        {
            sp1.Open();
            if (sp1.IsOpen)
            {
                try
                {
                    sp1.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(300);

                    sp1.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(300);

                    sp1.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(300);

                    sp1.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(300);

                    sp1.Close();
                }
                catch { }
            }
        }

        private void init2()
        {
            sp2.Open();
            if (sp2.IsOpen)
            {
                try
                {
                    sp2.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp2.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp2.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp2.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp2.Close();
                }
                catch { }
            }
        }

        private void init3()
        {
            sp3.Open();
            if (sp3.IsOpen)
            {
                try
                {
                    sp3.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp3.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp3.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp3.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp3.Close();
                }
                catch { }
            }
        }

        private void init4()
        {
            sp4.Open();
            if (sp4.IsOpen)
            {
                try
                {
                    sp4.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp4.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp4.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp4.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp4.Close();
                }
                catch { }
            }
        }

        private void init5()
        {
            sp5.Open();
            if (sp5.IsOpen)
            {
                try
                {
                    sp5.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp5.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp5.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp5.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp5.Close();
                }
                catch { }
            }
        }

        private void init6()
        {
            sp6.Open();
            if (sp6.IsOpen)
            {
                try
                {
                    sp6.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp6.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp6.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp6.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp6.Close();
                }
                catch { }
            }
        }

        private void init7()
        {
            sp7.Open();
            if (sp7.IsOpen)
            {
                try
                {
                    sp7.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp7.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp7.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp7.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp7.Close();
                }
                catch { }
            }
        }

        private void init8()
        {
            sp8.Open();
            if (sp8.IsOpen)
            {
                try
                {
                    sp8.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp8.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp8.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp8.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp8.Close();
                }
                catch { }
            }
        }

        private void init9()
        {
            sp9.Open();
            if (sp9.IsOpen)
            {
                try
                {
                    sp9.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp9.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp9.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp9.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp9.Close();
                }
                catch { }
            }
        }

        private void init10()
        {
            sp10.Open();
            if (sp10.IsOpen)
            {
                try
                {
                    sp10.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp10.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp10.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp10.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp10.Close();
                }
                catch { }
            }
        }

        private void init11()
        {
            sp11.Open();
            if (sp11.IsOpen)
            {
                try
                {
                    sp11.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp11.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp11.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp11.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp11.Close();
                }
                catch { }
            }
        }

        private void init12()
        {
            sp12.Open();
            if (sp12.IsOpen)
            {
                try
                {
                    sp12.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp12.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp12.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp12.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp12.Close();
                }
                catch { }
            }
        }

        private void init13()
        {
            sp13.Open();
            if (sp13.IsOpen)
            {
                try
                {
                    sp13.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp13.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp13.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp13.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp13.Close();
                }
                catch { }
            }
        }

        private void init14()
        {
            sp14.Open();
            if (sp14.IsOpen)
            {
                try
                {
                    sp14.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp14.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp14.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp14.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp14.Close();
                }
                catch { }
            }
        }

        private void init15()
        {
            sp15.Open();
            if (sp15.IsOpen)
            {
                try
                {
                    sp15.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp15.Write(changeDN25, 0, changeDN25.Length);

                    Thread.Sleep(3000);

                    sp15.Write(open_cal_old, 0, open_cal_old.Length);

                    Thread.Sleep(3000);

                    sp15.Write(CAL_INIT, 0, CAL_INIT.Length);

                    Thread.Sleep(3000);

                    sp15.Close();
                }
                catch { }
            }
        }

        private void btn_threadinit_Click(object sender, EventArgs e)
        {
            Thread i1 = new Thread(new ThreadStart(init1));
            i1.Start();
            Thread i2 = new Thread(new ThreadStart(init2));
            i2.Start();
            Thread i3 = new Thread(new ThreadStart(init3));
            i3.Start();
            Thread i4 = new Thread(new ThreadStart(init4));
            i4.Start();
            Thread i5 = new Thread(new ThreadStart(init5));
            i5.Start();
            Thread i6 = new Thread(new ThreadStart(init6));
            i6.Start();
            Thread i7 = new Thread(new ThreadStart(init7));
            i7.Start();
            Thread i8 = new Thread(new ThreadStart(init8));
            i8.Start();
            Thread i9 = new Thread(new ThreadStart(init9));
            i9.Start();
            Thread i10 = new Thread(new ThreadStart(init10));
            i10.Start();
            Thread i11 = new Thread(new ThreadStart(init11));
            i11.Start();
            Thread i12 = new Thread(new ThreadStart(init12));
            i12.Start();
            Thread i13 = new Thread(new ThreadStart(init13));
            i13.Start();
            Thread i14 = new Thread(new ThreadStart(init14));
            i14.Start();
            Thread i15 = new Thread(new ThreadStart(init15));
            i15.Start();
        }

        private void btn_threadwrite_Click(object sender, EventArgs e)
        {
            Thread w1 = new Thread(new ThreadStart(write1));
            w1.Start();
            Thread w2 = new Thread(new ThreadStart(write2));
            w2.Start();
            Thread w3 = new Thread(new ThreadStart(write3));
            w3.Start();
            Thread w4 = new Thread(new ThreadStart(write4));
            w4.Start();
            Thread w5 = new Thread(new ThreadStart(write5));
            w5.Start();
            Thread w6 = new Thread(new ThreadStart(write6));
            w6.Start();
            Thread w7 = new Thread(new ThreadStart(write7));
            w7.Start();
            Thread w8 = new Thread(new ThreadStart(write8));
            w8.Start();
            Thread w9 = new Thread(new ThreadStart(write9));
            w9.Start();
            Thread w10 = new Thread(new ThreadStart(write10));
            w10.Start();
            Thread w11 = new Thread(new ThreadStart(write11));
            w11.Start();
            Thread w12 = new Thread(new ThreadStart(write12));
            w12.Start();
            Thread w13 = new Thread(new ThreadStart(write13));
            w13.Start();
            Thread w14 = new Thread(new ThreadStart(write14));
            w14.Start();
            Thread w15 = new Thread(new ThreadStart(write15));
            w15.Start();
        }

        private void calibratedata_Load(object sender, EventArgs e)
        {
            //SerialPort spp = new SerialPort("COM1", 9600, Parity.None, 8, StopBits.One);

            //spp.Open();
            //if (spp.IsOpen)
            //{
            //    try
            //    {
            //        spp.Write(open_cal_old, 0, open_cal_old.Length);

            //        Thread.Sleep(300);

            //        spp.Write(changeDN25, 0, changeDN25.Length);

            //        Thread.Sleep(300);

            //        spp.Write(open_cal_old, 0, open_cal_old.Length);

            //        Thread.Sleep(300);

            //        spp.Write(CAL_INIT, 0, CAL_INIT.Length);

            //        Thread.Sleep(300);

            //        spp.Close();
            //    }
            //    catch { }
            //}
        }
    }
}
