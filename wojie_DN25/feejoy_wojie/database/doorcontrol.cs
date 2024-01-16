using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;

namespace feejoy_wojie.database
{
    class doorcontrol
    {
        SerialPort sp = new SerialPort("COM2", 38400, Parity.None, 8, StopBits.One);
        byte[] crc;
        byte[] tx_data = new byte[8] { 0x01, 0x05, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00 };

        public static byte[] CRC(byte[] pByte)
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
        private void door_control(int ch,bool state)
        {
            switch (ch)
            {
                case 1:
                    tx_data[3] = 0x01;
                    break;
                case 2:
                    tx_data[3] = 0x02;
                    break;
                case 3:
                    tx_data[3] = 0x03;
                    break;
                case 4:
                    tx_data[3] = 0x04;
                    break;
                case 5:
                    tx_data[3] = 0x05;
                    break;
                default:
                    break;
            }
            if (state)
            {
                tx_data[4] = 0xFF;
            }
            else
            {
                tx_data[4] = 0x00;
            }

            try
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    crc = CRC(tx_data);
                    tx_data[6] = crc[1];
                    tx_data[7] = crc[0];
                    sp.Write(tx_data, 0, 8);
                }
                sp.Close();
            }
            catch { }
        }
    }
}
