using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.IO.Ports;

namespace feejoy_wojie.database
{
    class environment
    {
        SerialPort sp = new SerialPort("COM2", 19200, Parity.None, 8, StopBits.Two);
        Thread e1 = null;
        public void open_com()
        {
            if (sp.IsOpen)
            {
                sp.Close();
            }
            if (!sp.IsOpen)
            {
                sp.ReadTimeout = 500;
                sp.WriteTimeout = 1000;
                try
                {
                    if (sp.IsOpen)
                    {
                        sp.Close();
                    }
                    sp.Open();///////
                }
                catch
                {
                }

                e1 = new Thread(new ThreadStart(data_read));
                e1.Start();
            }
        }

        public void data_read()
        {
            while (true)
            {
                while (sp.IsOpen)
                {
                    byte[] txdata = new byte[8];
                    byte[] rxdata = new byte[100];
                    byte[] crc;
                    txdata[0] = 0xF0;
                    txdata[1] = 0x03;
                    txdata[2] = 0x00;
                    txdata[3] = 0x00;
                    txdata[4] = 0x00;
                    txdata[5] = 0x04;
                    crc = CRC(txdata);
                    txdata[6] = crc[1];
                    txdata[7] = crc[0];
                    sp.Write(txdata, 0, txdata.Length);

                    Thread.Sleep(500);

                    int count = sp.Read(rxdata, 0, sp.BytesToRead);
                    if (sp.IsOpen)
                    {
                        sp.DiscardOutBuffer();
                    }

                    Thread.Sleep(500);

                    byte[] rx_hum = { rxdata[4], rxdata[3], rxdata[6], rxdata[5] };
                    plan_data.hum = BitConverter.ToSingle(rx_hum, 0);
                    byte[] rx_temp = { rxdata[8], rxdata[7], rxdata[10], rxdata[9] };
                    plan_data.temp = BitConverter.ToSingle(rx_temp, 0);
                }
            }
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
    }
}
    
