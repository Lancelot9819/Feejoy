using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using feejoy_wojie.database;
using System.Threading;

namespace feejoy_wojie
{
    class weight_comm
    {
        SerialPort sp_weight = new SerialPort("COM1", 1200, Parity.Odd, 8, StopBits.One);

        byte[] read_value = new byte[50];
        List<byte> buffer = new List<byte>();
        byte FIRST = 0x2B;
        string sum;

        public void sp_open()
        {
            sp_weight.Dispose();
            if (sp_weight.IsOpen)
            {
                sp_weight.Close();
            }
            Thread.Sleep(50);
            if (!sp_weight.IsOpen)
            {
                sp_weight.Open();
            }
        }

        public void sp_close()
        {
            if (sp_weight.IsOpen)
            {
                sp_weight.Close();
            }
        }

        public void set_zero()
        {
            try
            {
                byte[] set_zero_buffer = new byte[2] { 0x1B, 0x54 };
                sp_weight.Write(set_zero_buffer, 0, set_zero_buffer.Length);
            }
            catch(Exception e)
            {

            }
        }

        public void read_weight()
        {
            sp_weight.DiscardInBuffer();

            Thread.Sleep(100);

            double d_weight;

            string weight= sp_weight.ReadLine().Replace("k", "").Replace("g", "").Replace(" ", "").Replace("+", "").Replace("-", "");

            if (weight.Contains("."))
            {
                try
                {
                    double.TryParse(weight, out d_weight);
                    plan_data.weight = d_weight.ToString();
                }
                catch (Exception e)
                {
                }
            }

            //sp_weight.Read(read_value, 0, read_value.Length);

            //buffer.AddRange(read_value);

            //int a = buffer.IndexOf(FIRST);

            //if (buffer[a + 1] != 0x20)
            //{
            //    string str5= (buffer[a + 1] - 0x30).ToString();
            //    sum = sum + str5;
            //}

            //if (buffer[a + 2] != 0x20)
            //{
            //    string str4 = (buffer[a + 2] - 0x30).ToString();
            //    sum = sum + str4;
            //}

            //if (buffer[a + 3] != 0x20)
            //{
            //    string str3 = (buffer[a + 3] - 0x30).ToString();
            //    sum = sum + str3;
            //}

            //if (buffer[a + 4] != 0x20)
            //{
            //    string str2 = (buffer[a + 4] - 0x30).ToString();
            //    sum = sum + str2;
            //}

            //if (buffer[a + 5] != 0x20)
            //{
            //    string str1 = (buffer[a + 5] - 0x30).ToString();
            //    sum = sum + str1;
            //}

            //if (buffer[a + 6] == 0x2E)
            //{
            //    sum = sum + ".";
            //}

            //if (buffer[a + 7] != 0x20)
            //{
            //    string str_1 = (buffer[a + 7] - 0x30).ToString();
            //    sum = sum + str_1;
            //}

            //if (buffer[a + 8] != 0x20)
            //{
            //    string str_2 = (buffer[a + 8] - 0x30).ToString();
            //    sum = sum + str_2;
            //}

            //if (buffer[a + 9] != 0x20)
            //{
            //    string str_3 = (buffer[a + 9] - 0x30).ToString();
            //    sum = sum + str_3;
            //}
            //plan_data.weight = sum;
        }
    }
}
