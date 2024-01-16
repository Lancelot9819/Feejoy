using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace feejoy_wojie.database
{
    class plan_data
    {
        private static string _operatorname;

        private static string _b1order;
        private static string _b2order;
        private static string _b3order;
        private static string _b4order;
        private static string _b5order;
        private static string _b6order;
        private static string _b7order;
        private static string _b8order;
        private static string _b9order;
        private static string _b10order;
        private static string _b11order;

        private static string _pl1 = "111";
        public static string pl1
        {
            get { return _pl1; }
            set { _pl1 = value; }
        }
        private static string _pl2 = "222";
        public static string pl2
        {
            get { return _pl2; }
            set { _pl2 = value; }
        }
        private static string _pl3 = "333";
        public static string pl3
        {
            get { return _pl3; }
            set { _pl3 = value; }
        }

        private static string _kd1 = "333";
        public static string kd1
        {
            get { return _kd1; }
            set { _kd1 = value; }
        }
        private static string _kd2 = "222";
        public static string kd2
        {
            get { return _kd2; }
            set { _kd2 = value; }
        }
        private static string _kd3 = "111";

        public static string kd3
        {
            get { return _kd3; }
            set { _kd3 = value; }
        }

        private static string _stp;
        private static string _b1p;
        private static string _b2p;
        private static string _b3p;
        private static string _b4p;
        private static string _b5p;
        private static string _b6p;
        private static string _b7p;
        private static string _b8p;
        private static string _b9p;
        private static string _b10p;
        private static string _b11p;

        private static float _b1_flow1;
        private static float _b1_flow2;
        private static float _b1_flow3;
        private static float _b1_flow4;
        private static float _b1_flow5;

        private static float _b2_flow1;
        private static float _b2_flow2;
        private static float _b2_flow3;
        private static float _b2_flow4;
        private static float _b2_flow5;

        private static float _b3_flow1;
        private static float _b3_flow2;
        private static float _b3_flow3;
        private static float _b3_flow4;
        private static float _b3_flow5;


        private static float _b4_flow1;
        private static float _b4_flow2;
        private static float _b4_flow3;
        private static float _b4_flow4;
        private static float _b4_flow5;

        private static float _b5_flow1;
        private static float _b5_flow2;
        private static float _b5_flow3;
        private static float _b5_flow4;
        private static float _b5_flow5;

        private static float _b6_flow1;
        private static float _b6_flow2;
        private static float _b6_flow3;
        private static float _b6_flow4;
        private static float _b6_flow5;

        private static float _b7_flow1;
        private static float _b7_flow2;
        private static float _b7_flow3;
        private static float _b7_flow4;
        private static float _b7_flow5;

        private static float _b8_flow1;
        private static float _b8_flow2;
        private static float _b8_flow3;
        private static float _b8_flow4;
        private static float _b8_flow5;

        private static float _b9_flow1;
        private static float _b9_flow2;
        private static float _b9_flow3;
        private static float _b9_flow4;
        private static float _b9_flow5;

        private static float _b10_flow1;
        private static float _b10_flow2;
        private static float _b10_flow3;
        private static float _b10_flow4;
        private static float _b10_flow5;

        private static float _b11_flow1;
        private static float _b11_flow2;
        private static float _b11_flow3;
        private static float _b11_flow4;
        private static float _b11_flow5;

        private static float _b1_k1;
        private static float _b1_k2;
        private static float _b1_k3;
        private static float _b1_k4;
        private static float _b1_k5;

        private static float _b2_k1;
        private static float _b2_k2;
        private static float _b2_k3;
        private static float _b2_k4;
        private static float _b2_k5;

        private static float _b3_k1;
        private static float _b3_k2;
        private static float _b3_k3;
        private static float _b3_k4;
        private static float _b3_k5;

        private static float _b4_k1;
        private static float _b4_k2;
        private static float _b4_k3;
        private static float _b4_k4;
        private static float _b4_k5;

        private static float _b5_k1;
        private static float _b5_k2;
        private static float _b5_k3;
        private static float _b5_k4;
        private static float _b5_k5;

        private static float _b6_k1;
        private static float _b6_k2;
        private static float _b6_k3;
        private static float _b6_k4;
        private static float _b6_k5;

        private static float _b7_k1;
        private static float _b7_k2;
        private static float _b7_k3;
        private static float _b7_k4;
        private static float _b7_k5;

        private static float _b8_k1;
        private static float _b8_k2;
        private static float _b8_k3;
        private static float _b8_k4;
        private static float _b8_k5;

        private static float _b9_k1;
        private static float _b9_k2;
        private static float _b9_k3;
        private static float _b9_k4;
        private static float _b9_k5;

        private static float _b10_k1;
        private static float _b10_k2;
        private static float _b10_k3;
        private static float _b10_k4;
        private static float _b10_k5;

        private static float _b11_k1;
        private static float _b11_k2;
        private static float _b11_k3;
        private static float _b11_k4;
        private static float _b11_k5;

        private static Int16 _stabletime;
        private static Int16 _exec_count = 0;

        private static Int16 _time1 = 30;
        private static Int16 _time2 = 30;
        private static Int16 _time3 = 30;
        private static Int16 _time4 = 60;
        private static Int16 _time5 = 60;
        private static Int16 _time6 = 60;
        private static Int16 _time7 = 90;
        private static Int16 _time8 = 90;
        private static Int16 _time9 = 90;
        private static Int16 _time10 = 120;
        private static Int16 _time11 = 120;
        private static Int16 _time12 = 120;
        private static Int16 _time13 = 150;
        private static Int16 _time14 = 150;
        private static Int16 _time15 = 150;

        private static Int16 _freq1 = 1;
        private static Int16 _freq2 = 1;
        private static Int16 _freq3 = 1;
        private static Int16 _freq4 = 2;
        private static Int16 _freq5 = 2;
        private static Int16 _freq6 = 2;
        private static Int16 _freq7 = 3;
        private static Int16 _freq8 = 3;
        private static Int16 _freq9 = 3;
        private static Int16 _freq10 = 4;
        private static Int16 _freq11 = 4;
        private static Int16 _freq12 = 4;
        private static Int16 _freq13 = 5;
        private static Int16 _freq14 = 5;
        private static Int16 _freq15 = 5;

        private static Boolean _exec1 = true;
        private static Boolean _exec2 = true;
        private static Boolean _exec3 = true;
        private static Boolean _exec4;
        private static Boolean _exec5;
        private static Boolean _exec6;
        private static Boolean _exec7;
        private static Boolean _exec8;
        private static Boolean _exec9;
        private static Boolean _exec10;
        private static Boolean _exec11;
        private static Boolean _exec12;
        private static Boolean _exec13;
        private static Boolean _exec14;
        private static Boolean _exec15;

        private static double _temp;

        public static double temp
        {
            get { return _temp; }
            set { _temp = value; }
        }

        private static double _hum;

        public static double hum
        {
            get { return _hum; }
            set { _hum = value; }
        }

        public static string operatorname
        {
            get { return _operatorname; }
            set { _operatorname = value; }
        }

        public static string b1order
        {
            get { return _b1order; }
            set { _b1order = value; }
        }

        public static string b2order
        {
            get { return _b2order; }
            set { _b2order = value; }
        }

        public static string b3order
        {
            get { return _b3order; }
            set { _b3order = value; }
        }

        public static string b4order
        {
            get { return _b4order; }
            set { _b4order = value; }
        }

        public static string b5order
        {
            get { return _b5order; }
            set { _b5order = value; }
        }

        public static string b6order
        {
            get { return _b6order; }
            set { _b6order = value; }
        }

        public static string b7order
        {
            get { return _b7order; }
            set { _b7order = value; }
        }

        public static string b8order
        {
            get { return _b8order; }
            set { _b8order = value; }
        }

        public static string b9order
        {
            get { return _b9order; }
            set { _b9order = value; }
        }

        public static string b10order
        {
            get { return _b10order; }
            set { _b10order = value; }
        }
        public static string b11order
        {
            get { return _b11order; }
            set { _b11order = value; }
        }

        public static string stp
        {
            get { return _stp; }
            set { _stp = value; }
        }

        public static string b1p
        {
            get { return _b1p; }
            set { _b1p = value; }
        }

        public static string b2p
        {
            get { return _b2p; }
            set { _b2p = value; }
        }
        public static string b3p
        {
            get { return _b3p; }
            set { _b3p = value; }
        }
        public static string b4p
        {
            get { return _b4p; }
            set { _b4p = value; }
        }
        public static string b5p
        {
            get { return _b5p; }
            set { _b5p = value; }
        }
        public static string b6p
        {
            get { return _b6p; }
            set { _b6p = value; }
        }
        public static string b7p
        {
            get { return _b7p; }
            set { _b7p = value; }
        }
        public static string b8p
        {
            get { return _b8p; }
            set { _b8p = value; }
        }
        public static string b9p
        {
            get { return _b9p; }
            set { _b9p = value; }
        }
        public static string b10p
        {
            get { return _b10p; }
            set { _b10p = value; }
        }
        public static string b11p
        {
            get { return _b11p; }
            set { _b11p = value; }
        }
        public static float b1_flow1
        {
            get { return _b1_flow1; }
            set { _b1_flow1 = value; }
        }

        public static float b1_flow2
        {
            get { return _b1_flow2; }
            set { _b1_flow2 = value; }
        }

        public static float b1_flow3
        {
            get { return _b1_flow3; }
            set { _b1_flow3 = value; }
        }

        public static float b1_flow4
        {
            get { return _b1_flow4; }
            set { _b1_flow4 = value; }
        }

        public static float b1_flow5
        {
            get { return _b1_flow5; }
            set { _b1_flow5 = value; }
        }

        public static float b2_flow1
        {
            get { return _b2_flow1; }
            set { _b2_flow1 = value; }
        }

        public static float b2_flow2
        {
            get { return _b2_flow2; }
            set { _b2_flow2 = value; }
        }

        public static float b2_flow3
        {
            get { return _b2_flow3; }
            set { _b2_flow3 = value; }
        }

        public static float b2_flow4
        {
            get { return _b2_flow4; }
            set { _b2_flow4 = value; }
        }

        public static float b2_flow5
        {
            get { return _b2_flow5; }
            set { _b2_flow5 = value; }
        }

        public static float b3_flow1
        {
            get { return _b3_flow1; }
            set { _b3_flow1 = value; }
        }

        public static float b3_flow2
        {
            get { return _b3_flow2; }
            set { _b3_flow2 = value; }
        }

        public static float b3_flow3
        {
            get { return _b3_flow3; }
            set { _b3_flow3 = value; }
        }

        public static float b3_flow4
        {
            get { return _b3_flow4; }
            set { _b3_flow4 = value; }
        }

        public static float b3_flow5
        {
            get { return _b3_flow5; }
            set { _b3_flow5 = value; }
        }

        public static float b4_flow1
        {
            get { return _b4_flow1; }
            set { _b4_flow1 = value; }
        }

        public static float b4_flow2
        {
            get { return _b4_flow2; }
            set { _b4_flow2 = value; }
        }

        public static float b4_flow3
        {
            get { return _b4_flow3; }
            set { _b4_flow3 = value; }
        }

        public static float b4_flow4
        {
            get { return _b4_flow4; }
            set { _b4_flow4 = value; }
        }

        public static float b4_flow5
        {
            get { return _b4_flow5; }
            set { _b4_flow5 = value; }
        }

        public static float b5_flow1
        {
            get { return _b5_flow1; }
            set { _b5_flow1 = value; }
        }

        public static float b5_flow2
        {
            get { return _b5_flow2; }
            set { _b5_flow2 = value; }
        }

        public static float b5_flow3
        {
            get { return _b5_flow3; }
            set { _b5_flow3 = value; }
        }

        public static float b5_flow4
        {
            get { return _b5_flow4; }
            set { _b5_flow4 = value; }
        }

        public static float b5_flow5
        {
            get { return _b5_flow5; }
            set { _b5_flow5 = value; }
        }

        public static float b6_flow1
        {
            get { return _b6_flow1; }
            set { _b6_flow1 = value; }
        }

        public static float b6_flow2
        {
            get { return _b6_flow2; }
            set { _b6_flow2 = value; }
        }

        public static float b6_flow3
        {
            get { return _b6_flow3; }
            set { _b6_flow3 = value; }
        }

        public static float b6_flow4
        {
            get { return _b6_flow4; }
            set { _b6_flow4 = value; }
        }

        public static float b6_flow5
        {
            get { return _b6_flow5; }
            set { _b6_flow5 = value; }
        }

        public static float b7_flow1
        {
            get { return _b7_flow1; }
            set { _b7_flow1 = value; }
        }

        public static float b7_flow2
        {
            get { return _b7_flow2; }
            set { _b7_flow2 = value; }
        }

        public static float b7_flow3
        {
            get { return _b7_flow3; }
            set { _b7_flow3 = value; }
        }

        public static float b7_flow4
        {
            get { return _b7_flow4; }
            set { _b7_flow4 = value; }
        }

        public static float b7_flow5
        {
            get { return _b7_flow5; }
            set { _b7_flow5 = value; }
        }

        public static float b8_flow1
        {
            get { return _b8_flow1; }
            set { _b8_flow1 = value; }
        }

        public static float b8_flow2
        {
            get { return _b8_flow2; }
            set { _b8_flow2 = value; }
        }

        public static float b8_flow3
        {
            get { return _b8_flow3; }
            set { _b8_flow3 = value; }
        }

        public static float b8_flow4
        {
            get { return _b8_flow4; }
            set { _b8_flow4 = value; }
        }

        public static float b8_flow5
        {
            get { return _b8_flow5; }
            set { _b8_flow5 = value; }
        }
        public static float b9_flow1
        {
            get { return _b9_flow1; }
            set { _b9_flow1 = value; }
        }

        public static float b9_flow2
        {
            get { return _b9_flow2; }
            set { _b9_flow2 = value; }
        }

        public static float b9_flow3
        {
            get { return _b9_flow3; }
            set { _b9_flow3 = value; }
        }

        public static float b9_flow4
        {
            get { return _b9_flow4; }
            set { _b9_flow4 = value; }
        }

        public static float b9_flow5
        {
            get { return _b9_flow5; }
            set { _b9_flow5 = value; }
        }

        public static float b10_flow1
        {
            get { return _b10_flow1; }
            set { _b10_flow1 = value; }
        }

        public static float b10_flow2
        {
            get { return _b10_flow2; }
            set { _b10_flow2 = value; }
        }

        public static float b10_flow3
        {
            get { return _b10_flow3; }
            set { _b10_flow3 = value; }
        }

        public static float b10_flow4
        {
            get { return _b10_flow4; }
            set { _b10_flow4 = value; }
        }

        public static float b10_flow5
        {
            get { return _b10_flow5; }
            set { _b10_flow5 = value; }
        }

        public static float b11_flow1
        {
            get { return _b11_flow1; }
            set { _b11_flow1 = value; }
        }

        public static float b11_flow2
        {
            get { return _b11_flow2; }
            set { _b11_flow2 = value; }
        }

        public static float b11_flow3
        {
            get { return _b11_flow3; }
            set { _b11_flow3 = value; }
        }

        public static float b11_flow4
        {
            get { return _b11_flow4; }
            set { _b11_flow4 = value; }
        }

        public static float b11_flow5
        {
            get { return _b11_flow5; }
            set { _b11_flow5 = value; }
        }

        public static float b1_k1
        {
            get { return _b1_k1; }
            set { _b1_k1 = value; }
        }

        public static float b1_k2
        {
            get { return _b1_k2; }
            set { _b1_k2 = value; }
        }

        public static float b1_k3
        {
            get { return _b1_k3; }
            set { _b1_k3 = value; }
        }

        public static float b1_k4
        {
            get { return _b1_k4; }
            set { _b1_k4 = value; }
        }

        public static float b1_k5
        {
            get { return _b1_k5; }
            set { _b1_k5 = value; }
        }

        public static float b2_k1
        {
            get { return _b2_k1; }
            set { _b2_k1 = value; }
        }

        public static float b2_k2
        {
            get { return _b2_k2; }
            set { _b2_k2 = value; }
        }

        public static float b2_k3
        {
            get { return _b2_k3; }
            set { _b2_k3 = value; }
        }

        public static float b2_k4
        {
            get { return _b2_k4; }
            set { _b2_k4 = value; }
        }

        public static float b2_k5
        {
            get { return _b2_k5; }
            set { _b2_k5 = value; }
        }

        public static float b3_k1
        {
            get { return _b3_k1; }
            set { _b3_k1 = value; }
        }

        public static float b3_k2
        {
            get { return _b3_k2; }
            set { _b3_k2 = value; }
        }

        public static float b3_k3
        {
            get { return _b3_k3; }
            set { _b3_k3 = value; }
        }

        public static float b3_k4
        {
            get { return _b3_k4; }
            set { _b3_k4 = value; }
        }

        public static float b3_k5
        {
            get { return _b3_k5; }
            set { _b3_k5 = value; }
        }

        public static float b4_k1
        {
            get { return _b4_k1; }
            set { _b4_k1 = value; }
        }

        public static float b4_k2
        {
            get { return _b4_k2; }
            set { _b4_k2 = value; }
        }

        public static float b4_k3
        {
            get { return _b4_k3; }
            set { _b4_k3 = value; }
        }

        public static float b4_k4
        {
            get { return _b4_k4; }
            set { _b4_k4 = value; }
        }

        public static float b4_k5
        {
            get { return _b4_k5; }
            set { _b4_k5 = value; }
        }

        public static float b5_k1
        {
            get { return _b5_k1; }
            set { _b5_k1 = value; }
        }

        public static float b5_k2
        {
            get { return _b5_k2; }
            set { _b5_k2 = value; }
        }

        public static float b5_k3
        {
            get { return _b5_k3; }
            set { _b5_k3 = value; }
        }

        public static float b5_k4
        {
            get { return _b5_k4; }
            set { _b5_k4 = value; }
        }

        public static float b5_k5
        {
            get { return _b5_k5; }
            set { _b5_k5 = value; }
        }

        public static float b6_k1
        {
            get { return _b6_k1; }
            set { _b6_k1 = value; }
        }

        public static float b6_k2
        {
            get { return _b6_k2; }
            set { _b6_k2 = value; }
        }

        public static float b6_k3
        {
            get { return _b6_k3; }
            set { _b6_k3 = value; }
        }

        public static float b6_k4
        {
            get { return _b6_k4; }
            set { _b6_k4 = value; }
        }

        public static float b6_k5
        {
            get { return _b6_k5; }
            set { _b6_k5 = value; }
        }

        public static float b7_k1
        {
            get { return _b7_k1; }
            set { _b7_k1 = value; }
        }

        public static float b7_k2
        {
            get { return _b7_k2; }
            set { _b7_k2 = value; }
        }

        public static float b7_k3
        {
            get { return _b7_k3; }
            set { _b7_k3 = value; }
        }

        public static float b7_k4
        {
            get { return _b7_k4; }
            set { _b7_k4 = value; }
        }

        public static float b7_k5
        {
            get { return _b7_k5; }
            set { _b7_k5 = value; }
        }

        public static float b8_k1
        {
            get { return _b8_k1; }
            set { _b8_k1 = value; }
        }

        public static float b8_k2
        {
            get { return _b8_k2; }
            set { _b8_k2 = value; }
        }

        public static float b8_k3
        {
            get { return _b8_k3; }
            set { _b8_k3 = value; }
        }

        public static float b8_k4
        {
            get { return _b8_k4; }
            set { _b8_k4 = value; }
        }

        public static float b8_k5
        {
            get { return _b8_k5; }
            set { _b8_k5 = value; }
        }

        public static float b9_k1
        {
            get { return _b9_k1; }
            set { _b9_k1 = value; }
        }

        public static float b9_k2
        {
            get { return _b9_k2; }
            set { _b9_k2 = value; }
        }

        public static float b9_k3
        {
            get { return _b9_k3; }
            set { _b9_k3 = value; }
        }

        public static float b9_k4
        {
            get { return _b9_k4; }
            set { _b9_k4 = value; }
        }

        public static float b9_k5
        {
            get { return _b9_k5; }
            set { _b9_k5 = value; }
        }

        public static float b10_k1
        {
            get { return _b10_k1; }
            set { _b10_k1 = value; }
        }

        public static float b10_k2
        {
            get { return _b10_k2; }
            set { _b10_k2 = value; }
        }

        public static float b10_k3
        {
            get { return _b10_k3; }
            set { _b10_k3 = value; }
        }

        public static float b10_k4
        {
            get { return _b10_k4; }
            set { _b10_k4 = value; }
        }

        public static float b10_k5
        {
            get { return _b10_k5; }
            set { _b10_k5 = value; }
        }

        public static float b11_k1
        {
            get { return _b11_k1; }
            set { _b11_k1 = value; }
        }

        public static float b11_k2
        {
            get { return _b11_k2; }
            set { _b11_k2 = value; }
        }

        public static float b11_k3
        {
            get { return _b11_k3; }
            set { _b11_k3 = value; }
        }

        public static float b11_k4
        {
            get { return _b11_k4; }
            set { _b11_k4 = value; }
        }

        public static float b11_k5
        {
            get { return _b11_k5; }
            set { _b11_k5 = value; }
        }

        public static Int16 stable_time
        {
            get { return _stabletime; }
            set { _stabletime = value; }
        }

        public static Int16 exec_count
        {
            get { return _exec_count; }
            set { _exec_count = value; }
        }

        public static Int16 time1
        {
            get { return _time1; }
            set { _time1 = value; }
        }
        public static Int16 time2
        {
            get { return _time2; }
            set { _time2 = value; }
        }
        public static Int16 time3
        {
            get { return _time3; }
            set { _time3 = value; }
        }
        public static Int16 time4
        {
            get { return _time4; }
            set { _time4 = value; }
        }
        public static Int16 time5
        {
            get { return _time5; }
            set { _time5 = value; }
        }
        public static Int16 time6
        {
            get { return _time6; }
            set { _time6 = value; }
        }
        public static Int16 time7
        {
            get { return _time7; }
            set { _time7 = value; }
        }
        public static Int16 time8
        {
            get { return _time8; }
            set { _time8 = value; }
        }
        public static Int16 time9
        {
            get { return _time9; }
            set { _time9 = value; }
        }
        public static Int16 time10
        {
            get { return _time10; }
            set { _time10 = value; }
        }
        public static Int16 time11
        {
            get { return _time11; }
            set { _time11 = value; }
        }
        public static Int16 time12
        {
            get { return _time12; }
            set { _time12 = value; }
        }
        public static Int16 time13
        {
            get { return _time13; }
            set { _time13 = value; }
        }
        public static Int16 time14
        {
            get { return _time14; }
            set { _time14 = value; }
        }
        public static Int16 time15
        {
            get { return _time15; }
            set { _time15 = value; }
        }

        public static Int16 freq1
        {
            get { return _freq1; }
            set { _freq1 = value; }
        }
        public static Int16 freq2
        {
            get { return _freq2; }
            set { _freq2 = value; }
        }
        public static Int16 freq3
        {
            get { return _freq3; }
            set { _freq3 = value; }
        }
        public static Int16 freq4
        {
            get { return _freq4; }
            set { _freq4 = value; }
        }
        public static Int16 freq5
        {
            get { return _freq5; }
            set { _freq5 = value; }
        }
        public static Int16 freq6
        {
            get { return _freq6; }
            set { _freq6 = value; }
        }
        public static Int16 freq7
        {
            get { return _freq7; }
            set { _freq7 = value; }
        }
        public static Int16 freq8
        {
            get { return _freq8; }
            set { _freq8 = value; }
        }
        public static Int16 freq9
        {
            get { return _freq9; }
            set { _freq9 = value; }
        }
        public static Int16 freq10
        {
            get { return _freq10; }
            set { _freq10 = value; }
        }
        public static Int16 freq11
        {
            get { return _freq11; }
            set { _freq11 = value; }
        }
        public static Int16 freq12
        {
            get { return _freq12; }
            set { _freq12 = value; }
        }
        public static Int16 freq13
        {
            get { return _freq13; }
            set { _freq13 = value; }
        }
        public static Int16 freq14
        {
            get { return _freq14; }
            set { _freq14 = value; }
        }
        public static Int16 freq15
        {
            get { return _freq15; }
            set { _freq15 = value; }
        }

        public static Boolean exec1
        {
            get { return _exec1; }
            set { _exec1 = value; }
        }
        public static Boolean exec2
        {
            get { return _exec2; }
            set { _exec2 = value; }
        }
        public static Boolean exec3
        {
            get { return _exec3; }
            set { _exec3 = value; }
        }
        public static Boolean exec4
        {
            get { return _exec4; }
            set { _exec4 = value; }
        }
        public static Boolean exec5
        {
            get { return _exec5; }
            set { _exec5 = value; }
        }
        public static Boolean exec6
        {
            get { return _exec6; }
            set { _exec6 = value; }
        }
        public static Boolean exec7
        {
            get { return _exec7; }
            set { _exec7 = value; }
        }
        public static Boolean exec8
        {
            get { return _exec8; }
            set { _exec8 = value; }
        }
        public static Boolean exec9
        {
            get { return _exec9; }
            set { _exec9 = value; }
        }
        public static Boolean exec10
        {
            get { return _exec10; }
            set { _exec10 = value; }
        }
        public static Boolean exec11
        {
            get { return _exec11; }
            set { _exec11 = value; }
        }
        public static Boolean exec12
        {
            get { return _exec12; }
            set { _exec12 = value; }
        }
        public static Boolean exec13
        {
            get { return _exec13; }
            set { _exec13 = value; }
        }
        public static Boolean exec14
        {
            get { return _exec14; }
            set { _exec14 = value; }
        }
        public static Boolean exec15
        {
            get { return _exec15; }
            set { _exec15 = value; }
        }
    }
}
