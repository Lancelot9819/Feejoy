using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace feejoy_wojie.database
{
    class TCPIP
    {
        Socket tcpsocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        private static byte[] result = new byte[1024];
        public Socket socket = null;
        public EndPoint endpoint = null;
        
        public void send(string ip, int port, string send_data)
        {
            IPAddress ipaddress = IPAddress.Parse(ip);
            EndPoint point = new IPEndPoint(ipaddress, port);

            try
            {
                tcpsocket.Connect(point);
                tcpsocket.Send(Encoding .UTF8.GetBytes(send_data));
            }
            catch
            { }
        }     
    }
}
    
