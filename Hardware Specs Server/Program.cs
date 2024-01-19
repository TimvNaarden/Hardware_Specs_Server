using System;
using System.Net.Sockets;
using System.Net;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Win32;
using System.Reflection;
using Hardware_Specs_GUI.Json;
using System.Collections.Generic;

namespace Hardware_Specs_Server
{
    public class Program
    {
        public static void Main()
        {
            // Auto startup
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            rk.SetValue("Hardware Specs Client", Assembly.GetExecutingAssembly().Location);

            // Set the default port number to listen on
            int port = 12345;

            // Check the online configuration for an port number
            using (WebClient client = new WebClient())
            {
                try
                {
                    string data = client.DownloadString("https://raw.githubusercontent.com/TimvNaarden/Hardware_specs_client/main/index.json");
                    Dictionary<string, object> ob = data.FromJson<Dictionary<string, object>>();
                    if (!int.TryParse((string)ob["port"], out port))
                    {
                        Console.WriteLine($"Couldn't convert {ob["port"]} to int.");
                    }
                } catch
                {
                    Console.WriteLine("Can't get the information form the web server.");
                }
            }
            
            // Create a UDP client to receive data
            using (UdpClient udpClient = new UdpClient(port))
            {
                Console.WriteLine($"Listening on port {port}...");

                while (true)
                {
                    // Receive UDP packet and get the sender's IP address
                    IPEndPoint remoteEP = null;
                    byte[] data = udpClient.Receive(ref remoteEP);

                    // Convert the received data to a string
                    string receivedMessage = Encoding.ASCII.GetString(data);

                    // Log the received packet to the console
                    Console.WriteLine($"Received from {remoteEP}: {receivedMessage}");
                    Save2Excel.Save(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "computers.xlsx"), receivedMessage);
                }
            }
        }
    }
}
