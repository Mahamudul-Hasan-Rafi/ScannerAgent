using ScannerAgent.Controllers;
using System;


namespace ScannerAgent
{
    public class Startup
    {
        public static void Start()
        {
            string url = "http://localhost:9257/";

            // Initialize controller
            ScannerController scannerController = new ScannerController();

            // Initialize listener and pass controller
            ScannerHttpListener listener = new ScannerHttpListener(url, scannerController);
            listener.Start();

            Console.WriteLine("Scanner service started at " + url);
        }
    }
}
