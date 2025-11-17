using System;

namespace ScannerAgent
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Startup.Start(); // starts HTTP listener
            Console.WriteLine("Press Ctrl+C to stop...");
            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
        }
    }
}
