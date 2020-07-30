

using System;

namespace CertManager
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Application Started" + Environment.NewLine + "Press any key to continue");
            Console.ReadKey();
            new Manager().Execute();
            Console.WriteLine("Press Enter to Exit");
            Console.ReadLine();
        }
    }
}