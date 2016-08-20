using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DWGtoPdfCoverterTrueView
{
    class Program
    {
        static void Main(string[] args)
        {
            //string file = @"C:\Users\amitp\Desktop\csv - test.csv";
            //string file = @"C:\Users\amitp\Desktop\csv - test - Copy.csv";
            
            if (args.Length == 1)
            {
                if (args[0].Equals("?") || args[0].Equals("help") || args[0].Equals("-help"))
                {
                    Console.WriteLine("USAGE : <exe file name> <path to csv file>");
                    return;
                }

                DWGTrueViewConverter converter = new DWGTrueViewConverter(args[0]);
                converter.Convert();
                converter.Dispose();
            }
            else
            {
                Console.WriteLine("USAGE : <exe file name> <path to csv file>");
            }
        }
    }
}
