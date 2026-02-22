using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: CoordinateRunner.exe <name> <x> <y> <z> [angleX] [angleY] [angleZ]");
                return 1;
            }

            string name = args[0];
            double x = double.Parse(args[1]);
            double y = double.Parse(args[2]);
            double z = double.Parse(args[3]);
            double angleX = args.Length > 4 ? double.Parse(args[4]) : 0;
            double angleY = args.Length > 5 ? double.Parse(args[5]) : 0;
            double angleZ = args.Length > 6 ? double.Parse(args[6]) : 0;

            try
            {
                SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                if (swApp == null)
                {
                    Console.WriteLine("Error: Could not connect to SolidWorks. Make sure it is running.");
                    return 1;
                }

                bool success = InsertCoordinate.InsertCoordinateSystem(swApp, name, x, y, z, angleX, angleY, angleZ);
                return success ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }
    }
}
