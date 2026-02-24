using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace sw_drawer
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return 1;
            }

            string command = args[0].ToLower();

            // Handle marker commands
            if (command == "marker")
            {
                return InsertMarker.Run(args) ? 0 : 1;
            }

            // Route to appropriate handler
            if (command == "coord" || command == "coordinate")
            {
                return HandleCoordinate(args);
            }
            else if (command == "vis" || command == "visibility")
            {
                return HandleVisibility(args);
            }
            else
            {
                // Legacy support: if first arg looks like a coordinate name, treat as coordinate command
                if (args.Length >= 4 && double.TryParse(args[1], out _))
                {
                    return HandleCoordinateLegacy(args);
                }
                
                Console.WriteLine($"Unknown command: {command}");
                PrintUsage();
                return 1;
            }
        }

        static int HandleCoordinate(string[] args)
        {
            // coord <name> <x> <y> <z> [angleX] [angleY] [angleZ]
            if (args.Length < 5)
            {
                Console.WriteLine("Usage: SuspensionTools coord <name> <x> <y> <z> [angleX] [angleY] [angleZ]");
                return 1;
            }

            string name = args[1];
            double x = double.Parse(args[2]);
            double y = double.Parse(args[3]);
            double z = double.Parse(args[4]);
            double angleX = args.Length > 5 ? double.Parse(args[5]) : 0;
            double angleY = args.Length > 6 ? double.Parse(args[6]) : 0;
            double angleZ = args.Length > 7 ? double.Parse(args[7]) : 0;

            return ExecuteWithSolidWorks(swApp => 
                InsertCoordinate.InsertCoordinateSystem(swApp, name, x, y, z, angleX, angleY, angleZ) ? 0 : 1);
        }

        static int HandleCoordinateLegacy(string[] args)
        {
            // Legacy: <name> <x> <y> <z> [angleX] [angleY] [angleZ]
            string name = args[0];
            double x = double.Parse(args[1]);
            double y = double.Parse(args[2]);
            double z = double.Parse(args[3]);
            double angleX = args.Length > 4 ? double.Parse(args[4]) : 0;
            double angleY = args.Length > 5 ? double.Parse(args[5]) : 0;
            double angleZ = args.Length > 6 ? double.Parse(args[6]) : 0;

            return ExecuteWithSolidWorks(swApp => 
                InsertCoordinate.InsertCoordinateSystem(swApp, name, x, y, z, angleX, angleY, angleZ) ? 0 : 1);
        }

        static int HandleVisibility(string[] args)
        {
            // vis <subcommand> <show|hide> [parameter]
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: SuspensionTools vis <subcommand> <show|hide> [parameter]");
                PrintVisibilityCommands();
                return 1;
            }

            string subCommand = args[1].ToLower();
            string visibleArg = args[2].ToLower();
            bool visible = visibleArg == "true" || visibleArg == "1" || visibleArg == "show";

            return ExecuteWithSolidWorks(swApp =>
            {
                int result = 0;

                switch (subCommand)
                {
                    case "feature":
                        if (args.Length < 4)
                        {
                            Console.WriteLine("Error: Feature name required.");
                            return 1;
                        }
                        bool success = FeatureVisibility.SetFeatureVisibility(swApp, args[3], visible);
                        result = success ? 1 : 0;
                        break;

                    case "suffix":
                        if (args.Length < 4)
                        {
                            Console.WriteLine("Error: Suffix required.");
                            return 1;
                        }
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, args[3], visible);
                        break;

                    case "substring":
                        if (args.Length < 4)
                        {
                            Console.WriteLine("Error: Substring required.");
                            return 1;
                        }
                        result = FeatureVisibility.SetFeatureVisibilityBySubstring(swApp, args[3], visible);
                        break;

                    case "front":
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_FRONT", visible);
                        result += FeatureVisibility.SetWheelVisibility(swApp, visible, frontOnly: true);
                        break;

                    case "rear":
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_REAR", visible);
                        result += FeatureVisibility.SetWheelVisibility(swApp, visible, rearOnly: true);
                        break;

                    case "wheels":
                        result = FeatureVisibility.SetWheelVisibility(swApp, visible);
                        break;

                    case "frontwheels":
                        result = FeatureVisibility.SetWheelVisibility(swApp, visible, frontOnly: true);
                        break;

                    case "rearwheels":
                        result = FeatureVisibility.SetWheelVisibility(swApp, visible, rearOnly: true);
                        break;

                    case "chassis":
                        result = FeatureVisibility.SetFeatureVisibilityBySubstring(swApp, "CHAS", visible);
                        break;

                    case "nonchassis":
                        result = SetNonChassisVisibility(swApp, visible);
                        break;

                    case "all":
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_FRONT", visible);
                        result += FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_REAR", visible);
                        result += FeatureVisibility.SetWheelVisibility(swApp, visible);
                        break;

                    default:
                        Console.WriteLine($"Unknown visibility command: {subCommand}");
                        PrintVisibilityCommands();
                        return 1;
                }

                Console.WriteLine($"Total features modified: {result}");
                return result > 0 ? 0 : 1;
            });
        }

        static int SetNonChassisVisibility(SldWorks swApp, bool visible)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null) return 0;

            int count = 0;
            Feature swFeature = (Feature)swModel.FirstFeature();

            while (swFeature != null)
            {
                string name = swFeature.Name;
                bool isSuspensionPoint = name.EndsWith("_FRONT") || name.EndsWith("_REAR") || name.Contains("_wheel");
                bool isChassisPoint = name.IndexOf("CHAS", StringComparison.OrdinalIgnoreCase) >= 0;

                if (isSuspensionPoint && !isChassisPoint)
                {
                    if (FeatureVisibility.SetFeatureVisibility(swApp, name, visible))
                    {
                        count++;
                    }
                }
                swFeature = (Feature)swFeature.GetNextFeature();
            }

            Console.WriteLine($"{(visible ? "Shown" : "Hidden")} {count} non-chassis features.");
            return count;
        }

        static int ExecuteWithSolidWorks(Func<SldWorks, int> action)
        {
            try
            {
                SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                if (swApp == null)
                {
                    Console.WriteLine("Error: Could not connect to SolidWorks. Make sure it is running.");
                    return 1;
                }
                return action(swApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("SuspensionTools - SolidWorks Suspension Geometry Tool");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  SuspensionTools coord <name> <x> <y> <z> [angleX] [angleY] [angleZ]");
            Console.WriteLine("  SuspensionTools vis <command> <show|hide> [parameter]");
            Console.WriteLine();
            Console.WriteLine("Legacy (backward compatible):");
            Console.WriteLine("  SuspensionTools <name> <x> <y> <z> [angleX] [angleY] [angleZ]");
            Console.WriteLine();
            PrintVisibilityCommands();
            Console.WriteLine("  marker create <csname> <radius_mm>  - Create marker at coordinate system");
            Console.WriteLine("  marker createall <radius_mm>        - Create markers at all coordinate systems");
            Console.WriteLine("  marker deleteall                    - Delete all markers");
            Console.WriteLine("  marker vis <all|group|front|rear|name> <show|hide> [param]");
        }

        static void PrintVisibilityCommands()
        {
            Console.WriteLine("Visibility Commands:");
            Console.WriteLine("  feature <show|hide> <name>     - Show/hide a specific feature");
            Console.WriteLine("  suffix <show|hide> <suffix>    - Show/hide features by suffix");
            Console.WriteLine("  substring <show|hide> <text>   - Show/hide features containing text");
            Console.WriteLine("  front <show|hide>              - Show/hide all front suspension");
            Console.WriteLine("  rear <show|hide>               - Show/hide all rear suspension");
            Console.WriteLine("  wheels <show|hide>             - Show/hide all wheels");
            Console.WriteLine("  frontwheels <show|hide>        - Show/hide front wheels only");
            Console.WriteLine("  rearwheels <show|hide>         - Show/hide rear wheels only");
            Console.WriteLine("  chassis <show|hide>            - Show/hide chassis points");
            Console.WriteLine("  nonchassis <show|hide>         - Show/hide non-chassis points");
            Console.WriteLine("  all <show|hide>                - Show/hide all suspension features");
        }
    }
}
