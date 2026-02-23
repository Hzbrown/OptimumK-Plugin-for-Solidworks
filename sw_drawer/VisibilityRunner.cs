using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace sw_drawer
{
    class VisibilityRunner
    {
        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return 1;
            }

            string command = args[0].ToLower();
            string visibleArg = args[1].ToLower();
            bool visible = visibleArg == "true" || visibleArg == "1" || visibleArg == "show";

            try
            {
                SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                if (swApp == null)
                {
                    Console.WriteLine("Error: Could not connect to SolidWorks. Make sure it is running.");
                    return 1;
                }

                int result = 0;

                switch (command)
                {
                    case "feature":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: Feature name required.");
                            return 1;
                        }
                        bool success = FeatureVisibility.SetFeatureVisibility(swApp, args[2], visible);
                        result = success ? 1 : 0;
                        break;

                    case "suffix":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: Suffix required.");
                            return 1;
                        }
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, args[2], visible);
                        break;

                    case "substring":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: Substring required.");
                            return 1;
                        }
                        result = FeatureVisibility.SetFeatureVisibilityBySubstring(swApp, args[2], visible);
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
                        // Show/hide features containing "Chassis" in name
                        result = FeatureVisibility.SetFeatureVisibilityBySubstring(swApp, "Chassis", visible);
                        break;

                    case "nonchassis":
                        // Show/hide all suspension points that are NOT chassis points
                        result = SetNonChassisVisibility(swApp, visible);
                        break;

                    case "all":
                        result = FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_FRONT", visible);
                        result += FeatureVisibility.SetFeatureVisibilityBySuffix(swApp, "_REAR", visible);
                        result += FeatureVisibility.SetWheelVisibility(swApp, visible);
                        break;

                    default:
                        Console.WriteLine($"Unknown command: {command}");
                        PrintUsage();
                        return 1;
                }

                Console.WriteLine($"Total features modified: {result}");
                return result > 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
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
                bool isChassisPoint = name.Contains("Chassis");

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

        static void PrintUsage()
        {
            Console.WriteLine("Usage: VisibilityRunner.exe <command> <show|hide> [parameter]");
            Console.WriteLine();
            Console.WriteLine("Commands:");
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
