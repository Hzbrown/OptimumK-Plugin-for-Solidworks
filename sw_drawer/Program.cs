using System;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return 1;
            }

            string command = args[0].ToLowerInvariant();

            try
            {
                bool success = false;

                switch (command)
                {
                    case "marker":
                        success = InsertMarker.Run(args);
                        break;

                    case "vis":
                        success = RunVisibilityCommand(args);
                        break;

                    case "hardpoints":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Usage: hardpoints <add|pose> [args]");
                            PrintUsage();
                            return 1;
                        }
                        
                        string subCommand = args[1].ToLowerInvariant();
                        
                        switch (subCommand)
                        {
                            case "add":
                                if (args.Length < 4)
                                {
                                    Console.WriteLine("Usage: hardpoints add <jsonPath> <markerPartPath>");
                                    PrintUsage();
                                    return 1;
                                }
                                success = HardpointRunner.RunAdd(args);
                                break;
                            case "pose":
                                if (args.Length < 4)
                                {
                                    Console.WriteLine("Usage: hardpoints pose <jsonPath> <configName>");
                                    PrintUsage();
                                    return 1;
                                }
                                success = HardpointRunner.RunPose(args);
                                break;
                            case "addwheels":
                                if (args.Length < 4)
                                {
                                    Console.WriteLine("Usage: hardpoints addwheels <jsonPath> <markerPartPath> [front|rear] [referenceDistanceMM]");
                                    PrintUsage();
                                    return 1;
                                }
                                success = HardpointRunner.RunAddWheels(args);
                                break;
                            default:
                                Console.WriteLine($"Unknown hardpoints command: {subCommand}");
                                PrintUsage();
                                return 1;
                        }
                        break;

                    default:
                        // Legacy: treat as coordinate system insertion
                        // Args: <name> <x> <y> <z> [rx] [ry] [rz]
                        if (args.Length >= 4)
                        {
                            success = InsertCoordinateSystem(args);
                        }
                        else
                        {
                            Console.WriteLine($"Unknown command: {command}");
                            PrintUsage();
                            return 1;
                        }
                        break;
                }

                return success ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("SuspensionTools.exe - SolidWorks Suspension Utilities");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  SuspensionTools.exe <name> <x> <y> <z> [rx] [ry] [rz]  - Insert coordinate system");
            Console.WriteLine("  SuspensionTools.exe marker <command> [args]            - Marker operations");
            Console.WriteLine("  SuspensionTools.exe vis <command> [args]               - Visibility control");
        }

        static bool InsertCoordinateSystem(string[] args)
        {
            string name = args[0];
            double x = double.Parse(args[1]) / 1000.0;  // mm to meters
            double y = double.Parse(args[2]) / 1000.0;
            double z = double.Parse(args[3]) / 1000.0;

            SldWorks swApp;
            try { swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); }
            catch { Console.WriteLine("Error: SolidWorks not running"); return false; }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null) { Console.WriteLine("Error: No active document"); return false; }

            // Create a 3D sketch point at the location, then use it for the coordinate system
            SketchManager sketchMgr = swModel.SketchManager;
            FeatureManager featMgr = swModel.FeatureManager;

            // Insert 3D sketch
            swModel.ClearSelection2(true);
            sketchMgr.Insert3DSketch(true);
            
            // Create point at location
            SketchPoint skPt = sketchMgr.CreatePoint(x, y, z);
            
            // Exit sketch
            sketchMgr.Insert3DSketch(true);
            
            // Get the sketch feature and rename it
            Feature sketchFeat = (Feature)swModel.FeatureByPositionReverse(0);
            string sketchName = name + "_RefSketch";
            if (sketchFeat != null)
            {
                sketchFeat.Name = sketchName;
            }

            // Select the point for coordinate system creation
            swModel.ClearSelection2(true);
            
            // Select the sketch point as origin
            bool selected = swModel.Extension.SelectByID2(
                "Point1@" + sketchName, "EXTSKETCHPOINT", x, y, z, false, 1, null,
                (int)swSelectOption_e.swSelectOptionDefault);

            if (!selected)
            {
                // Try selecting by coordinates
                selected = swModel.Extension.SelectByID2(
                    "", "EXTSKETCHPOINT", x, y, z, false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);
            }

            // Create coordinate system
            Feature csFeat = featMgr.InsertCoordinateSystem(false, false, false);
            
            if (csFeat != null)
            {
                csFeat.Name = name;
                Console.WriteLine($"Created: {name} at ({x*1000:F1}, {y*1000:F1}, {z*1000:F1}) mm");
                swModel.ClearSelection2(true);
                return true;
            }
            
            Console.WriteLine($"Failed to create: {name}");
            swModel.ClearSelection2(true);
            return false;
        }

        static bool RunVisibilityCommand(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: vis <all|front|rear|substring> <show|hide> [param]");
                return false;
            }

            string target = args[1].ToLowerInvariant();
            bool visible = args[2].ToLowerInvariant() == "show";
            string param = args.Length > 3 ? args[3] : null;

            SldWorks swApp;
            try { swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); }
            catch { Console.WriteLine("Error: SolidWorks not running"); return false; }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null) { Console.WriteLine("Error: No active document"); return false; }

            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                string name = feat.Name;
                bool shouldModify = false;

                switch (target)
                {
                    case "all":
                        shouldModify = name.Contains("_FRONT") || name.Contains("_REAR") || name.Contains("_wheel");
                        break;
                    case "front":
                        shouldModify = name.Contains("_FRONT") || name.StartsWith("FL_") || name.StartsWith("FR_");
                        break;
                    case "rear":
                        shouldModify = name.Contains("_REAR") || name.StartsWith("RL_") || name.StartsWith("RR_");
                        break;
                    case "wheels":
                        shouldModify = name.Contains("_wheel");
                        break;
                    case "frontwheels":
                        shouldModify = name.StartsWith("FL_") || name.StartsWith("FR_");
                        break;
                    case "rearwheels":
                        shouldModify = name.StartsWith("RL_") || name.StartsWith("RR_");
                        break;
                    case "chassis":
                        shouldModify = name.Contains("CHAS_");
                        break;
                    case "nonchassis":
                        shouldModify = (name.Contains("_FRONT") || name.Contains("_REAR")) && !name.Contains("CHAS_");
                        break;
                    case "substring":
                        shouldModify = param != null && name.IndexOf(param, StringComparison.OrdinalIgnoreCase) >= 0;
                        break;
                }

                if (shouldModify && feat.GetTypeName2() == "CoordSys")
                {
                    try
                    {
                        feat.SetSuppression2(
                            visible ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature
                                    : (int)swFeatureSuppressionAction_e.swSuppressFeature,
                            (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        count++;
                    }
                    catch { }
                }

                feat = (Feature)feat.GetNextFeature();
            }

            Console.WriteLine($"{count} features {(visible ? "shown" : "hidden")}");
            return true;
        }
    }
}
