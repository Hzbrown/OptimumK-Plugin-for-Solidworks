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
            bool shouldEnsureReleased = command == "marker" || command == "hardpoints" || command == "release";

            try
            {
                bool success = false;

                switch (command)
                {
                    case "release":
                        success = RunReleaseCommand();
                        break;

                    case "marker":
                        success = InsertMarker.Run(args);
                        break;

                    case "vis":
                        success = RunVisibilityCommand(args);
                        break;

                    case "hardpoints":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Usage: hardpoints <add|pose|insertpose> [args]");
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
                            case "insertpose":
                                if (args.Length < 4)
                                {
                                    Console.WriteLine("Usage: hardpoints insertpose <jsonPath> <poseName>");
                                    PrintUsage();
                                    return 1;
                                }
                                success = HardpointRunner.RunInsertPose(args);
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
            finally
            {
                if (shouldEnsureReleased)
                {
                    TryReleaseSolidWorksCommandState(false);
                }
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
            Console.WriteLine("  SuspensionTools.exe release                            - Release SolidWorks command state");
        }

        static bool RunReleaseCommand()
        {
            return TryReleaseSolidWorksCommandState(true);
        }

        static bool TryReleaseSolidWorksCommandState(bool verbose)
        {
            SldWorks swApp;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                if (verbose)
                {
                    Console.WriteLine("Error: SolidWorks not running");
                }
                return false;
            }

            bool success = true;

            try
            {
                swApp.CommandInProgress = false;
                if (verbose)
                {
                    Console.WriteLine("SolidWorks CommandInProgress reset");
                }
            }
            catch (Exception ex)
            {
                success = false;
                if (verbose)
                {
                    Console.WriteLine($"Warning: Failed to reset CommandInProgress: {ex.Message}");
                }
            }

            try
            {
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel != null)
                {
                    try { swModel.ClearSelection2(true); } catch { }

                    try
                    {
                        if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                        {
                            AssemblyDoc swAssy = (AssemblyDoc)swModel;
                            swAssy.EditAssembly();
                        }
                    }
                    catch { }

                    try { swModel.EditRebuild3(); } catch { }
                }
            }
            catch (Exception ex)
            {
                success = false;
                if (verbose)
                {
                    Console.WriteLine($"Warning: Failed to normalize active document state: {ex.Message}");
                }
            }

            return success;
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
                Console.WriteLine("Usage: vis <all|front|rear|wheels|frontwheels|rearwheels|chassis|nonchassis|substring|feature> <show|hide> [param]");
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

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Console.WriteLine("Error: Active document must be an assembly for component visibility control");
                return false;
            }

            AssemblyDoc swAssy = (AssemblyDoc)swModel;

            if (target == "substring" && string.IsNullOrWhiteSpace(param))
            {
                Console.WriteLine("Usage: vis substring <show|hide> <text>");
                return false;
            }

            if (target == "feature" && string.IsNullOrWhiteSpace(param))
            {
                Console.WriteLine("Usage: vis feature <show|hide> <name>");
                return false;
            }

            Func<string, bool> matcher;
            switch (target)
            {
                case "all":
                    matcher = IsHardpointLikeName;
                    break;
                case "front":
                    matcher = name => IsFrontLikeName(name);
                    break;
                case "rear":
                    matcher = name => IsRearLikeName(name);
                    break;
                case "wheels":
                    matcher = name => IsWheelLikeName(name);
                    break;
                case "frontwheels":
                    matcher = name => IsWheelLikeName(name) && IsFrontLikeName(name);
                    break;
                case "rearwheels":
                    matcher = name => IsWheelLikeName(name) && IsRearLikeName(name);
                    break;
                case "chassis":
                    matcher = name => IsChassisLikeName(name);
                    break;
                case "nonchassis":
                    matcher = name => IsHardpointLikeName(name) && !IsChassisLikeName(name);
                    break;
                case "substring":
                    matcher = name => name.IndexOf(param, StringComparison.OrdinalIgnoreCase) >= 0;
                    break;
                case "feature":
                    matcher = name => string.Equals(name, param, StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    Console.WriteLine($"Unknown visibility target: {target}");
                    return false;
            }

            int componentCount = SetMatchingComponentVisibility(swModel, swAssy, visible, matcher);
            int coordinateCount = SetMatchingCoordinateVisibility(swModel, visible, matcher);

            Console.WriteLine($"{componentCount + coordinateCount} items {(visible ? "shown" : "hidden")} " +
                              $"({componentCount} components, {coordinateCount} coordinate systems)");
            return true;
        }

        static int SetMatchingComponentVisibility(ModelDoc2 swModel, AssemblyDoc swAssy, bool visible, Func<string, bool> matcher)
        {
            if (swModel == null)
            {
                return 0;
            }

            object componentsObj = swAssy.GetComponents(false);
            object[] components = componentsObj as object[];
            if (components == null || matcher == null)
            {
                return 0;
            }

            int count = 0;

            foreach (object obj in components)
            {
                Component2 comp = obj as Component2;
                if (comp == null)
                {
                    continue;
                }

                string normalized = NormalizeComponentName(comp.Name2);
                if (!matcher(normalized))
                {
                    continue;
                }

                if (TrySetComponentVisibility(swModel, comp, visible))
                {
                    count++;
                }
            }

            return count;
        }

        static bool TrySetComponentVisibility(ModelDoc2 swModel, Component2 comp, bool visible)
        {
            if (swModel == null || comp == null)
            {
                return false;
            }

            try
            {
                swModel.ClearSelection2(true);
                bool selected = comp.Select4(false, null, false);
                if (!selected)
                {
                    return false;
                }

                // Requested behavior: hide/show selected components via ModelDoc2 APIs.
                if (visible)
                {
                    swModel.ShowComponent2();
                }
                else
                {
                    swModel.HideComponent2();
                }
                swModel.ClearSelection2(true);
                return true;
            }
            catch
            {
                // Fall through to compatibility paths.
            }

            int visibilityState = visible
                ? (int)swComponentVisibilityState_e.swComponentVisible
                : (int)swComponentVisibilityState_e.swComponentHidden;

            // Compatibility API path: IComponent2.SetVisibility(...)
            try
            {
                comp.SetVisibility(
                    visibilityState,
                    (int)swInConfigurationOpts_e.swThisConfiguration,
                    null);
                return true;
            }
            catch
            {
                // Fall through to property setter for broader compatibility.
            }

            // Fallback API path: IComponent2.Visible property
            try
            {
                comp.Visible = visibilityState;
                return true;
            }
            catch
            {
                return false;
            }
        }

        static int SetMatchingCoordinateVisibility(ModelDoc2 swModel, bool visible, Func<string, bool> matcher)
        {
            if (swModel == null || matcher == null)
            {
                return 0;
            }

            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                string name = feat.Name ?? string.Empty;
                if (feat.GetTypeName2() == "CoordSys" && matcher(name))
                {
                    try
                    {
                        bool updated = feat.SetSuppression2(
                            visible ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature
                                    : (int)swFeatureSuppressionAction_e.swSuppressFeature,
                            (int)swInConfigurationOpts_e.swThisConfiguration,
                            null);
                        if (updated)
                        {
                            count++;
                        }
                    }
                    catch
                    {
                        // Ignore per-feature failures.
                    }
                }

                feat = (Feature)feat.GetNextFeature();
            }

            return count;
        }

        static bool IsHardpointLikeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return name.IndexOf("_FRONT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("_REAR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("_wheel", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.StartsWith("FL_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("FR_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("RL_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("RR_", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsWheelLikeName(string name)
        {
            return !string.IsNullOrWhiteSpace(name) &&
                   name.IndexOf("_wheel", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        static bool IsFrontLikeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return name.IndexOf("_FRONT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.StartsWith("FL_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("FR_", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsRearLikeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return name.IndexOf("_REAR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.StartsWith("RL_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("RR_", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsChassisLikeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return name.IndexOf("CHAS_", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("Chassis", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        static string NormalizeComponentName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }

            string normalized = name.Trim();

            int caretIdx = normalized.IndexOf('^');
            if (caretIdx > 0)
            {
                normalized = normalized.Substring(0, caretIdx);
            }

            int angleIdx = normalized.IndexOf('<');
            if (angleIdx > 0)
            {
                normalized = normalized.Substring(0, angleIdx);
            }

            int dashIdx = normalized.LastIndexOf('-');
            if (dashIdx > 0 && dashIdx < normalized.Length - 1)
            {
                bool numericSuffix = true;
                for (int i = dashIdx + 1; i < normalized.Length; i++)
                {
                    if (!char.IsDigit(normalized[i]))
                    {
                        numericSuffix = false;
                        break;
                    }
                }

                if (numericSuffix)
                {
                    normalized = normalized.Substring(0, dashIdx);
                }
            }

            return normalized.Trim();
        }
    }
}
