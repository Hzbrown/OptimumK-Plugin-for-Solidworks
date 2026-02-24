using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public class InsertMarker
    {
        // Color mapping based on name prefixes (RGB values 0-255)
        private static readonly Dictionary<string, int[]> ColorMap = new Dictionary<string, int[]>
        {
            { "CHAS_", new[] { 255, 0, 0 } },       // Red - Chassis
            { "UPRI_", new[] { 0, 0, 255 } },       // Blue - Upright
            { "ROCK_", new[] { 0, 128, 255 } },     // Light Blue - Rocker
            { "NSMA_", new[] { 255, 192, 203 } },   // Pink - Non-Sprung Mass
            { "PUSH_", new[] { 0, 255, 0 } },       // Green - Pushrod
            { "TIER_", new[] { 255, 165, 0 } },     // Orange - Tie Rod
            { "DAMP_", new[] { 128, 0, 128 } },     // Purple - Damper
            { "ARBA_", new[] { 255, 255, 0 } },     // Yellow - ARB
            { "_FRONT", new[] { 0, 200, 200 } },    // Cyan - Front (fallback)
            { "_REAR", new[] { 200, 100, 0 } },     // Brown - Rear (fallback)
            { "wheel", new[] { 64, 64, 64 } },      // Dark Gray - Wheels
        };

        private static readonly int[] DefaultColor = new[] { 128, 128, 128 }; // Gray default

        public static bool Run(string[] args)
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return false;
            }

            string subCommand = args[1].ToLowerInvariant();

            switch (subCommand)
            {
                case "createall":
                    double radius = args.Length > 2 ? double.Parse(args[2]) : 5.0;
                    string markerPath = args.Length > 3 ? args[3] : FindMarkerPath();
                    return CreateAllMarkers(radius, markerPath);

                case "deleteall":
                    return DeleteAllMarkers();

                case "vis":
                    if (args.Length < 4)
                    {
                        Console.WriteLine("Usage: marker vis <all|front|rear|name> <show|hide> [filter]");
                        return false;
                    }
                    string target = args[2].ToLowerInvariant();
                    bool visible = args[3].ToLowerInvariant() == "show";
                    string filter = args.Length > 4 ? args[4] : null;
                    return SetMarkerVisibility(target, visible, filter);

                default:
                    Console.WriteLine($"Unknown marker command: {subCommand}");
                    PrintUsage();
                    return false;
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: SuspensionTools.exe marker <command> [args]");
            Console.WriteLine("Commands:");
            Console.WriteLine("  createall <radius_mm> [marker_path]  - Create markers at all coordinate systems");
            Console.WriteLine("  deleteall                            - Delete all marker components");
            Console.WriteLine("  vis <all|front|rear|name> <show|hide> [filter]");
        }

        private static string FindMarkerPath()
        {
            // Look for Marker.sldprt in common locations (check both cases)
            string exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] searchPaths = new[]
            {
                Path.Combine(exeDir, "Marker.SLDPRT"),
                Path.Combine(exeDir, "Marker.sldprt"),
                Path.Combine(exeDir, "..", "..", "..", "Marker.SLDPRT"),
                Path.Combine(exeDir, "..", "..", "..", "Marker.sldprt"),
                Path.Combine(exeDir, "..", "..", "..", "..", "Marker.SLDPRT"),
                Path.Combine(exeDir, "..", "..", "..", "..", "Marker.sldprt"),
                @"C:\Users\harri\OptimumK Plugin for Solidworks\Marker.SLDPRT",
                @"C:\Users\harri\OptimumK Plugin for Solidworks\Marker.sldprt",
            };

            foreach (string path in searchPaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    Console.WriteLine($"Found Marker at: {fullPath}");
                    return fullPath;
                }
            }

            throw new FileNotFoundException(
                "Marker.SLDPRT not found. Please place it in the project folder or specify path.");
        }

        private static int[] GetColorForName(string name)
        {
            string upper = name.ToUpperInvariant();
            foreach (var kvp in ColorMap)
            {
                if (upper.Contains(kvp.Key.ToUpperInvariant()))
                {
                    return kvp.Value;
                }
            }
            return DefaultColor;
        }

        private static bool CreateAllMarkers(double radiusMm, string markerPath)
        {
            SldWorks swApp;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                Console.WriteLine("Error: SolidWorks is not running");
                return false;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document");
                return false;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Console.WriteLine("Error: Active document must be an assembly");
                return false;
            }

            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            string assyTitle = swModel.GetTitle();
            string assyPath = swModel.GetPathName();

            // Check if assembly is read-only or has other issues
            Console.WriteLine($"Assembly: {assyTitle}");
            Console.WriteLine($"Assembly Path: {assyPath}");
            Console.WriteLine($"Read-only: {swModel.IsOpenedReadOnly()}");
            Console.WriteLine($"View-only: {swModel.IsOpenedViewOnly()}");

            // Resolve to absolute path
            markerPath = Path.GetFullPath(markerPath);
            
            // Verify marker file exists
            if (!File.Exists(markerPath))
            {
                Console.WriteLine($"Error: Marker file not found at {markerPath}");
                return false;
            }
            
            Console.WriteLine($"Using marker file: {markerPath}");

            // Close marker if it's already open (from previous run)
            ModelDoc2 existingMarker = swApp.GetOpenDocumentByName(markerPath) as ModelDoc2;
            if (existingMarker != null)
            {
                Console.WriteLine("Closing previously opened Marker document...");
                swApp.CloseDoc(existingMarker.GetTitle());
            }

            // IMPORTANT: Pre-load the marker document into memory
            int openErrors = 0;
            int openWarnings = 0;
            Console.WriteLine("Pre-loading Marker document...");
            ModelDoc2 markerDoc = swApp.OpenDoc6(
                markerPath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref openErrors,
                ref openWarnings
            );

            if (markerDoc == null)
            {
                Console.WriteLine($"Error: Could not open Marker document (errors={openErrors}, warnings={openWarnings})");
                return false;
            }
            Console.WriteLine($"Marker document loaded: {markerDoc.GetTitle()}");

            // Re-activate the assembly
            int activateErrors = 0;
            swApp.ActivateDoc2(assyTitle, false, ref activateErrors);
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swAssy = (AssemblyDoc)swModel;

            // Find all coordinate systems in the assembly
            List<CoordSystemInfo> coordSystems = GetAllCoordinateSystems(swModel);
            Console.WriteLine($"Found {coordSystems.Count} coordinate systems");

            if (coordSystems.Count == 0)
            {
                Console.WriteLine("No coordinate systems found to mark");
                swApp.CloseDoc(markerDoc.GetTitle());
                return true;
            }

            // Force the doc into a "safe to modify" state before insertion
            swApp.CommandInProgress = true;
            swModel.ClearSelection2(true);
            
            // Ensure we're editing the assembly, not a component
            swAssy.EditAssembly();
            
            // Force rebuild to ensure clean state
            swModel.ForceRebuild3(false);

            // Store inserted components for post-processing
            List<Component2> insertedComponents = new List<Component2>();
            List<string> componentNames = new List<string>();
            
            int created = 0;
            int skipped = 0;

            foreach (var csInfo in coordSystems)
            {
                string markerName = csInfo.Name + "_Marker";

                // Check if marker already exists
                if (ComponentExists(swAssy, markerName))
                {
                    Console.WriteLine($"Skipping {csInfo.Name} - marker already exists");
                    skipped++;
                    continue;
                }

                try
                {
                    Console.WriteLine($"Creating marker for: {csInfo.Name} at ({csInfo.X*1000:F1}, {csInfo.Y*1000:F1}, {csInfo.Z*1000:F1}) mm");

                    // Re-activate the assembly before each insert
                    int errors = 0;
                    swApp.ActivateDoc2(assyTitle, false, ref errors);
                    
                    // Refresh swModel and swAssy references
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    swAssy = (AssemblyDoc)swModel;
                    
                    // Ensure we're not editing a component
                    swAssy.EditAssembly();
                    swModel.ClearSelection2(true);

                    // Try AddComponent4 instead of AddComponent5
                    // AddComponent4 signature: (PathName, ConfigName, X, Y, Z)
                    Component2 newComp = (Component2)swAssy.AddComponent4(
                        markerPath,
                        "",          // configuration name (blank = default)
                        csInfo.X,
                        csInfo.Y,
                        csInfo.Z
                    );

                    if (newComp == null)
                    {
                        // Try alternative: AddComponent5 with explicit options
                        Console.WriteLine($"  AddComponent4 failed, trying AddComponent5...");
                        newComp = swAssy.AddComponent5(
                            markerPath,
                            (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                            "",          // configuration name
                            false,       // suppress
                            "",          // component name
                            csInfo.X,
                            csInfo.Y,
                            csInfo.Z
                        );
                    }

                    if (newComp == null)
                    {
                        // Last resort: try AddComponent2
                        Console.WriteLine($"  AddComponent5 failed, trying AddComponent2...");
                        newComp = (Component2)swAssy.AddComponent2(
                            markerPath,
                            csInfo.X,
                            csInfo.Y,
                            csInfo.Z
                        );
                    }

                    if (newComp == null)
                    {
                        Console.WriteLine($"  All AddComponent methods failed for {csInfo.Name}");
                        Console.WriteLine($"  Check: Is the assembly saved? Is it read-only? Is a sketch or feature being edited?");
                        continue;
                    }

                    Console.WriteLine($"  Component inserted: {newComp.Name2}");
                    
                    // Store for later processing
                    insertedComponents.Add(newComp);
                    componentNames.Add(markerName);

                    // Apply color immediately
                    ApplyColorToComponent(newComp, csInfo.Name);

                    created++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error creating marker for {csInfo.Name}: {ex.Message}");
                    Console.WriteLine($"  Stack trace: {ex.StackTrace}");
                }
            }

            // Restore command state
            swApp.CommandInProgress = false;

            // Post-process: make virtual and rename
            Console.WriteLine($"\nProcessing {insertedComponents.Count} components...");
            for (int i = 0; i < insertedComponents.Count; i++)
            {
                Component2 comp = insertedComponents[i];
                string newName = componentNames[i];
                
                try
                {
                    // Make virtual
                    bool madeVirtual = comp.MakeVirtual2(false);
                    if (madeVirtual)
                    {
                        Console.WriteLine($"  Made virtual: {comp.Name2}");
                    }
                    
                    // Rename
                    comp.Name2 = newName;
                    Console.WriteLine($"  Renamed to: {newName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Warning processing {comp.Name2}: {ex.Message}");
                }
            }

            // Close the marker document
            Console.WriteLine("Closing Marker document...");
            swApp.ActivateDoc2(assyTitle, false, ref activateErrors);
            swApp.CloseDoc(markerDoc.GetTitle());

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.EditRebuild3();
            Console.WriteLine($"\nCreated {created} markers, skipped {skipped} existing");
            return true;
        }

        private static List<CoordSystemInfo> GetAllCoordinateSystems(ModelDoc2 swModel)
        {
            var coordSystems = new List<CoordSystemInfo>();

            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "CoordSys")
                {
                    try
                    {
                        CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)feat.GetDefinition();
                        if (csData != null)
                        {
                            // Access selection needed to get transform
                            bool accessOk = csData.AccessSelections(swModel, null);
                            
                            MathTransform transform = csData.Transform;
                            if (transform != null)
                            {
                                double[] td = (double[])transform.ArrayData;
                                // ArrayData: [0-8] rotation matrix, [9-11] translation (meters)
                                coordSystems.Add(new CoordSystemInfo
                                {
                                    Name = feat.Name,
                                    Feature = feat,
                                    X = td[9],
                                    Y = td[10],
                                    Z = td[11]
                                });
                            }
                            
                            csData.ReleaseSelectionAccess();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not get data for {feat.Name}: {ex.Message}");
                    }
                }
                feat = (Feature)feat.GetNextFeature();
            }

            return coordSystems;
        }

        private static bool ComponentExists(AssemblyDoc swAssy, string componentName)
        {
            object[] components = (object[])swAssy.GetComponents(false);
            if (components == null) return false;

            foreach (object obj in components)
            {
                Component2 comp = (Component2)obj;
                if (comp.Name2.Contains(componentName))
                    return true;
            }
            return false;
        }

        private static void ApplyColorToComponent(Component2 comp, string csName)
        {
            try
            {
                int[] rgb = GetColorForName(csName);
                double[] matProps = new double[9];
                matProps[0] = rgb[0] / 255.0;  // R
                matProps[1] = rgb[1] / 255.0;  // G
                matProps[2] = rgb[2] / 255.0;  // B
                matProps[3] = 1.0;             // Ambient
                matProps[4] = 1.0;             // Diffuse
                matProps[5] = 0.2;             // Specular
                matProps[6] = 0.3;             // Shininess
                matProps[7] = 0.0;             // Transparency
                matProps[8] = 0.0;             // Emission

                comp.MaterialPropertyValues = matProps;
                Console.WriteLine($"  Applied color RGB({rgb[0]},{rgb[1]},{rgb[2]})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Warning: Could not set color: {ex.Message}");
            }
        }

        private static void AddMateToCoordSystem(SldWorks swApp, ModelDoc2 swModel, AssemblyDoc swAssy,
            Component2 markerComp, CoordSystemInfo csInfo)
        {
            try
            {
                swModel.ClearSelection2(true);

                // Get the origin point of the marker component
                ModelDoc2 markerDoc = (ModelDoc2)markerComp.GetModelDoc2();
                if (markerDoc == null)
                {
                    Console.WriteLine("  Warning: Could not access marker model for mating");
                    return;
                }

                // Select the coordinate system's origin entity
                // The coordinate system feature has an OriginEntity we can use for mating
                CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)csInfo.Feature.GetDefinition();
                object originEntity = csData.OriginEntity;

                if (originEntity == null)
                {
                    Console.WriteLine("  Warning: Could not get origin entity from coordinate system");
                    // Fall back to just positioning (already done during insert)
                    return;
                }

                // Select the coordinate system origin entity
                SelectionMgr selMgr = (SelectionMgr)swModel.SelectionManager;
                
                // Select coord system origin (mark 1)
                SelectData selData1 = (SelectData)selMgr.CreateSelectData();
                selData1.Mark = 1;
                
                bool sel1 = swModel.Extension.SelectByID2(
                    csInfo.Name, "COORDSYS", 0, 0, 0, false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!sel1)
                {
                    Console.WriteLine($"  Warning: Could not select coordinate system {csInfo.Name}");
                    return;
                }

                // Select marker component origin (mark 2) 
                string markerOriginName = markerComp.Name2 + "@" + swModel.GetTitle().Replace(".SLDASM", "") + "/Origin";
                bool sel2 = swModel.Extension.SelectByID2(
                    "Point1@Origin@" + markerComp.Name2 + "@" + swModel.GetTitle(),
                    "EXTSKETCHPOINT", 0, 0, 0, true, 2, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!sel2)
                {
                    // Try alternate selection for origin
                    sel2 = swModel.Extension.SelectByID2(
                        "Origin@" + markerComp.Name2 + "@" + swModel.GetTitle(),
                        "ORIGINFOLDER", 0, 0, 0, true, 2, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                if (sel1)
                {
                    // Add coincident mate
                    int mateError = 0;
                    Mate2 mate = swAssy.AddMate5(
                        (int)swMateType_e.swMateCOINCIDENT,
                        (int)swMateAlign_e.swMateAlignALIGNED,
                        false,  // Flip
                        0, 0, 0, 0, 0, 0, 0, 0,  // Distances and angles (11 doubles total)
                        false,  // ForPositioningOnly
                        false,  // LockRotation
                        (int)swMateWidthOptions_e.swMateWidth_Centered,  // WidthMateOption
                        out mateError);

                    if (mate != null && mateError == 0)
                    {
                        Console.WriteLine($"  Added mate to {csInfo.Name}");
                    }
                    else
                    {
                        Console.WriteLine($"  Note: Component positioned at coordinate system (mate skipped, error: {mateError})");
                    }
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Note: Mate creation skipped: {ex.Message}");
                swModel.ClearSelection2(true);
            }
        }

        private static bool DeleteAllMarkers()
        {
            SldWorks swApp;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                Console.WriteLine("Error: SolidWorks is not running");
                return false;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document");
                return false;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Console.WriteLine("Error: Active document must be an assembly");
                return false;
            }

            AssemblyDoc swAssy = (AssemblyDoc)swModel;

            // Find all marker components
            List<Component2> markersToDelete = new List<Component2>();
            object[] components = (object[])swAssy.GetComponents(false);

            if (components != null)
            {
                foreach (object obj in components)
                {
                    Component2 comp = (Component2)obj;
                    if (comp.Name2.Contains("_Marker"))
                    {
                        markersToDelete.Add(comp);
                    }
                }
            }

            Console.WriteLine($"Found {markersToDelete.Count} markers to delete");

            int deleted = 0;
            foreach (Component2 comp in markersToDelete)
            {
                try
                {
                    swModel.ClearSelection2(true);
                    comp.Select4(false, null, false);
                    swModel.EditDelete();
                    deleted++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not delete {comp.Name2}: {ex.Message}");
                }
            }

            swModel.EditRebuild3();
            Console.WriteLine($"Deleted {deleted} markers");
            return true;
        }

        private static bool SetMarkerVisibility(string target, bool visible, string filter)
        {
            SldWorks swApp;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                Console.WriteLine("Error: SolidWorks is not running");
                return false;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document");
                return false;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Console.WriteLine("Error: Active document must be an assembly");
                return false;
            }

            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            object[] components = (object[])swAssy.GetComponents(false);

            if (components == null)
            {
                Console.WriteLine("No components in assembly");
                return true;
            }

            int count = 0;
            foreach (object obj in components)
            {
                Component2 comp = (Component2)obj;
                string compName = comp.Name2;

                if (!compName.Contains("_Marker")) continue;

                bool shouldModify = false;
                string baseName = compName.Replace("_Marker", "").Split('-')[0];

                switch (target)
                {
                    case "all":
                        shouldModify = true;
                        break;
                    case "front":
                        shouldModify = baseName.ToUpperInvariant().Contains("FRONT");
                        break;
                    case "rear":
                        shouldModify = baseName.ToUpperInvariant().Contains("REAR");
                        break;
                    case "name":
                        shouldModify = filter != null &&
                            baseName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
                        break;
                }

                if (shouldModify)
                {
                    try
                    {
                        int suppState = visible
                            ? (int)swComponentSuppressionState_e.swComponentResolved
                            : (int)swComponentSuppressionState_e.swComponentSuppressed;

                        comp.SetSuppression2(suppState);
                        count++;
                    }
                    catch { }
                }
            }

            Console.WriteLine($"{count} markers {(visible ? "shown" : "hidden")}");
            return true;
        }

        private class CoordSystemInfo
        {
            public string Name { get; set; }
            public Feature Feature { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
        }
    }
}
