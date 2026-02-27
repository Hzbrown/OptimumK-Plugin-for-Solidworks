using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    // State enum for progress reporting
    public enum MarkerCreationState
    {
        Initializing,
        LoadingMarker,
        ScanningCoordSystems,
        InsertingComponents,
        MatingMarkers,
        PostProcessing,
        Cleanup,
        Complete
    }

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

        private static readonly int[] TieRodColor = new[] { 255, 165, 0 }; // Orange (Tie Rod)
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

            // Tie rod tokens should take precedence over CHAS_/UPRI_ mixed names
            // like CHAS_TiePnt_L and UPRI_TiePnt_R.
            if (upper.Contains("TIER_") ||
                upper.Contains("TIEROD") ||
                upper.Contains("TIE_ROD") ||
                upper.Contains("TIE ROD") ||
                upper.Contains("TIEPNT") ||
                upper.Contains("TIE_PNT"))
            {
                return TieRodColor;
            }

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
            // Report initial state
            Console.WriteLine("STATE:Initializing");
            
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

            Console.WriteLine($"Assembly: {assyTitle}");

            // Resolve to absolute path
            markerPath = Path.GetFullPath(markerPath);
            
            if (!File.Exists(markerPath))
            {
                Console.WriteLine($"Error: Marker file not found at {markerPath}");
                return false;
            }

            // State: Loading marker document
            Console.WriteLine("STATE:LoadingMarker");
            Console.WriteLine($"Using marker file: {markerPath}");

            ModelDoc2 existingMarker = swApp.GetOpenDocumentByName(markerPath) as ModelDoc2;
            if (existingMarker != null)
            {
                Console.WriteLine("Closing previously opened Marker document...");
                swApp.CloseDoc(existingMarker.GetTitle());
            }

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

            int activateErrors = 0;
            swApp.ActivateDoc2(assyTitle, false, ref activateErrors);
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swAssy = (AssemblyDoc)swModel;

            // State: Scanning coordinate systems
            Console.WriteLine("STATE:ScanningCoordSystems");

            int totalCoordSystems = CountCoordinateSystems(swModel);
            List<CoordSystemInfo> coordSystems = GetAllCoordinateSystems(swModel);

            // Filter out existing markers
            var coordSystemsToProcess = new List<CoordSystemInfo>();
            int existingCount = 0;
            foreach (var csInfo in coordSystems)
            {
                            string markerName = csInfo.Name;
                if (ComponentExists(swAssy, markerName))
                    existingCount++;
                else
                    coordSystemsToProcess.Add(csInfo);
            }

            // Precompute total steps: scan + insert + mate + post (all per target)
            int scanTasks = coordSystems.Count;
            int insertTasks = coordSystemsToProcess.Count;
            int mateTasks = coordSystemsToProcess.Count;
            int postTasks = coordSystemsToProcess.Count;
            int totalTasks = scanTasks + insertTasks + mateTasks + postTasks;

            Console.WriteLine($"TOTAL:{totalTasks}");
            Console.WriteLine($"Found {coordSystems.Count} coordinate systems ({existingCount} already have markers, {coordSystemsToProcess.Count} to create)");

            if (coordSystemsToProcess.Count == 0)
            {
                Console.WriteLine("No new coordinate systems to mark");
                swApp.CloseDoc(markerDoc.GetTitle());
                Console.WriteLine("STATE:Complete");
                return true;
            }

            // Progress starts at 0; drive PROGRESS through every phase
            int progressCount = 0;

            // Emit scan progress (already have data, just advance bar)
            foreach (var _ in coordSystems)
            {
                progressCount++;
                Console.WriteLine($"PROGRESS:{progressCount}");
            }

            // Do NOT force CommandInProgress=true here. If the runner process is aborted,
            // SolidWorks can remain UI-locked (feature tree not clickable).
            // We rely on normal API calls without entering command-in-progress mode.
            swModel.ClearSelection2(true);
            swAssy.EditAssembly();
            swModel.ForceRebuild3(false);

            var workItems = new List<MarkerWorkItem>();
            Console.WriteLine("STATE:InsertingComponents");

            foreach (var csInfo in coordSystemsToProcess)
            {
                var item = new MarkerWorkItem { CsInfo = csInfo, NewName = csInfo.Name + "_Marker" };
                try
                {
                    Console.WriteLine($"Creating marker for: {csInfo.Name} at ({csInfo.X*1000:F1}, {csInfo.Y*1000:F1}, {csInfo.Z*1000:F1}) mm");

                    int errors = 0;
                    swApp.ActivateDoc2(assyTitle, false, ref errors);
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    swAssy = (AssemblyDoc)swModel;
                    swAssy.EditAssembly();
                    swModel.ClearSelection2(true);

                    Component2 newComp = (Component2)swAssy.AddComponent4(markerPath, "", csInfo.X, csInfo.Y, csInfo.Z);
                    if (newComp == null)
                        newComp = swAssy.AddComponent5(markerPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", csInfo.X, csInfo.Y, csInfo.Z);
                    if (newComp == null)
                        newComp = (Component2)swAssy.AddComponent2(markerPath, csInfo.X, csInfo.Y, csInfo.Z);

                    if (newComp != null)
                    {
                        item.Comp = newComp;
                        ApplyColorToComponent(newComp, csInfo.Name);
                    }
                    else
                    {
                        Console.WriteLine($"  All AddComponent methods failed for {csInfo.Name}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error creating marker for {csInfo.Name}: {ex.Message}");
                }
                workItems.Add(item);
                progressCount++;
                Console.WriteLine($"PROGRESS:{progressCount}");
            }

            Console.WriteLine("STATE:MatingMarkers");
            foreach (var item in workItems)
            {
                try
                {
                    if (item.Comp != null)
                        AddMateToCoordSystem(swApp, swModel, swAssy, item.Comp, item.CsInfo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Note: Mate skipped for {item.CsInfo.Name}: {ex.Message}");
                }
                progressCount++;
                Console.WriteLine($"PROGRESS:{progressCount}");
            }

            // Ensure SolidWorks command mode is released (best effort).
            try { swApp.CommandInProgress = false; } catch { }

            Console.WriteLine("STATE:PostProcessing");
            foreach (var item in workItems)
            {
                try
                {
                    if (item.Comp != null)
                    {
                        bool madeVirtual = item.Comp.MakeVirtual2(false);
                        if (madeVirtual) Console.WriteLine($"  Made virtual: {item.Comp.Name2}");
                        item.Comp.Name2 = item.NewName;
                        Console.WriteLine($"  Renamed to: {item.NewName}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Warning processing {item.NewName}: {ex.Message}");
                }
                progressCount++;
                Console.WriteLine($"PROGRESS:{progressCount}");
            }

            Console.WriteLine("STATE:Cleanup");
            Console.WriteLine("Closing Marker document...");
            swApp.ActivateDoc2(assyTitle, false, ref activateErrors);
            swApp.CloseDoc(markerDoc.GetTitle());

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.EditRebuild3();
            
            Console.WriteLine("STATE:Complete");
            Console.WriteLine($"Created {workItems.FindAll(w => w.Comp != null).Count} markers, skipped {existingCount} existing");
            return true;
        }

        private static int CountCoordinateSystems(ModelDoc2 swModel)
        {
            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "CoordSys")
                {
                    count++;
                }
                feat = (Feature)feat.GetNextFeature();
            }
            return count;
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
                        if (csData != null && csData.AccessSelections(swModel, null))
                        {
                            MathTransform transform = csData.Transform;
                            if (transform != null)
                            {
                                double[] td = (double[])transform.ArrayData;
                                coordSystems.Add(new CoordSystemInfo { Name = feat.Name, Feature = feat, X = td[9], Y = td[10], Z = td[11] });
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
                    Mate2 mate = swAssy.AddMate5(
                        (int)swMateType_e.swMateCOINCIDENT,
                        (int)swMateAlign_e.swMateAlignALIGNED,
                        false,  // Flip
                        0, 0, 0, 0, 0, 0, 0, 0,  // Distances and angles (11 doubles total)
                        false,  // ForPositioningOnly
                        false,  // LockRotation
                        (int)swMateWidthOptions_e.swMateWidth_Centered,  // WidthMateOption
                        out int mateError);

                    if (mate != null && mateError == 0)
                    {
                        Feature mateFeat = (Feature)swModel.FeatureByPositionReverse(0);
                        if (mateFeat != null)
                        {
                            bool axisAligned = TryEnableCoincidentMateAxisAlignment(swModel, mateFeat);
                            if (axisAligned)
                            {
                                Console.WriteLine($"  Enabled coincident mate axis alignment for {csInfo.Name}");
                            }
                        }

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

        private static bool TryEnableCoincidentMateAxisAlignment(ModelDoc2 swModel, Feature mateFeature)
        {
            if (swModel == null || mateFeature == null)
            {
                return false;
            }

            // Macro-style first to mirror VBA behavior directly.
            if (TryEnableCoincidentMateAxisAlignmentByMacro(swModel, mateFeature))
            {
                return true;
            }

            // Fallback: definition edit when macro-style call is unavailable.
            return TryEnableCoincidentMateAxisAlignmentByDefinition(swModel, mateFeature);
        }

        private static bool TryEnableCoincidentMateAxisAlignmentByDefinition(ModelDoc2 swModel, Feature mateFeature)
        {
            try
            {
                object mateDefinition = mateFeature.GetDefinition();
                if (mateDefinition == null)
                {
                    return false;
                }

                Type definitionType = mateDefinition.GetType();
                var accessSelections = definitionType.GetMethod("AccessSelections");
                bool selectionAccessOpen = false;

                if (accessSelections != null)
                {
                    object accessResult = accessSelections.Invoke(mateDefinition, new object[] { swModel, null });
                    if (accessResult is bool && !(bool)accessResult)
                    {
                        return false;
                    }
                    selectionAccessOpen = true;
                }

                bool changed = false;
                changed |= TrySetIntProperty(
                    mateDefinition,
                    "MateAlignment",
                    (int)swMateAlign_e.swMateAlignALIGNED);
                changed |= TrySetBooleanProperty(mateDefinition, "AlignAxes", true);
                changed |= TrySetBooleanProperty(mateDefinition, "AlignAxis", true);

                if (!changed)
                {
                    if (selectionAccessOpen)
                    {
                        var releaseSelections = definitionType.GetMethod("ReleaseSelectionAccess");
                        releaseSelections?.Invoke(mateDefinition, null);
                    }
                    return false;
                }

                bool modified = mateFeature.ModifyDefinition(mateDefinition, swModel, null);

                if (selectionAccessOpen)
                {
                    var releaseSelections = definitionType.GetMethod("ReleaseSelectionAccess");
                    releaseSelections?.Invoke(mateDefinition, null);
                }

                return modified;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryEnableCoincidentMateAxisAlignmentByMacro(ModelDoc2 swModel, Feature mateFeature)
        {
            try
            {
                string mateName = mateFeature.Name;
                if (string.IsNullOrWhiteSpace(mateName))
                {
                    return false;
                }

                swModel.ClearSelection2(true);
                bool selected = swModel.Extension.SelectByID2(
                    mateName,
                    "MATE",
                    0, 0, 0,
                    false,
                    0,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!selected)
                {
                    swModel.ClearSelection2(true);
                    return false;
                }

                object modelObject = swModel;
                Type modelType = modelObject.GetType();
                System.Reflection.MethodInfo[] methods = modelType.GetMethods();

                for (int i = 0; i < methods.Length; i++)
                {
                    System.Reflection.MethodInfo method = methods[i];
                    if (!string.Equals(method.Name, "AddMate5", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    System.Reflection.ParameterInfo[] parameters = method.GetParameters();
                    if (!TryBuildMacroAddMate5Arguments(parameters, out object[] args, out int errorIndex))
                    {
                        continue;
                    }

                    object invokeResult;
                    try
                    {
                        invokeResult = method.Invoke(modelObject, args);
                    }
                    catch
                    {
                        continue;
                    }

                    int errorStatus = 0;
                    if (errorIndex >= 0 && errorIndex < args.Length && args[errorIndex] is int)
                    {
                        errorStatus = (int)args[errorIndex];
                    }

                    if (invokeResult != null && errorStatus == 0)
                    {
                        swModel.ClearSelection2(true);
                        return true;
                    }
                }

                swModel.ClearSelection2(true);
                return false;
            }
            catch
            {
                try { swModel.ClearSelection2(true); } catch { }
                return false;
            }
        }

        private static bool TryBuildMacroAddMate5Arguments(
            System.Reflection.ParameterInfo[] parameters,
            out object[] args,
            out int errorIndex)
        {
            args = null;
            errorIndex = -1;

            if (parameters == null || parameters.Length == 0)
            {
                return false;
            }

            int[] intValues = new[]
            {
                (int)swMateType_e.swMateCOORDINATE,
                (int)swMateAlign_e.swMateAlignALIGNED,
                0
            };

            bool[] boolValues = new[]
            {
                false,
                false,
                false,
                false
            };

            double[] doubleValues = new[]
            {
                0.0,
                0.001, 0.001, 0.001, 0.001,
                0.5235987755983, 0.5235987755983, 0.5235987755983
            };

            int intCursor = 0;
            int boolCursor = 0;
            int doubleCursor = 0;

            args = new object[parameters.Length];

            for (int i = 0; i < parameters.Length; i++)
            {
                Type paramType = parameters[i].ParameterType;
                Type elementType = paramType.IsByRef ? paramType.GetElementType() : paramType;

                if (parameters[i].IsOut || (paramType.IsByRef && elementType == typeof(int)))
                {
                    args[i] = 0;
                    errorIndex = i;
                    continue;
                }

                if (elementType == typeof(int))
                {
                    args[i] = intCursor < intValues.Length ? intValues[intCursor++] : 0;
                    continue;
                }

                if (elementType == typeof(bool))
                {
                    args[i] = boolCursor < boolValues.Length ? boolValues[boolCursor++] : false;
                    continue;
                }

                if (elementType == typeof(double))
                {
                    args[i] = doubleCursor < doubleValues.Length ? doubleValues[doubleCursor++] : 0.0;
                    continue;
                }

                if (elementType == typeof(object))
                {
                    args[i] = null;
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool TrySetBooleanProperty(object target, string propertyName, bool value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var propertyInfo = target.GetType().GetProperty(
                    propertyName,
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.IgnoreCase);

                if (propertyInfo == null || !propertyInfo.CanWrite || propertyInfo.PropertyType != typeof(bool))
                {
                    return false;
                }

                propertyInfo.SetValue(target, value, null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySetIntProperty(object target, string propertyName, int value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var propertyInfo = target.GetType().GetProperty(
                    propertyName,
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.IgnoreCase);

                if (propertyInfo == null || !propertyInfo.CanWrite)
                {
                    return false;
                }

                Type propertyType = propertyInfo.PropertyType;
                if (propertyType == typeof(int))
                {
                    propertyInfo.SetValue(target, value, null);
                    return true;
                }

                if (propertyType.IsEnum)
                {
                    object enumValue = Enum.ToObject(propertyType, value);
                    propertyInfo.SetValue(target, enumValue, null);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
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

        private class MarkerWorkItem
        {
            public CoordSystemInfo CsInfo { get; set; }
            public Component2 Comp { get; set; }
            public string NewName { get; set; }
        }
    }
}
