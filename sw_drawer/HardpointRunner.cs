using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public class HardpointRunner
    {
        // Color mapping based on actual JSON name prefixes (RGB values 0-255)
        private static readonly Dictionary<string, int[]> ColorMap = new Dictionary<string, int[]>
        {
            { "CHAS_", new[] { 255, 0, 0 } },       // Red - Chassis
            { "wheel", new[] { 0, 255, 0 } },       // Green - Wheels
            { "UPRI_", new[] { 0, 0, 255 } },       // Blue - Upright
        };

        private static readonly int[] DefaultColor = new[] { 128, 0, 128 }; // Purple for other components

        private static void ReportState(HardpointState state)
        {
            Console.WriteLine($"STATE:{state}");
        }

        private static void ReportProgress(int count)
        {
            Console.WriteLine($"PROGRESS:{count}");
        }

        public static bool RunAdd(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: hardpoints add <jsonPath> <markerPartPath>");
                return false;
            }

            string jsonPath = args[2];
            string markerPartPath = args[3];

            ReportState(HardpointState.Initializing);
            
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

            // Validate marker file exists
            if (!File.Exists(markerPartPath))
            {
                Console.WriteLine($"Error: Marker part not found at {markerPartPath}");
                return false;
            }

            try
            {
                // Load JSON data from both Front and Rear suspension files
                ReportState(HardpointState.LoadingJson);
                
                string frontJsonPath = Path.Combine(Path.GetDirectoryName(jsonPath), "Front_Suspension.json");
                string rearJsonPath = Path.Combine(Path.GetDirectoryName(jsonPath), "Rear_Suspension.json");
                double rearReferenceOffsetMm = LoadRearReferenceOffsetMm(Path.GetDirectoryName(jsonPath));
                
                var hardpoints = new List<HardpointInfo>();
                
                // Load front suspension
                if (File.Exists(frontJsonPath))
                {
                    var frontData = LoadJsonData(frontJsonPath);
                    ExtractHardpointsWithSuffix(frontData, "_FRONT", hardpoints, 0.0);
                    // Also extract wheels from front suspension
                    ExtractWheelsFromJson(frontData, "_FRONT", hardpoints, false, 0.0);
                    Console.WriteLine($"Loaded {hardpoints.Count} hardpoints from Front_Suspension.json");
                }
                
                // Load rear suspension
                int frontCount = hardpoints.Count;
                if (File.Exists(rearJsonPath))
                {
                    var rearData = LoadJsonData(rearJsonPath);
                    ExtractHardpointsWithSuffix(rearData, "_REAR", hardpoints, rearReferenceOffsetMm);
                    // Also extract wheels from rear suspension
                    ExtractWheelsFromJson(rearData, "_REAR", hardpoints, true, rearReferenceOffsetMm);
                    Console.WriteLine($"Loaded {hardpoints.Count - frontCount} hardpoints from Rear_Suspension.json");
                }
                
                if (hardpoints.Count == 0)
                {
                    Console.WriteLine("Error: No hardpoints found in JSON files");
                    return false;
                }
                
                // Precompute total steps (insert + make virtual + rename + float + rename CS + color + folder + rebuild)
                int totalSteps = hardpoints.Count * 6 + 2;
                Console.WriteLine($"TOTAL:{totalSteps}");
                int progressCount = 0;

                // Load marker part
                ReportState(HardpointState.LoadingMarkerPart);
                ModelDoc2 markerDoc = LoadMarkerPart(swApp, markerPartPath);
                if (markerDoc == null)
                {
                    return false;
                }

                int activateErrors = 0;
                swApp.ActivateDoc2(swModel.GetTitle(), false, ref activateErrors);
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swAssy = (AssemblyDoc)swModel;

                // Start insertion process
                swModel.ClearSelection2(true);
                swAssy.EditAssembly();

                var insertedParts = new List<VirtualPartInfo>();
                
                // Insert all components and immediately process each one
                ReportState(HardpointState.InsertingBodies);
                foreach (var hardpoint in hardpoints)
                {
                    // Step 1: Insert component
                    var partInfo = InsertVirtualMarkerPart(swAssy, markerDoc, hardpoint);
                    if (partInfo != null)
                    {
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        // Step 2: Make virtual immediately
                        MakeComponentVirtual(partInfo);
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        // Step 3: Rename immediately
                        RenameVirtualPart(swAssy, partInfo);
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        // Step 4: Float the component in the assembly (unfix)
                        FloatComponent(swAssy, swModel, partInfo);
                        progressCount++;
                        ReportProgress(progressCount);

                        // Step 5: Rename internal coordinate system using EditPart2.
                        // For wheels, also apply camber/toe orientation to the internal CS.
                        if (IsWheelHardpoint(partInfo.BaseName))
                        {
                            RenameAndOrientWheelCoordinateSystem(swApp, swAssy, swModel, partInfo);
                        }
                        else
                        {
                            RenameInternalCoordinateSystem(swApp, swAssy, swModel, partInfo);
                        }
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        // Step 6: Apply color
                        ApplyColorToVirtualPart(swAssy, partInfo);
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        insertedParts.Add(partInfo);
                    }
                }

                // Step 6: Create folder and move all parts at once
                ReportState(HardpointState.CreatingHardpointsFolder);
                if (insertedParts.Count > 0)
                {
                    CreateContainingFolderFromComponents(swModel, "Hardpoints", insertedParts);
                }
                progressCount++;
                ReportProgress(progressCount);

                // Step 7: Rebuild
                swModel.EditRebuild3();
                progressCount++;
                ReportProgress(progressCount);

                ReportState(HardpointState.Complete);
                Console.WriteLine($"Successfully inserted {insertedParts.Count} virtual marker parts");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        private static Dictionary<string, object> LoadJsonData(string jsonPath)
        {
            string jsonContent = File.ReadAllText(jsonPath);
            // Simple JSON parsing for .NET Framework 4.8 compatibility
            return ParseJson(jsonContent);
        }

        private static List<HardpointInfo> ExtractHardpoints(Dictionary<string, object> jsonData)
        {
            var hardpoints = new List<HardpointInfo>();

            // Front_Suspension.json structure: top level has "Double A-Arm", "Push Pull", "Wheels"
            // Rear_Suspension.json structure: top level has "Double A-Arm", "Push Pull", "Wheels"
            // Determine if Front or Rear by checking which file was loaded
            
            // Extract from all sections except Wheels
            foreach (var section in jsonData)
            {
                if (section.Key == "Wheels")
                {
                    // Handle wheels separately
                    ExtractWheels(section.Value as Dictionary<string, object>, hardpoints);
                }
                else
                {
                    // This is a suspension section (e.g., "Double A-Arm", "Push Pull")
                    // Since we're reading from Front_Suspension.json, use _FRONT suffix
                    Dictionary<string, object> sectionObj = section.Value as Dictionary<string, object>;
                    if (sectionObj != null)
                    {
                        foreach (var pointProperty in sectionObj)
                        {
                            string pointName = pointProperty.Key;
                            List<object> coords = pointProperty.Value as List<object>;
                            if (coords != null && coords.Count >= 3)
                            {
                                hardpoints.Add(new HardpointInfo
                                {
                                    BaseName = pointName,
                                    Suffix = "_FRONT",
                                    X = Convert.ToDouble(coords[0]),
                                    Y = Convert.ToDouble(coords[1]),
                                    Z = Convert.ToDouble(coords[2])
                                });
                            }
                        }
                    }
                }
            }

            return hardpoints;
        }

        private static void ExtractHardpointsFromSection(Dictionary<string, object> section, string suffix, List<HardpointInfo> hardpoints)
        {
            if (section == null) return;

            foreach (var property in section)
            {
                string sectionName = property.Key;
                if (sectionName == "Wheels") continue;

                Dictionary<string, object> sectionObj = property.Value as Dictionary<string, object>;
                if (sectionObj != null)
                {
                    foreach (var pointProperty in sectionObj)
                    {
                        string pointName = pointProperty.Key;
                        List<object> coords = pointProperty.Value as List<object>;
                        if (coords != null && coords.Count >= 3)
                        {
                            hardpoints.Add(new HardpointInfo
                            {
                                BaseName = pointName,
                                Suffix = suffix,
                                X = Convert.ToDouble(coords[0]),
                                Y = Convert.ToDouble(coords[1]),
                                Z = Convert.ToDouble(coords[2])
                            });
                        }
                    }
                }
            }
        }

        private static void ExtractWheels(Dictionary<string, object> wheels, List<HardpointInfo> hardpoints)
        {
            if (wheels == null) return;

            if (wheels.TryGetValue("Half Track", out object halfTrackObj) &&
                wheels.TryGetValue("Tire Diameter", out object tireDiameterObj))
            {
                Dictionary<string, object> halfTrack = halfTrackObj as Dictionary<string, object>;
                Dictionary<string, object> tireDiameter = tireDiameterObj as Dictionary<string, object>;
                
                if (halfTrack != null && tireDiameter != null &&
                    halfTrack.TryGetValue("left", out object halfTrackValue) &&
                    tireDiameter.TryGetValue("left", out object tireDiameterValue))
                {
                    double halfTrackDouble = Convert.ToDouble(halfTrackValue);
                    double tireRadius = Convert.ToDouble(tireDiameterValue) / 2.0;

                    // Front wheels
                    hardpoints.Add(new HardpointInfo
                    {
                        BaseName = "FL_wheel",
                        Suffix = "_FRONT",
                        X = 0, // Will be set by reference distance
                        Y = halfTrackDouble,
                        Z = tireRadius
                    });

                    hardpoints.Add(new HardpointInfo
                    {
                        BaseName = "FR_wheel", 
                        Suffix = "_FRONT",
                        X = 0,
                        Y = -halfTrackDouble,
                        Z = tireRadius
                    });

                    // Rear wheels (will be adjusted by reference distance in pose update)
                    hardpoints.Add(new HardpointInfo
                    {
                        BaseName = "RL_wheel",
                        Suffix = "_REAR",
                        X = 0,
                        Y = halfTrackDouble,
                        Z = tireRadius
                    });

                    hardpoints.Add(new HardpointInfo
                    {
                        BaseName = "RR_wheel",
                        Suffix = "_REAR", 
                        X = 0,
                        Y = -halfTrackDouble,
                        Z = tireRadius
                    });
                }
            }
        }

        private static void ExtractWheelsFromJson(Dictionary<string, object> jsonData, string suffix, List<HardpointInfo> hardpoints)
        {
            bool isRear = string.Equals(suffix, "_REAR", StringComparison.OrdinalIgnoreCase);
            ExtractWheelsFromJson(jsonData, suffix, hardpoints, isRear, 0.0);
        }

        private static void ExtractWheelsFromJson(
            Dictionary<string, object> jsonData,
            string suffix,
            List<HardpointInfo> hardpoints,
            bool isRear,
            double referenceDistanceMm)
        {
            if (!(jsonData.TryGetValue("Wheels", out object wheelsObj) && wheelsObj is Dictionary<string, object> wheelsData))
            {
                return;
            }

            try
            {
                // Mirror draw_suspension.py InsertWheel logic.
                double halfTrack      = GetWheelValue(wheelsData, "Half Track", "left");
                double tireDiameter   = GetWheelValue(wheelsData, "Tire Diameter", "left");
                double lateralOffset  = GetWheelValue(wheelsData, "Lateral Offset", "left", 0.0);
                double verticalOffset = GetWheelValue(wheelsData, "Vertical Offset", "left", 0.0);
                double longOffset     = GetWheelValue(wheelsData, "Longitudinal Offset", "left", 0.0);
                double camber         = GetWheelValue(wheelsData, "Static Camber", "left", 0.0);
                double toe            = GetWheelValue(wheelsData, "Static Toe", "left", 0.0);

                double xBase = referenceDistanceMm + longOffset;
                double yBase = halfTrack + lateralOffset;
                double z     = tireDiameter / 2.0 + verticalOffset;

                // IMPORTANT:
                // Derive wheel naming from suffix, not caller-provided bool, to prevent
                // mixed names like RL_wheel_FRONT / FL_wheel_REAR when a caller passes
                // an incorrect rear/front flag.
                bool suffixIsRear = string.Equals(suffix, "_REAR", StringComparison.OrdinalIgnoreCase);
                if (isRear != suffixIsRear)
                {
                    Console.WriteLine($"Warning: Wheel extraction flag mismatch (isRear={isRear}, suffix='{suffix}'). Using suffix for naming.");
                }

                string prefix = suffixIsRear ? "R" : "F";

                // Left wheel (positive Y)
                hardpoints.Add(new HardpointInfo
                {
                    BaseName = $"{prefix}L_wheel",
                    Suffix   = suffix,
                    X        = xBase,
                    Y        = yBase,
                    Z        = z,
                    AngleX   = camber,
                    AngleY   = 0,
                    AngleZ   = toe
                });

                // Right wheel (negative Y, mirrored angles)
                hardpoints.Add(new HardpointInfo
                {
                    BaseName = $"{prefix}R_wheel",
                    Suffix   = suffix,
                    X        = xBase,
                    Y        = -yBase,
                    Z        = z,
                    AngleX   = -camber,
                    AngleY   = 0,
                    AngleZ   = -toe
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not extract wheel data for suffix '{suffix}': {ex.Message}");
            }
        }

        private static Dictionary<string, object> ParseJson(string json)
        {
            // Very simple JSON parser for basic structures
            var result = new Dictionary<string, object>();
            json = json.Trim();
            
            if (json.StartsWith("{") && json.EndsWith("}"))
            {
                json = json.Substring(1, json.Length - 2);
                var pairs = SplitJsonPairs(json);
                
                foreach (var pair in pairs)
                {
                    var colonIndex = pair.IndexOf(':');
                    if (colonIndex > 0)
                    {
                        var key = pair.Substring(0, colonIndex).Trim().Trim('"');
                        var value = pair.Substring(colonIndex + 1).Trim();
                        
                        result[key] = ParseJsonValue(value);
                    }
                }
            }
            
            return result;
        }

        private static object ParseJsonValue(string value)
        {
            value = value.Trim();
            
            if (value.StartsWith("[") && value.EndsWith("]"))
            {
                // Parse array
                value = value.Substring(1, value.Length - 2);
                var items = SplitJsonArray(value);
                var list = new List<object>();
                
                foreach (var item in items)
                {
                    list.Add(ParseJsonValue(item.Trim()));
                }
                
                return list;
            }
            else if (value.StartsWith("{") && value.EndsWith("}"))
            {
                // Parse object
                return ParseJson(value);
            }
            else if (double.TryParse(value, out double number))
            {
                return number;
            }
            else
            {
                return value.Trim('"');
            }
        }

        private static List<string> SplitJsonPairs(string json)
        {
            var pairs = new List<string>();
            int braceCount = 0;
            int bracketCount = 0;
            int quoteCount = 0;
            int start = 0;
            
            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];
                
                if (c == '"' && (i == 0 || json[i - 1] != '\\'))
                {
                    quoteCount++;
                }
                else if (quoteCount % 2 == 0)
                {
                    if (c == '{') braceCount++;
                    else if (c == '}') braceCount--;
                    else if (c == '[') bracketCount++;
                    else if (c == ']') bracketCount--;
                    else if (c == ',' && braceCount == 0 && bracketCount == 0)
                    {
                        pairs.Add(json.Substring(start, i - start));
                        start = i + 1;
                    }
                }
            }
            
            if (start < json.Length)
            {
                pairs.Add(json.Substring(start));
            }
            
            return pairs;
        }

        private static List<string> SplitJsonArray(string array)
        {
            var items = new List<string>();
            int braceCount = 0;
            int bracketCount = 0;
            int quoteCount = 0;
            int start = 0;
            
            for (int i = 0; i < array.Length; i++)
            {
                char c = array[i];
                
                if (c == '"' && (i == 0 || array[i - 1] != '\\'))
                {
                    quoteCount++;
                }
                else if (quoteCount % 2 == 0)
                {
                    if (c == '{') braceCount++;
                    else if (c == '}') braceCount--;
                    else if (c == '[') bracketCount++;
                    else if (c == ']') bracketCount--;
                    else if (c == ',' && braceCount == 0 && bracketCount == 0)
                    {
                        items.Add(array.Substring(start, i - start));
                        start = i + 1;
                    }
                }
            }
            
            if (start < array.Length)
            {
                items.Add(array.Substring(start));
            }
            
            return items;
        }

        private static ModelDoc2 LoadMarkerPart(SldWorks swApp, string markerPath)
        {
            ModelDoc2 existingMarker = swApp.GetOpenDocumentByName(markerPath) as ModelDoc2;
            if (existingMarker != null)
            {
                Console.WriteLine("Closing previously opened Marker document...");
                swApp.CloseDoc(existingMarker.GetTitle());
            }

            int openErrors = 0;
            int openWarnings = 0;
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
                return null;
            }

            Console.WriteLine($"Marker document loaded: {markerDoc.GetTitle()}");
            return markerDoc;
        }

        private static BodyInfo InsertMarkerBody(PartDoc swPart, ModelDoc2 markerDoc, HardpointInfo hardpoint)
        {
            try
            {
                // Convert mm to meters for SolidWorks
                double x = hardpoint.X / 1000.0;
                double y = hardpoint.Y / 1000.0;
                double z = hardpoint.Z / 1000.0;

                // Get body count before insertion
                object bodiesObjBefore = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
                int bodyCountBefore = 0;
                if (bodiesObjBefore != null)
                {
                    object[] bodiesBefore = bodiesObjBefore as object[];
                    bodyCountBefore = bodiesBefore != null ? bodiesBefore.Length : 0;
                }

                // Use InsertPart2 to insert the marker as a solid body
                Feature insertFeature = swPart.InsertPart2(markerDoc.GetPathName(), 0);
                
                if (insertFeature == null)
                {
                    Console.WriteLine($"Failed to insert part for {hardpoint.BaseName}");
                    return null;
                }

                // Get the newly inserted bodies
                object bodiesObjAfter = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
                object[] bodiesAfter = bodiesObjAfter as object[];
                
                if (bodiesAfter != null && bodiesAfter.Length > bodyCountBefore)
                {
                    // The last body should be the one we just inserted
                    Body2 body = bodiesAfter[bodiesAfter.Length - 1] as Body2;
                    if (body != null)
                    {
                        return new BodyInfo
                        {
                            Body = body,
                            OriginalName = body.Name,
                            BaseName = hardpoint.BaseName,
                            Suffix = hardpoint.Suffix,
                            X = x,
                            Y = y,
                            Z = z
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting body for {hardpoint.BaseName}: {ex.Message}");
            }

            return null;
        }

        private static void RenameBody(PartDoc swPart, BodyInfo bodyInfo)
        {
            try
            {
                string newName = $"{bodyInfo.BaseName}{bodyInfo.Suffix}";
                bodyInfo.Body.Name = newName;
                bodyInfo.RenamedName = newName;
                Console.WriteLine($"Renamed body to: {newName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming body: {ex.Message}");
            }
        }

        private static void ApplyColorToBody(PartDoc swPart, BodyInfo bodyInfo)
        {
            try
            {
                int[] rgb = GetColorForName(bodyInfo.BaseName);
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

                bodyInfo.Body.MaterialPropertyValues = matProps;
                Console.WriteLine($"Applied color RGB({rgb[0]},{rgb[1]},{rgb[2]}) to {bodyInfo.RenamedName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying color: {ex.Message}");
            }
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

        private static void CreateCoordinateSystem(ModelDoc2 swModel, BodyInfo bodyInfo)
        {
            try
            {
                // Coordinate system name: CS_<BodyName> where BodyName includes _FRONT/_REAR
                string csName = $"CS_{bodyInfo.RenamedName}";

                SketchManager sketchMgr = swModel.SketchManager;
                FeatureManager featMgr = swModel.FeatureManager;

                // Insert 3D sketch with a point at the hardpoint location
                swModel.ClearSelection2(true);
                sketchMgr.Insert3DSketch(true);
                sketchMgr.CreatePoint(bodyInfo.X, bodyInfo.Y, bodyInfo.Z);
                sketchMgr.Insert3DSketch(true);

                // Get the sketch feature and name it
                Feature sketchFeat = (Feature)swModel.FeatureByPositionReverse(0);
                string sketchName = csName + "_RefSketch";
                if (sketchFeat != null)
                {
                    sketchFeat.Name = sketchName;
                }

                // Select the sketch point as origin for coordinate system
                swModel.ClearSelection2(true);
                bool selected = swModel.Extension.SelectByID2(
                    "Point1@" + sketchName, "EXTSKETCHPOINT",
                    bodyInfo.X, bodyInfo.Y, bodyInfo.Z,
                    false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!selected)
                {
                    selected = swModel.Extension.SelectByID2(
                        "", "EXTSKETCHPOINT",
                        bodyInfo.X, bodyInfo.Y, bodyInfo.Z,
                        false, 1, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                // Create coordinate system
                Feature csFeat = featMgr.InsertCoordinateSystem(false, false, false);
                if (csFeat != null)
                {
                    csFeat.Name = csName;
                    Console.WriteLine($"Created coordinate system: {csName}");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not create coordinate system {csName}");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating coordinate system for {bodyInfo.RenamedName}: {ex.Message}");
            }
        }

        private static void CreateHardpointsFolder(PartDoc swPart, List<BodyInfo> bodies)
        {
            try
            {
                // Get the feature manager and create a folder
                ModelDoc2 swModel = (ModelDoc2)swPart;
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Select all the bodies to group them
                swModel.ClearSelection2(true);
                foreach (var bodyInfo in bodies)
                {
                    bodyInfo.Body.Select2(false, null);
                }

                // Create feature folder
                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                if (folder != null)
                {
                    folder.Name = "Hardpoints";
                    Console.WriteLine("Created Hardpoints folder");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Hardpoints folder: {ex.Message}");
            }
        }

        public static bool RunPose(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: hardpoints pose <jsonPath> <configName>");
                return false;
            }

            string jsonPath = args[2];
            string configName = args[3];

            ReportState(HardpointState.Initializing);

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

            try
            {
                // Load JSON data
                ReportState(HardpointState.LoadingJson);
                var jsonData = LoadJsonData(jsonPath);
                var hardpoints = ExtractHardpoints(jsonData);

                // Get all virtual parts in the assembly
                ReportState(HardpointState.LoadingMarkerPart);
                object componentsObj = swAssy.GetComponents(false);
                object[] components = componentsObj as object[];
                
                if (components == null || components.Length == 0)
                {
                    Console.WriteLine("Error: No components found in assembly");
                    return false;
                }

                // Precompute total steps
                int totalSteps = 0;
                totalSteps += hardpoints.Count; // Loading JSON
                totalSteps += components.Length; // Processing existing components
                totalSteps += hardpoints.Count; // Creating coordinate systems
                totalSteps += hardpoints.Count; // Creating mates
                totalSteps += 1; // Creating coordinate system folder
                totalSteps += 1; // Rebuilding

                Console.WriteLine($"TOTAL:{totalSteps}");
                int progressCount = 0;

                // Find virtual parts that match our hardpoints
                var matchingParts = new List<VirtualPartTransformInfo>();
                foreach (Component2 comp in components)
                {
                    string compName = comp.Name2;
                    foreach (var hardpoint in hardpoints)
                    {
                        string expectedName = $"{hardpoint.BaseName}{hardpoint.Suffix}";
                        if (compName == expectedName)
                        {
                            matchingParts.Add(new VirtualPartTransformInfo
                            {
                                Component = comp,
                                Hardpoint = hardpoint,
                                CoordSystemName = $"{configName}_{hardpoint.BaseName}{hardpoint.Suffix}"
                            });
                            break;
                        }
                    }
                    progressCount++;
                    ReportProgress(progressCount);
                }

                ReportState(HardpointState.CreatingCoordinateSystems);
                foreach (var partTransform in matchingParts)
                {
                    CreatePoseCoordinateSystem(swModel, partTransform, configName);
                    progressCount++;
                    ReportProgress(progressCount);
                }

                ReportState(HardpointState.CreatingCoordinateSystems);
                foreach (var partTransform in matchingParts)
                {
                    UpdateMateToPoseCoordinateSystem(swModel, swAssy, partTransform);
                    progressCount++;
                    ReportProgress(progressCount);
                }

                ReportState(HardpointState.CreatingTransformsFolder);
                CreatePoseCoordinateSystemFolder(swModel, matchingParts, configName);
                progressCount++;
                ReportProgress(progressCount);

                swModel.EditRebuild3();
                progressCount++;
                ReportProgress(progressCount);

                ReportState(HardpointState.Complete);
                Console.WriteLine($"Successfully updated pose for {matchingParts.Count} virtual parts");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Insert pose for active configuration only:
        /// - create assembly coordinate systems for all hardpoints
        /// - hide the new coordinate systems
        /// - mate each hardpoint component origin to corresponding pose coordinate system
        /// - place all pose coordinate systems into "&lt;PoseName&gt; Transforms" folder
        /// </summary>
        public static bool RunInsertPose(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: hardpoints insertpose <jsonPath> <poseName>");
                return false;
            }

            string jsonPath = args[2];
            string poseName = args[3];

            ReportState(HardpointState.Initializing);

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

            try
            {
                ReportState(HardpointState.LoadingJson);

                string jsonDir = Path.GetDirectoryName(jsonPath);
                string frontJsonPath = Path.Combine(jsonDir, "Front_Suspension.json");
                string rearJsonPath = Path.Combine(jsonDir, "Rear_Suspension.json");
                double rearReferenceOffsetMm = LoadRearReferenceOffsetMm(jsonDir);

                var hardpoints = new List<HardpointInfo>();

                if (File.Exists(frontJsonPath))
                {
                    var frontData = LoadJsonData(frontJsonPath);
                    ExtractHardpointsWithSuffix(frontData, "_FRONT", hardpoints, 0.0);
                    ExtractWheelsFromJson(frontData, "_FRONT", hardpoints, false, 0.0);
                }

                if (File.Exists(rearJsonPath))
                {
                    var rearData = LoadJsonData(rearJsonPath);
                    ExtractHardpointsWithSuffix(rearData, "_REAR", hardpoints, rearReferenceOffsetMm);
                    ExtractWheelsFromJson(rearData, "_REAR", hardpoints, true, rearReferenceOffsetMm);
                }

                if (hardpoints.Count == 0)
                {
                    Console.WriteLine("Error: No hardpoints found in Front/Rear JSON files");
                    return false;
                }

                ConfigurationManager cfgMgr = swModel.ConfigurationManager;
                Configuration activeCfg = cfgMgr != null ? cfgMgr.ActiveConfiguration : null;
                if (activeCfg == null)
                {
                    Console.WriteLine("Error: Could not determine active configuration");
                    return false;
                }
                string activeConfigName = activeCfg.Name;
                Console.WriteLine($"Active configuration: {activeConfigName}");

                object componentsObj = swAssy.GetComponents(false);
                object[] components = componentsObj as object[];
                if (components == null || components.Length == 0)
                {
                    Console.WriteLine("Error: No components found in assembly");
                    return false;
                }

                List<Component2> hardpointComponents = InsertPoseGetComponentsForPoseMatching(swModel, components);
                Console.WriteLine($"Pose matching component pool: {hardpointComponents.Count}");

                int totalSteps = hardpoints.Count * 2 + 2;
                Console.WriteLine($"TOTAL:{totalSteps}");
                int progressCount = 0;

                var poseCoordFeatures = new List<Feature>();

                ReportState(HardpointState.CreatingCoordinateSystems);
                foreach (var hardpoint in hardpoints)
                {
                    string hardpointName = $"{hardpoint.BaseName}{hardpoint.Suffix}";
                    string poseCoordName = $"{poseName} {hardpointName}";

                    // Reuse InsertCoordinate implementation for coordinate creation behavior.
                    Feature csFeat = InsertCoordinate.InsertCoordinateSystemFeature(
                        swModel,
                        poseCoordName,
                        hardpoint.X,
                        hardpoint.Y,
                        hardpoint.Z,
                        hardpoint.AngleX,
                        hardpoint.AngleY,
                        hardpoint.AngleZ,
                        createAtOrigin: false,
                        folderName: null,
                        hideInGui: true);

                    if (csFeat != null)
                    {
                        InsertPoseSetFeatureToActiveConfigurationOnly(swModel, csFeat, activeConfigName);
                        poseCoordFeatures.Add(csFeat);
                    }

                    progressCount++;
                    ReportProgress(progressCount);

                    Component2 targetComp = InsertPoseFindComponentByHardpointName(hardpointComponents, hardpointName);
                    if (targetComp != null)
                    {
                        Console.WriteLine($"Matched hardpoint '{hardpointName}' to component '{targetComp.Name2}'");
                        InsertPoseCreateCoordinateSystemMate(
                            swModel,
                            swAssy,
                            targetComp,
                            hardpointName,
                            poseCoordName,
                            activeConfigName);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Could not find hardpoint component '{hardpointName}' for mating");
                    }

                    progressCount++;
                    ReportProgress(progressCount);
                }

                ReportState(HardpointState.CreatingTransformsFolder);
                InsertPoseCreateOrPopulateTransformsFolder(swModel, poseCoordFeatures, $"{poseName} Transforms");
                progressCount++;
                ReportProgress(progressCount);

                ReportState(HardpointState.Rebuilding);
                swModel.EditRebuild3();
                progressCount++;
                ReportProgress(progressCount);

                ReportState(HardpointState.Complete);
                Console.WriteLine($"Inserted pose '{poseName}' for {poseCoordFeatures.Count} hardpoints in active config '{activeConfigName}'");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }

        private static void InsertPoseSetFeatureToActiveConfigurationOnly(ModelDoc2 swModel, Feature feat, string activeConfigName)
        {
            if (swModel == null || feat == null || string.IsNullOrWhiteSpace(activeConfigName))
            {
                return;
            }

            try
            {
                object configNamesObj = swModel.GetConfigurationNames();
                string[] configNames = configNamesObj as string[];
                if (configNames == null || configNames.Length == 0)
                {
                    return;
                }

                foreach (string cfgName in configNames)
                {
                    int state = string.Equals(cfgName, activeConfigName, StringComparison.OrdinalIgnoreCase)
                        ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature
                        : (int)swFeatureSuppressionAction_e.swSuppressFeature;

                    feat.SetSuppression2(
                        state,
                        (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                        new string[] { cfgName });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not scope feature '{feat.Name}' to active configuration: {ex.Message}");
            }
        }

        private static Component2 InsertPoseFindComponentByHardpointName(List<Component2> components, string hardpointName)
        {
            if (components == null || string.IsNullOrWhiteSpace(hardpointName))
            {
                return null;
            }

            string targetNormalized = InsertPoseNormalizeComponentName(hardpointName);
            Component2 startsWithMatch = null;
            Component2 containsMatch = null;

            foreach (Component2 comp in components)
            {
                if (comp == null)
                {
                    continue;
                }

                string compName = comp.Name2 ?? string.Empty;
                string compNormalized = InsertPoseNormalizeComponentName(compName);

                if (string.Equals(compName, hardpointName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(compNormalized, targetNormalized, StringComparison.OrdinalIgnoreCase))
                {
                    return comp;
                }

                if (startsWithMatch == null &&
                    (compName.StartsWith(hardpointName, StringComparison.OrdinalIgnoreCase) ||
                     compNormalized.StartsWith(targetNormalized, StringComparison.OrdinalIgnoreCase)))
                {
                    startsWithMatch = comp;
                }

                if (containsMatch == null &&
                    (compName.IndexOf(hardpointName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                     compNormalized.IndexOf(targetNormalized, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    containsMatch = comp;
                }
            }

            return startsWithMatch ?? containsMatch;
        }

        private static List<Component2> InsertPoseGetComponentsForPoseMatching(ModelDoc2 swModel, object[] allComponents)
        {
            var result = new List<Component2>();
            if (allComponents == null)
            {
                return result;
            }

            HashSet<string> hardpointFolderNames = InsertPoseGetHardpointFolderComponentNames(swModel);
            if (hardpointFolderNames.Count == 0)
            {
                foreach (object obj in allComponents)
                {
                    Component2 comp = obj as Component2;
                    if (comp != null)
                    {
                        result.Add(comp);
                    }
                }

                Console.WriteLine($"Warning: Hardpoints folder not found/empty. Using all {result.Count} components for pose matching.");
                return result;
            }

            foreach (object obj in allComponents)
            {
                Component2 comp = obj as Component2;
                if (comp == null)
                {
                    continue;
                }

                string normalized = InsertPoseNormalizeComponentName(comp.Name2);
                if (hardpointFolderNames.Contains(normalized))
                {
                    result.Add(comp);
                }
            }

            if (result.Count == 0)
            {
                foreach (object obj in allComponents)
                {
                    Component2 comp = obj as Component2;
                    if (comp != null)
                    {
                        result.Add(comp);
                    }
                }

                Console.WriteLine($"Warning: No components matched Hardpoints folder names. Falling back to all {result.Count} components.");
            }

            return result;
        }

        private static HashSet<string> InsertPoseGetHardpointFolderComponentNames(ModelDoc2 swModel)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (swModel == null)
            {
                return names;
            }

            try
            {
                Feature hardpointsFolder = InsertPoseFindFolder(swModel, "Hardpoints");
                if (hardpointsFolder == null)
                {
                    return names;
                }

                // Try IFeatureFolder.GetFeatures (SOLIDWORKS API example style)
                FeatureFolder featureFolder = hardpointsFolder.GetSpecificFeature2() as FeatureFolder;
                if (featureFolder != null)
                {
                    object[] folderFeatures = featureFolder.GetFeatures() as object[];
                    if (folderFeatures != null)
                    {
                        foreach (object obj in folderFeatures)
                        {
                            Feature feat = obj as Feature;
                            if (feat == null)
                            {
                                continue;
                            }

                            string normalized = InsertPoseNormalizeComponentName(feat.Name);
                            if (!string.IsNullOrWhiteSpace(normalized))
                            {
                                names.Add(normalized);
                            }
                        }
                    }
                }

                // Fallback: iterate subfeatures under folder
                if (names.Count == 0)
                {
                    Feature subFeat = (Feature)hardpointsFolder.GetFirstSubFeature();
                    while (subFeat != null)
                    {
                        string normalized = InsertPoseNormalizeComponentName(subFeat.Name);
                        if (!string.IsNullOrWhiteSpace(normalized))
                        {
                            names.Add(normalized);
                        }

                        subFeat = (Feature)subFeat.GetNextSubFeature();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read Hardpoints folder contents: {ex.Message}");
            }

            return names;
        }

        private static string InsertPoseNormalizeComponentName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }

            string normalized = name.Trim();

            // Strip display wrappers like "[ ... ]" if present
            int leftBracket = normalized.IndexOf('[');
            int rightBracket = normalized.IndexOf(']');
            if (leftBracket >= 0 && rightBracket > leftBracket)
            {
                normalized = normalized.Substring(leftBracket + 1, rightBracket - leftBracket - 1).Trim();
            }

            // Strip assembly path/instance decorations
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

            // Strip trailing instance suffix like "-1", "-23"
            int dashIdx = normalized.LastIndexOf('-');
            if (dashIdx > 0 && dashIdx < normalized.Length - 1)
            {
                bool allDigits = true;
                for (int i = dashIdx + 1; i < normalized.Length; i++)
                {
                    if (!char.IsDigit(normalized[i]))
                    {
                        allDigits = false;
                        break;
                    }
                }

                if (allDigits)
                {
                    normalized = normalized.Substring(0, dashIdx);
                }
            }

            return normalized.Trim();
        }

        private static void InsertPoseCreateCoordinateSystemMate(
            ModelDoc2 swModel,
            AssemblyDoc swAssy,
            Component2 comp,
            string componentCoordName,
            string poseCoordName,
            string activeConfigName)
        {
            if (swModel == null || swAssy == null || comp == null ||
                string.IsNullOrWhiteSpace(componentCoordName) ||
                string.IsNullOrWhiteSpace(poseCoordName))
            {
                return;
            }

            try
            {
                swModel.ClearSelection2(true);

                bool csSelected = swModel.Extension.SelectByID2(
                    poseCoordName,
                    "COORDSYS",
                    0, 0, 0,
                    false,
                    1,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!csSelected)
                {
                    Console.WriteLine($"Warning: Could not select pose CSys '{poseCoordName}' for mate");
                    swModel.ClearSelection2(true);
                    return;
                }

                bool compCsSelected = swModel.Extension.SelectByID2(
                    componentCoordName + "@" + comp.Name2 + "@" + swModel.GetTitle(),
                    "COORDSYS",
                    0, 0, 0,
                    true,
                    2,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!compCsSelected)
                {
                    Console.WriteLine($"Warning: Could not select component CSys '{componentCoordName}' for '{comp.Name2}'");
                    swModel.ClearSelection2(true);
                    return;
                }

                bool mateCreated = InsertPoseCreateCoincidentMateFromSelection(
                    swModel,
                    swAssy,
                    (int)swMateAlign_e.swMateAlignALIGNED,
                    out Feature mateFeature);

                if (!mateCreated)
                {
                    Console.WriteLine($"Warning: Could not create coincident mate via CreateMateData for '{comp.Name2}'");
                }
                else if (mateFeature != null)
                {
                    bool axisAligned = TryEnableCoincidentMateAxisAlignment(swModel, mateFeature);
                    if (axisAligned)
                    {
                        Console.WriteLine($"Enabled coincident mate axis alignment for '{comp.Name2}'");
                    }

                    InsertPoseSetFeatureToActiveConfigurationOnly(swModel, mateFeature, activeConfigName);
                }

                // Explicit axis-to-axis aligned mate (X-axis) to enforce orientation.
                bool axisMateCreated = InsertPoseCreateCoordinateSystemAxisMate(
                    swModel,
                    swAssy,
                    comp,
                    componentCoordName,
                    poseCoordName,
                    activeConfigName,
                    "X");

                if (!axisMateCreated)
                {
                    Console.WriteLine($"Warning: Could not create explicit X-axis aligned mate for '{comp.Name2}'");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to mate '{comp.Name2}' to '{poseCoordName}': {ex.Message}");
                try { swModel.ClearSelection2(true); } catch { }
            }
        }

        private static bool InsertPoseCreateCoordinateSystemAxisMate(
            ModelDoc2 swModel,
            AssemblyDoc swAssy,
            Component2 comp,
            string componentCoordName,
            string poseCoordName,
            string activeConfigName,
            string axisLetter)
        {
            if (swModel == null || swAssy == null || comp == null ||
                string.IsNullOrWhiteSpace(componentCoordName) ||
                string.IsNullOrWhiteSpace(poseCoordName) ||
                string.IsNullOrWhiteSpace(axisLetter))
            {
                return false;
            }

            try
            {
                swModel.ClearSelection2(true);

                bool poseAxisSelected = InsertPoseSelectCoordinateSystemAxisByName(
                    swModel,
                    poseCoordName,
                    null,
                    axisLetter,
                    false,
                    1);

                if (!poseAxisSelected)
                {
                    swModel.ClearSelection2(true);
                    return false;
                }

                bool compAxisSelected = InsertPoseSelectCoordinateSystemAxisByName(
                    swModel,
                    componentCoordName,
                    comp,
                    axisLetter,
                    true,
                    2);

                if (!compAxisSelected)
                {
                    swModel.ClearSelection2(true);
                    return false;
                }

                bool axisMateCreated = InsertPoseCreateCoincidentMateFromSelection(
                    swModel,
                    swAssy,
                    (int)swMateAlign_e.swMateAlignALIGNED,
                    out Feature mateFeature);

                if (!axisMateCreated)
                {
                    swModel.ClearSelection2(true);
                    return false;
                }

                if (mateFeature != null)
                {
                    TryEnableCoincidentMateAxisAlignment(swModel, mateFeature);
                    InsertPoseSetFeatureToActiveConfigurationOnly(swModel, mateFeature, activeConfigName);
                }

                swModel.ClearSelection2(true);
                return true;
            }
            catch
            {
                try { swModel.ClearSelection2(true); } catch { }
                return false;
            }
        }

        private static bool InsertPoseSelectCoordinateSystemAxisByName(
            ModelDoc2 swModel,
            string coordSystemName,
            Component2 component,
            string axisLetter,
            bool append,
            int selectionMark)
        {
            if (swModel == null || string.IsNullOrWhiteSpace(coordSystemName) || string.IsNullOrWhiteSpace(axisLetter))
            {
                return false;
            }

            string asmTitle = swModel.GetTitle();
            string compName = component != null ? (component.Name2 ?? string.Empty) : null;
            string axisPrefix = axisLetter.Trim().ToUpperInvariant();

            var candidates = new List<string>();
            if (component == null)
            {
                candidates.Add($"{axisPrefix} Axis1@{coordSystemName}");
                candidates.Add($"{axisPrefix} Axis@{coordSystemName}");
                candidates.Add($"{axisPrefix}Axis1@{coordSystemName}");
                candidates.Add($"{axisPrefix}Axis@{coordSystemName}");
            }
            else
            {
                candidates.Add($"{axisPrefix} Axis1@{coordSystemName}@{compName}@{asmTitle}");
                candidates.Add($"{axisPrefix} Axis@{coordSystemName}@{compName}@{asmTitle}");
                candidates.Add($"{axisPrefix}Axis1@{coordSystemName}@{compName}@{asmTitle}");
                candidates.Add($"{axisPrefix}Axis@{coordSystemName}@{compName}@{asmTitle}");
            }

            foreach (string candidate in candidates)
            {
                bool selected = swModel.Extension.SelectByID2(
                    candidate,
                    "AXIS",
                    0, 0, 0,
                    append,
                    selectionMark,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (selected)
                {
                    return true;
                }

                selected = swModel.Extension.SelectByID2(
                    candidate,
                    "DATUMAXIS",
                    0, 0, 0,
                    append,
                    selectionMark,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (selected)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool InsertPoseCreateCoincidentMateFromSelection(
            ModelDoc2 swModel,
            AssemblyDoc swAssy,
            int mateAlignment,
            out Feature createdMateFeature)
        {
            createdMateFeature = null;

            if (swModel == null || swAssy == null)
            {
                return false;
            }

            try
            {
                SelectionMgr selectionMgr = swModel.SelectionManager as SelectionMgr;
                if (selectionMgr == null)
                {
                    return false;
                }

                object entity1 = selectionMgr.GetSelectedObject6(1, -1);
                object entity2 = selectionMgr.GetSelectedObject6(2, -1);
                if (entity1 == null || entity2 == null)
                {
                    return false;
                }

                object coincMateData = swAssy.GetType().InvokeMember(
                    "CreateMateData",
                    System.Reflection.BindingFlags.InvokeMethod |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance,
                    null,
                    swAssy,
                    new object[] { (int)swMateType_e.swMateCOINCIDENT });

                if (coincMateData == null)
                {
                    return false;
                }

                Type mateDataType = coincMateData.GetType();
                var entitiesProp = mateDataType.GetProperty("EntitiesToMate");
                if (entitiesProp == null || !entitiesProp.CanWrite)
                {
                    return false;
                }

                entitiesProp.SetValue(coincMateData, new object[] { entity1, entity2 }, null);
                TrySetIntProperty(coincMateData, "MateAlignment", mateAlignment);

                object createdMate = swAssy.GetType().InvokeMember(
                    "CreateMate",
                    System.Reflection.BindingFlags.InvokeMethod |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance,
                    null,
                    swAssy,
                    new object[] { coincMateData });

                if (createdMate == null)
                {
                    return false;
                }

                createdMateFeature = (Feature)swModel.FeatureByPositionReverse(0);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void InsertPoseCreateOrPopulateTransformsFolder(ModelDoc2 swModel, List<Feature> coordFeatures, string folderName)
        {
            if (swModel == null || coordFeatures == null || coordFeatures.Count == 0 || string.IsNullOrWhiteSpace(folderName))
            {
                return;
            }

            try
            {
                FeatureManager featMgr = swModel.FeatureManager;
                if (featMgr == null)
                {
                    return;
                }

                Feature existingFolder = InsertPoseFindFolder(swModel, folderName);
                if (existingFolder != null)
                {
                    foreach (Feature feat in coordFeatures)
                    {
                        if (feat != null)
                        {
                            featMgr.MoveToFolder(folderName, feat.Name, false);
                        }
                    }
                    Console.WriteLine($"Added coordinate systems to existing folder '{folderName}'");
                    return;
                }

                swModel.ClearSelection2(true);
                int selectedCount = 0;
                foreach (Feature feat in coordFeatures)
                {
                    if (feat != null && feat.Select2(true, 0))
                    {
                        selectedCount++;
                    }
                }

                if (selectedCount == 0)
                {
                    swModel.ClearSelection2(true);
                    return;
                }

                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created folder '{folderName}'");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to create/populate folder '{folderName}': {ex.Message}");
                try { swModel.ClearSelection2(true); } catch { }
            }
        }

        private static Feature InsertPoseFindFolder(ModelDoc2 swModel, string folderName)
        {
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "FtrFolder" &&
                    string.Equals(feat.Name, folderName, StringComparison.OrdinalIgnoreCase))
                {
                    return feat;
                }

                feat = (Feature)feat.GetNextFeature();
            }

            return null;
        }

        private static bool SelectComponentOriginForMate(ModelDoc2 swModel, Component2 comp, int selectionMark)
        {
            if (swModel == null || comp == null)
            {
                return false;
            }

            string assemblyTitle = swModel.GetTitle();
            string componentName = comp.Name2 ?? string.Empty;

            bool selected = swModel.Extension.SelectByID2(
                "Origin@" + componentName + "@" + assemblyTitle,
                "ORIGINFOLDER",
                0, 0, 0,
                true,
                selectionMark,
                null,
                (int)swSelectOption_e.swSelectOptionDefault);

            if (!selected)
            {
                selected = swModel.Extension.SelectByID2(
                    "Point1@Origin@" + componentName + "@" + assemblyTitle,
                    "EXTSKETCHPOINT",
                    0, 0, 0,
                    true,
                    selectionMark,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);
            }

            return selected;
        }

        private static bool TryEnableCoincidentMateAxisAlignment(ModelDoc2 swModel, Feature mateFeature)
        {
            if (swModel == null || mateFeature == null)
            {
                return false;
            }

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
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not enable coincident mate axis alignment: {ex.Message}");
                return false;
            }
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

        private static void CreateOrUpdateTransform(PartDoc swPart, BodyTransformInfo bodyTransform, string configName)
        {
            try
            {
                // Check if transform already exists
                Feature existingTransform = FindTransformFeature(swPart, bodyTransform.TransformName);
                
                if (existingTransform != null)
                {
                    // Update existing transform
                    UpdateTransform(swPart, existingTransform, bodyTransform);
                    Console.WriteLine($"Updated transform: {bodyTransform.TransformName}");
                }
                else
                {
                    // Create new transform
                    CreateTransform(swPart, bodyTransform);
                    Console.WriteLine($"Created transform: {bodyTransform.TransformName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating/updating transform for {bodyTransform.Body.Name}: {ex.Message}");
            }
        }

        private static Feature FindTransformFeature(PartDoc swPart, string transformName)
        {
            Feature feat = (Feature)swPart.FirstFeature();
            while (feat != null)
            {
                if (feat.Name == transformName)
                {
                    return feat;
                }
                feat = (Feature)feat.GetNextFeature();
            }
            return null;
        }

        private static void CreateTransform(PartDoc swPart, BodyTransformInfo bodyTransform)
        {
            // Create Move/Copy Body feature
            ModelDoc2 swModel = (ModelDoc2)swPart;
            FeatureManager featMgr = swModel.FeatureManager;
            
            // Select the body
            swModel.ClearSelection2(true);
            bodyTransform.Body.Select2(false, null);

            // Create Move/Copy Body feature
            // InsertMoveCopyBody2 signature: (transX, transY, transZ, rotX, rotY, rotZ, scaleX, scaleY, scaleZ, scale, copy, numPattern)
            Feature transformFeat = featMgr.InsertMoveCopyBody2(
                bodyTransform.Hardpoint.X / 1000.0,  // Translation X
                bodyTransform.Hardpoint.Y / 1000.0,  // Translation Y
                bodyTransform.Hardpoint.Z / 1000.0,  // Translation Z
                0, 0, 0,  // Rotation (radians)
                1.0, 1.0, 1.0,  // Scale X, Y, Z
                1.0,      // Uniform scale
                false,    // Copy (false = move)
                0         // Number of patterns
            );

            if (transformFeat != null)
            {
                transformFeat.Name = bodyTransform.TransformName;
                Console.WriteLine($"Created transform feature: {bodyTransform.TransformName}");
            }

            swModel.ClearSelection2(true);
        }

        private static void UpdateTransform(PartDoc swPart, Feature transformFeat, BodyTransformInfo bodyTransform)
        {
            // For simplicity, delete the old transform and create a new one
            // This is easier than trying to modify the existing transform's properties
            ModelDoc2 swModel = (ModelDoc2)swPart;
            
            // Delete the existing transform
            swModel.ClearSelection2(true);
            transformFeat.Select2(false, 0);
            swModel.EditDelete();
            
            // Create a new transform with updated position
            CreateTransform(swPart, bodyTransform);
        }

        private static void UpdateTransformSuppression(PartDoc swPart, List<BodyTransformInfo> transforms, string targetConfig)
        {
            try
            {
                // Get all configurations
                ModelDoc2 swModel = (ModelDoc2)swPart;
                object configNamesObj = swModel.GetConfigurationNames();
                string[] configNames = configNamesObj as string[];
                
                if (configNames == null || configNames.Length == 0)
                {
                    Console.WriteLine("No configurations found in part");
                    return;
                }
                
                foreach (var transform in transforms)
                {
                    Feature transformFeat = FindTransformFeature(swPart, transform.TransformName);
                    if (transformFeat == null) continue;

                    foreach (string configName in configNames)
                    {
                        bool shouldSuppress = configName != targetConfig;
                        
                        int suppressState = shouldSuppress
                            ? (int)swFeatureSuppressionAction_e.swSuppressFeature
                            : (int)swFeatureSuppressionAction_e.swUnSuppressFeature;

                        transformFeat.SetSuppression2(
                            suppressState,
                            (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                            configName
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating transform suppression: {ex.Message}");
            }
        }

        private static void CreateTransformsFolder(PartDoc swPart, List<BodyTransformInfo> transforms, string configName)
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)swPart;
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Select all transform features
                swModel.ClearSelection2(true);
                foreach (var transform in transforms)
                {
                    Feature transformFeat = FindTransformFeature(swPart, transform.TransformName);
                    if (transformFeat != null)
                    {
                        transformFeat.Select2(false, 0);
                    }
                }

                // Create feature folder
                string folderName = $"{configName} Transforms";
                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created {folderName} folder");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Transforms folder: {ex.Message}");
            }
        }

        // Assembly-specific methods for virtual part handling
        private static VirtualPartInfo InsertVirtualMarkerPart(AssemblyDoc swAssy, ModelDoc2 markerDoc, HardpointInfo hardpoint)
        {
            try
            {
                // Insert ALL marker parts at assembly origin (0,0,0).
                // Actual position comes in the pose step via mates.
                Component2 newComp = (Component2)swAssy.AddComponent4(markerDoc.GetPathName(), "", 0, 0, 0);

                if (newComp == null)
                {
                    Console.WriteLine($"Failed to insert virtual part for {hardpoint.BaseName}");
                    return null;
                }

                return new VirtualPartInfo
                {
                    Component = newComp,
                    OriginalName = newComp.Name2,
                    BaseName = hardpoint.BaseName,
                    Suffix = hardpoint.Suffix,
                    X = hardpoint.X / 1000.0,
                    Y = hardpoint.Y / 1000.0,
                    Z = hardpoint.Z / 1000.0,
                    AngleX = hardpoint.AngleX,
                    AngleY = hardpoint.AngleY,
                    AngleZ = hardpoint.AngleZ
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting virtual part for {hardpoint.BaseName}: {ex.Message}");
                return null;
            }
        }

        private static void ExtractHardpointsWithSuffix(Dictionary<string, object> jsonData, string suffix, List<HardpointInfo> hardpoints)
        {
            ExtractHardpointsWithSuffix(jsonData, suffix, hardpoints, 0.0);
        }

        private static void ExtractHardpointsWithSuffix(
            Dictionary<string, object> jsonData,
            string suffix,
            List<HardpointInfo> hardpoints,
            double xOffsetMm)
        {
            // Extract from all sections except Wheels
            foreach (var section in jsonData)
            {
                if (section.Key == "Wheels")
                {
                    continue; // Wheels handled separately if needed
                }

                // This is a suspension section (e.g., "Double A-Arm", "Push Pull")
                Dictionary<string, object> sectionObj = section.Value as Dictionary<string, object>;
                if (sectionObj != null)
                {
                    foreach (var pointProperty in sectionObj)
                    {
                        string pointName = pointProperty.Key;
                        List<object> coords = pointProperty.Value as List<object>;
                        if (coords != null && coords.Count >= 3)
                        {
                            hardpoints.Add(new HardpointInfo
                            {
                                BaseName = pointName,
                                Suffix = suffix,
                                X = Convert.ToDouble(coords[0]) + xOffsetMm,
                                Y = Convert.ToDouble(coords[1]),
                                Z = Convert.ToDouble(coords[2])
                            });
                        }
                    }
                }
            }
        }

        private static void MakeComponentVirtual(VirtualPartInfo partInfo)
        {
            try
            {
                // Make the component virtual using MakeVirtual2
                // Parameter: deleteExternalFile (false = keep the file reference but make virtual)
                bool success = partInfo.Component.MakeVirtual2(false);
                if (success)
                {
                    Console.WriteLine($"Made component virtual: {partInfo.Component.Name2}");
                }
                else
                {
                    Console.WriteLine($"Warning: Failed to make component virtual: {partInfo.Component.Name2}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error making component virtual: {ex.Message}");
            }
        }

        private static void RenameVirtualPart(AssemblyDoc swAssy, VirtualPartInfo partInfo)
        {
            try
            {
                string newName = $"{partInfo.BaseName}{partInfo.Suffix}";
                partInfo.Component.Name2 = newName;
                partInfo.RenamedName = newName;
                Console.WriteLine($"Renamed virtual part to: {newName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming virtual part: {ex.Message}");
            }
        }

        private static void FloatComponent(AssemblyDoc swAssy, ModelDoc2 swModel, VirtualPartInfo partInfo)
        {
            try
            {
                if (swAssy == null || swModel == null || partInfo?.Component == null)
                {
                    Console.WriteLine("Warning: Cannot float component because assembly/model/component is null");
                    return;
                }

                swModel.ClearSelection2(true);
                bool selected = partInfo.Component.Select4(false, null, false);

                // Fallback selection by name if direct component selection fails
                if (!selected)
                {
                    string assyTitle = swModel.GetTitle();
                    string compName = partInfo.Component.Name2;
                    selected = swModel.Extension.SelectByID2(
                        $"{compName}@{assyTitle}",
                        "COMPONENT",
                        0, 0, 0,
                        false,
                        0,
                        null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                if (!selected)
                {
                    Console.WriteLine($"Warning: Could not select component to float: {partInfo.RenamedName}");
                    swModel.ClearSelection2(true);
                    return;
                }

                swAssy.UnfixComponent();
                Console.WriteLine($"Floated component: {partInfo.RenamedName}");
                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error floating component {partInfo?.RenamedName}: {ex.Message}");
                try { swModel?.ClearSelection2(true); } catch { }
            }
        }

        /// <summary>
        /// Renames the coordinate system inside a virtual marker part using EditPart2.
        /// EditPart2 enters in-context editing of the selected virtual component.
        /// </summary>
        private static void RenameInternalCoordinateSystem(SldWorks swApp, AssemblyDoc swAssy, ModelDoc2 swModel, VirtualPartInfo partInfo)
        {
            string newCsName = string.IsNullOrWhiteSpace(partInfo.RenamedName)
                ? $"{partInfo.BaseName}{partInfo.Suffix}"
                : partInfo.RenamedName;
            try
            {
                // Select the virtual component
                swModel.ClearSelection2(true);
                partInfo.Component.Select4(false, null, false);

                // Enter in-context part editing via EditPart2
                // Parameters: Silent=true, AllowReadOnly=false, Information (out)
                int editInfo = 0;
                int editResult = swAssy.EditPart2(true, false, ref editInfo);

                if (editResult == (int)swEditPartCommandStatus_e.swEditPartSuccessful)
                {
                    // Prefer the component's part document (reliable in in-context editing)
                    ModelDoc2 partDoc = (ModelDoc2)partInfo.Component.GetModelDoc2();
                    if (partDoc == null)
                    {
                        partDoc = (ModelDoc2)swApp.ActiveDoc;
                    }

                    if (partDoc != null)
                    {
                        bool renamed = false;
                        Feature defaultCs = GetInternalCoordinateSystemFeature(partDoc);
                        if (defaultCs != null && defaultCs.GetTypeName2() == "CoordSys")
                        {
                            defaultCs.Name = newCsName;
                            renamed = true;
                        }

                        if (renamed)
                        {
                            Console.WriteLine($"Renamed internal coordinate system to: {newCsName}");
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Could not find internal coordinate system to rename for {partInfo.RenamedName}");
                        }

                        partDoc.EditRebuild3();
                    }
                    // Exit in-context editing, return to assembly
                    swAssy.EditAssembly();
                }
                else
                {
                    Console.WriteLine($"Warning: EditPart2 returned {editResult} (info={editInfo}) for {partInfo.RenamedName}");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming internal coordinate system for {partInfo.RenamedName}: {ex.Message}");
                // Try to recover by returning to assembly edit mode
                try { swAssy.EditAssembly(); } catch { }
            }
        }

        private static void RenameEmbeddedCoordinateSystem(ModelDoc2 swModel, VirtualPartInfo partInfo)
        {
            try
            {
                // Find the embedded coordinate system in the virtual part
                Feature feat = (Feature)swModel.FirstFeature();
                while (feat != null)
                {
                    if (feat.GetTypeName2() == "CoordSys")
                    {
                        // Check if this coordinate system belongs to our virtual part
                        // For now, we'll assume the first coordinate system found is the one we want
                        string csName = $"CS_{partInfo.RenamedName}";
                        feat.Name = csName;
                        Console.WriteLine($"Renamed embedded coordinate system to: {csName}");
                        break;
                    }
                    feat = (Feature)feat.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming embedded coordinate system: {ex.Message}");
            }
        }

        private static void ApplyColorToVirtualPart(AssemblyDoc swAssy, VirtualPartInfo partInfo)
        {
            try
            {
                int[] rgb = GetColorForName(partInfo.BaseName);
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

                partInfo.Component.MaterialPropertyValues = matProps;
                Console.WriteLine($"Applied color RGB({rgb[0]},{rgb[1]},{rgb[2]}) to {partInfo.RenamedName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying color to virtual part: {ex.Message}");
            }
        }

        private static void CreateInitialCoordinateSystem(ModelDoc2 swModel, VirtualPartInfo partInfo, string configName)
        {
            try
            {
                // Coordinate system name: <ConfigName>_<HardpointName>
                string csName = $"{configName}_{partInfo.BaseName}{partInfo.Suffix}";

                SketchManager sketchMgr = swModel.SketchManager;
                FeatureManager featMgr = swModel.FeatureManager;

                // Insert 3D sketch with a point at the hardpoint location
                swModel.ClearSelection2(true);
                sketchMgr.Insert3DSketch(true);
                sketchMgr.CreatePoint(partInfo.X, partInfo.Y, partInfo.Z);
                sketchMgr.Insert3DSketch(true);

                // Get the sketch feature and name it
                Feature sketchFeat = (Feature)swModel.FeatureByPositionReverse(0);
                string sketchName = csName + "_RefSketch";
                if (sketchFeat != null)
                {
                    sketchFeat.Name = sketchName;
                }

                // Select the sketch point as origin for coordinate system
                swModel.ClearSelection2(true);
                bool selected = swModel.Extension.SelectByID2(
                    "Point1@" + sketchName, "EXTSKETCHPOINT",
                    partInfo.X, partInfo.Y, partInfo.Z,
                    false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!selected)
                {
                    selected = swModel.Extension.SelectByID2(
                        "", "EXTSKETCHPOINT",
                        partInfo.X, partInfo.Y, partInfo.Z,
                        false, 1, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                // Create coordinate system
                Feature csFeat = featMgr.InsertCoordinateSystem(false, false, false);
                if (csFeat != null)
                {
                    csFeat.Name = csName;
                    // Hide the coordinate system
                    csFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    Console.WriteLine($"Created and hid coordinate system: {csName}");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not create coordinate system {csName}");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating initial coordinate system for {partInfo.RenamedName}: {ex.Message}");
            }
        }

        private static void CreateMateToCoordinateSystem(ModelDoc2 swModel, AssemblyDoc swAssy, VirtualPartInfo partInfo)
        {
            try
            {
                // Select the coordinate system origin and virtual part origin for mating
                swModel.ClearSelection2(true);
                
                // Select coordinate system origin (mark 1)
                bool sel1 = swModel.Extension.SelectByID2(
                    partInfo.BaseName + partInfo.Suffix, "COORDSYS", 0, 0, 0, false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!sel1)
                {
                    Console.WriteLine($"Warning: Could not select coordinate system for {partInfo.RenamedName}");
                    return;
                }

                // Select virtual part origin (mark 2)
                bool sel2 = SelectComponentOriginForMate(swModel, partInfo.Component, 2);

                if (!sel2)
                {
                    Console.WriteLine($"Warning: Could not select virtual part origin for {partInfo.RenamedName}");
                    swModel.ClearSelection2(true);
                    return;
                }

                // Add coincident mate with axis alignment
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
                            Console.WriteLine($"Enabled coincident mate axis alignment for {partInfo.RenamedName}");
                        }
                    }

                    Console.WriteLine($"Added coincident mate with axis alignment for {partInfo.RenamedName}");
                }
                else
                {
                    Console.WriteLine($"Note: Component positioned at coordinate system (mate skipped, error: {mateError})");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating mate for {partInfo.RenamedName}: {ex.Message}");
                swModel.ClearSelection2(true);
            }
        }


        private static void CreateWheelHardpointsFolder(AssemblyDoc swAssy, List<VirtualPartInfo> parts)
        {
            try
            {
                // Get the feature manager and create a folder
                ModelDoc2 swModel = (ModelDoc2)swAssy;
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Step 1: Create empty folder first (before any components)
                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore);
                if (folder != null)
                {
                    folder.Name = "Wheel Hardpoints";
                    Console.WriteLine("Created empty Wheel Hardpoints folder");
                }
                
                // Step 2: Get the folder feature by name
                folder = (Feature)swAssy.FeatureByName("Wheel Hardpoints");
                if (folder == null)
                {
                    Console.WriteLine("Warning: Could not find Wheel Hardpoints folder after creation");
                    return;
                }

                // Step 3: Build array of components to move
                var componentsToMove = new List<object>();
                foreach (var partInfo in parts)
                {
                    if (partInfo.Component != null)
                    {
                        componentsToMove.Add(partInfo.Component);
                    }
                }

                Console.WriteLine($"Moving {componentsToMove.Count} components to Wheel Hardpoints folder");

                // Step 4: Move components into the folder using ReorderComponents
                if (componentsToMove.Count > 0)
                {
                    object[] compArray = componentsToMove.ToArray();
                    bool success = swAssy.ReorderComponents(
                        compArray,
                        folder,
                        (int)swReorderComponentsWhere_e.swReorderComponents_LastInFolder);
                    
                    if (success)
                    {
                        Console.WriteLine($"Successfully moved {compArray.Length} components to Wheel Hardpoints folder");
                    }
                    else
                    {
                        Console.WriteLine("Warning: ReorderComponents returned false");
                    }
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Wheel Hardpoints folder: {ex.Message}");
            }
        }

        private static void CreateHardpointsFolder(AssemblyDoc swAssy, List<VirtualPartInfo> parts)
        {
            try
            {
                Console.WriteLine($"CreateHardpointsFolder called with {parts.Count} parts");
                
                // Get the feature manager and create a folder
                ModelDoc2 swModel = (ModelDoc2)swAssy;
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Step 1: Create empty folder first (before any components)
                Console.WriteLine("Creating empty folder...");
                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore);
                if (folder != null)
                {
                    folder.Name = "Hardpoints";
                    Console.WriteLine($"Created empty folder, named it: {folder.Name}");
                }
                else
                {
                    Console.WriteLine("InsertFeatureTreeFolder2 returned null");
                }
                
                // Step 2: Get the folder feature by name
                folder = (Feature)swAssy.FeatureByName("Hardpoints");
                if (folder == null)
                {
                    Console.WriteLine("Warning: Could not find Hardpoints folder after creation, trying to find 'Folder1'");
                    // The default name might be "Folder1" if the rename didn't work
                    folder = (Feature)swAssy.FeatureByName("Folder1");
                    if (folder != null)
                    {
                        folder.Name = "Hardpoints";
                        Console.WriteLine("Found and renamed Folder1 to Hardpoints");
                    }
                }
                
                if (folder == null)
                {
                    Console.WriteLine("ERROR: Could not find any folder feature after creation");
                    return;
                }

                Console.WriteLine($"Folder found: {folder.Name}");

                // Step 3: Build array of components to move - use stored Component references directly
                var componentsToMove = new List<object>();
                foreach (var partInfo in parts)
                {
                    if (partInfo.Component != null)
                    {
                        componentsToMove.Add(partInfo.Component);
                        Console.WriteLine($"  Will move: {partInfo.RenamedName}");
                    }
                    else
                    {
                        Console.WriteLine($"  WARNING: Null component for {partInfo.RenamedName}");
                    }
                }

                Console.WriteLine($"Total components to move: {componentsToMove.Count}");

                // Step 4: Move components into the folder using ReorderComponents
                if (componentsToMove.Count > 0)
                {
                    object[] compArray = componentsToMove.ToArray();
                    Console.WriteLine($"Calling ReorderComponents with {compArray.Length} components...");
                    bool success = swAssy.ReorderComponents(
                        compArray,
                        folder,
                        (int)swReorderComponentsWhere_e.swReorderComponents_LastInFolder);
                    
                    Console.WriteLine($"ReorderComponents result: {success}");
                }
                else
                {
                    Console.WriteLine("No components to move!");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Hardpoints folder: {ex.Message}");
                Console.WriteLine($"Stack: {ex.StackTrace}");
            }
        }


        private static void CreatePoseCoordinateSystem(ModelDoc2 swModel, VirtualPartTransformInfo partTransform, string configName)
        {
            try
            {
                // Coordinate system name: <ConfigName>_<HardpointName>
                string csName = partTransform.CoordSystemName;

                SketchManager sketchMgr = swModel.SketchManager;
                FeatureManager featMgr = swModel.FeatureManager;

                // Insert 3D sketch with a point at the hardpoint location
                swModel.ClearSelection2(true);
                sketchMgr.Insert3DSketch(true);
                sketchMgr.CreatePoint(partTransform.Hardpoint.X / 1000.0, partTransform.Hardpoint.Y / 1000.0, partTransform.Hardpoint.Z / 1000.0);
                sketchMgr.Insert3DSketch(true);

                // Get the sketch feature and name it
                Feature sketchFeat = (Feature)swModel.FeatureByPositionReverse(0);
                string sketchName = csName + "_RefSketch";
                if (sketchFeat != null)
                {
                    sketchFeat.Name = sketchName;
                }

                // Select the sketch point as origin for coordinate system
                swModel.ClearSelection2(true);
                bool selected = swModel.Extension.SelectByID2(
                    "Point1@" + sketchName, "EXTSKETCHPOINT",
                    partTransform.Hardpoint.X / 1000.0, partTransform.Hardpoint.Y / 1000.0, partTransform.Hardpoint.Z / 1000.0,
                    false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!selected)
                {
                    selected = swModel.Extension.SelectByID2(
                        "", "EXTSKETCHPOINT",
                        partTransform.Hardpoint.X / 1000.0, partTransform.Hardpoint.Y / 1000.0, partTransform.Hardpoint.Z / 1000.0,
                        false, 1, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                // Create coordinate system
                Feature csFeat = featMgr.InsertCoordinateSystem(false, false, false);
                if (csFeat != null)
                {
                    csFeat.Name = csName;
                    // Hide the coordinate system
                    csFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    Console.WriteLine($"Created and hid pose coordinate system: {csName}");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not create pose coordinate system {csName}");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating pose coordinate system for {partTransform.Component.Name2}: {ex.Message}");
            }
        }

        private static void UpdateMateToPoseCoordinateSystem(ModelDoc2 swModel, AssemblyDoc swAssy, VirtualPartTransformInfo partTransform)
        {
            try
            {
                // Select the pose coordinate system origin and virtual part origin for mating
                swModel.ClearSelection2(true);
                
                // Select pose coordinate system origin (mark 1)
                bool sel1 = swModel.Extension.SelectByID2(
                    partTransform.CoordSystemName, "COORDSYS", 0, 0, 0, false, 1, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!sel1)
                {
                    Console.WriteLine($"Warning: Could not select pose coordinate system for {partTransform.Component.Name2}");
                    return;
                }

                // Select virtual part origin (mark 2)
                bool sel2 = SelectComponentOriginForMate(swModel, partTransform.Component, 2);

                if (!sel2)
                {
                    Console.WriteLine($"Warning: Could not select virtual part origin for {partTransform.Component.Name2}");
                    swModel.ClearSelection2(true);
                    return;
                }

                // Add coincident mate with axis alignment
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
                            Console.WriteLine($"Enabled coincident mate axis alignment for {partTransform.Component.Name2}");
                        }
                    }

                    Console.WriteLine($"Updated mate to pose coordinate system for {partTransform.Component.Name2}");
                }
                else
                {
                    Console.WriteLine($"Note: Component positioned at pose coordinate system (mate skipped, error: {mateError})");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating mate for {partTransform.Component.Name2}: {ex.Message}");
                swModel.ClearSelection2(true);
            }
        }

        private static void CreatePoseCoordinateSystemFolder(ModelDoc2 swModel, List<VirtualPartTransformInfo> parts, string configName)
        {
            try
            {
                // Get the feature manager and create a folder
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Select all the pose coordinate systems to group them
                swModel.ClearSelection2(true);
                foreach (var partTransform in parts)
                {
                    Feature csFeat = FindCoordinateSystemFeature(swModel, partTransform.CoordSystemName);
                    if (csFeat != null)
                    {
                        csFeat.Select2(false, 0);
                    }
                }

                // Create feature folder
                string folderName = $"{configName} Coordinate Systems";
                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created {folderName} folder");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating pose coordinate system folder: {ex.Message}");
            }
        }

        private static Feature FindCoordinateSystemFeature(ModelDoc2 swModel, string csName)
        {
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.Name == csName)
                {
                    return feat;
                }
                feat = (Feature)feat.GetNextFeature();
            }
            return null;
        }

        private class HardpointInfo
        {
            public string BaseName { get; set; }
            public string Suffix { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
            // Rotation angles in degrees (used for wheels: camber/toe)
            public double AngleX { get; set; }
            public double AngleY { get; set; }
            public double AngleZ { get; set; }
        }

        private class BodyInfo
        {
            public Body2 Body { get; set; }
            public string OriginalName { get; set; }
            public string RenamedName { get; set; }
            public string BaseName { get; set; }
            public string Suffix { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
        }

        private class BodyTransformInfo
        {
            public Body2 Body { get; set; }
            public HardpointInfo Hardpoint { get; set; }
            public string TransformName { get; set; }
        }

        private class VirtualPartInfo
        {
            public Component2 Component { get; set; }
            public string OriginalName { get; set; }
            public string RenamedName { get; set; }
            public string BaseName { get; set; }
            public string Suffix { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
            // Rotation angles in degrees (used for wheels: camber/toe)
            public double AngleX { get; set; }
            public double AngleY { get; set; }
            public double AngleZ { get; set; }
        }

        private class VirtualPartTransformInfo
        {
            public Component2 Component { get; set; }
            public HardpointInfo Hardpoint { get; set; }
            public string CoordSystemName { get; set; }
        }

        // =====================================================================
        // Wheel insertion – separate from suspension hardpoints
        // Mirrors draw_suspension.py InsertWheel logic.
        // All virtual markers are placed at assembly origin (0,0,0);
        // toe/camber are stored in the part's internal coordinate system.
        // =====================================================================

        /// <summary>
        /// Inserts wheel marker virtual parts from a JSON suspension file.
        /// Usage: hardpoints addwheels <jsonPath> <markerPartPath> [rear] [referenceDistance]
        /// </summary>
        public static bool RunAddWheels(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: hardpoints addwheels <jsonPath> <markerPartPath> [rear] [referenceDistanceMM]");
                return false;
            }

            string jsonPath       = args[2];
            string markerPartPath = args[3];
            bool   isRear         = args.Length > 4 && args[4].ToLowerInvariant() == "rear";
            double refDistance    = 0.0;
            if (args.Length > 5) double.TryParse(args[5], out refDistance);

            ReportState(HardpointState.Initializing);

            SldWorks swApp;
            try { swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); }
            catch { Console.WriteLine("Error: SolidWorks is not running"); return false; }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null) { Console.WriteLine("Error: No active document"); return false; }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            { Console.WriteLine("Error: Active document must be an assembly"); return false; }

            AssemblyDoc swAssy = (AssemblyDoc)swModel;

            if (!File.Exists(markerPartPath))
            { Console.WriteLine($"Error: Marker part not found at {markerPartPath}"); return false; }

            try
            {
                ReportState(HardpointState.LoadingJson);
                var jsonData = LoadJsonData(jsonPath);

                // Extract wheels section
                if (!jsonData.TryGetValue("Wheels", out object wheelsObj) || !(wheelsObj is Dictionary<string, object> wheelsData))
                {
                    Console.WriteLine("Error: No 'Wheels' section found in JSON");
                    return false;
                }

                var wheelHardpoints = BuildWheelHardpoints(wheelsData, isRear, refDistance);
                if (wheelHardpoints.Count == 0)
                { Console.WriteLine("Error: Could not extract wheel data"); return false; }

                int totalSteps = wheelHardpoints.Count * 6 + 2;
                Console.WriteLine($"TOTAL:{totalSteps}");
                int progress = 0;

                // Load marker part
                ReportState(HardpointState.LoadingMarkerPart);
                ModelDoc2 markerDoc = LoadMarkerPart(swApp, markerPartPath);
                if (markerDoc == null) return false;

                int activateErrors = 0;
                swApp.ActivateDoc2(swModel.GetTitle(), false, ref activateErrors);
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swAssy  = (AssemblyDoc)swModel;

                swModel.ClearSelection2(true);
                swAssy.EditAssembly();

                var insertedParts = new List<VirtualPartInfo>();

                ReportState(HardpointState.InsertingBodies);
                foreach (var hp in wheelHardpoints)
                {
                    // 1. Insert at origin
                    var partInfo = InsertVirtualMarkerPart(swAssy, markerDoc, hp);
                    if (partInfo == null) continue;
                    progress++; ReportProgress(progress);

                    // 2. Make virtual
                    MakeComponentVirtual(partInfo);
                    progress++; ReportProgress(progress);

                    // 3. Rename component
                    RenameVirtualPart(swAssy, partInfo);
                    progress++; ReportProgress(progress);

                    // 4. Float the component in the assembly (unfix)
                    FloatComponent(swAssy, swModel, partInfo);
                    progress++; ReportProgress(progress);

                    // 5. Rename internal CS AND apply wheel angles (camber / toe) via EditPart2
                    RenameAndOrientWheelCoordinateSystem(swApp, swAssy, swModel, partInfo);
                    progress++; ReportProgress(progress);

                    // 6. Apply green colour for wheels
                    ApplyColorToVirtualPart(swAssy, partInfo);
                    progress++; ReportProgress(progress);

                    insertedParts.Add(partInfo);
                }

                // Folder + move all components in one reorder call
                ReportState(HardpointState.CreatingHardpointsFolder);
                if (insertedParts.Count > 0)
                {
                    CreateContainingFolderFromComponents(swModel, "Wheel Hardpoints", insertedParts);
                }
                progress++; ReportProgress(progress);

                swModel.EditRebuild3();
                progress++; ReportProgress(progress);

                ReportState(HardpointState.Complete);
                Console.WriteLine($"Successfully inserted {insertedParts.Count} wheel marker parts");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }

        private static Feature EnsureAssemblyFolder(AssemblyDoc swAssy, ModelDoc2 swModel, string folderName, Component2 anchorComponent)
        {
            try
            {
                if (swAssy == null || swModel == null)
                {
                    Console.WriteLine($"Cannot create folder '{folderName}': invalid assembly/model document");
                    return null;
                }

                Feature existing = (Feature)swAssy.FeatureByName(folderName);
                if (existing != null)
                {
                    Console.WriteLine($"Using existing folder '{folderName}'");
                    return existing;
                }

                if (anchorComponent == null)
                {
                    Console.WriteLine($"Cannot create folder '{folderName}': anchor component is null");
                    return null;
                }

                FeatureManager featMgr = swModel.FeatureManager;
                if (featMgr == null)
                {
                    Console.WriteLine($"Cannot create folder '{folderName}': FeatureManager is null");
                    return null;
                }

                swModel.ClearSelection2(true);
                bool selected = anchorComponent.Select4(false, null, false);
                Console.WriteLine($"Anchor select for folder '{folderName}': {selected}");
                if (!selected)
                {
                    return null;
                }

                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore);
                swModel.ClearSelection2(true);

                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created folder '{folderName}'");
                    return folder;
                }

                Console.WriteLine($"InsertFeatureTreeFolder2 returned null for '{folderName}'");
                return (Feature)swAssy.FeatureByName(folderName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating folder '{folderName}': {ex.Message}");
                return null;
            }
        }

        private static Feature CreateContainingFolderFromComponents(ModelDoc2 swModel, string folderName, List<VirtualPartInfo> parts)
        {
            try
            {
                if (swModel == null || parts == null || parts.Count == 0)
                {
                    Console.WriteLine($"Cannot create folder '{folderName}': no model or components");
                    return null;
                }

                FeatureManager featMgr = swModel.FeatureManager;
                if (featMgr == null)
                {
                    Console.WriteLine($"Cannot create folder '{folderName}': FeatureManager is null");
                    return null;
                }

                swModel.ClearSelection2(true);

                int selectedCount = 0;
                foreach (var part in parts)
                {
                    if (part?.Component == null)
                    {
                        continue;
                    }

                    bool selected = part.Component.Select4(true, null, false);
                    if (selected)
                    {
                        selectedCount++;
                    }
                }

                Console.WriteLine($"Selected {selectedCount} components for containing folder '{folderName}'");

                if (selectedCount == 0)
                {
                    swModel.ClearSelection2(true);
                    return null;
                }

                Feature folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                swModel.ClearSelection2(true);

                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created containing folder '{folderName}' with {selectedCount} components");
                    return folder;
                }

                Console.WriteLine($"InsertFeatureTreeFolder2(Containing) returned null for '{folderName}'");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating containing folder '{folderName}': {ex.Message}");
                return null;
            }
        }

        private static void MoveComponentsToFolder(AssemblyDoc swAssy, List<VirtualPartInfo> parts, Feature folder, string folderName)
        {
            try
            {
                if (swAssy == null || parts == null || folder == null)
                {
                    return;
                }

                var componentObjects = new List<object>();
                foreach (var part in parts)
                {
                    if (part?.Component != null)
                    {
                        componentObjects.Add(part.Component);
                    }
                }

                if (componentObjects.Count == 0)
                {
                    Console.WriteLine($"No components found to move into '{folderName}'");
                    return;
                }

                object[] components = componentObjects.ToArray();
                bool moved = swAssy.ReorderComponents(
                    components,
                    folder,
                    (int)swReorderComponentsWhere_e.swReorderComponents_LastInFolder);

                if (moved)
                {
                    Console.WriteLine($"Moved {components.Length} components into '{folderName}'");
                }
                else
                {
                    Console.WriteLine($"Warning: Failed to move components into '{folderName}'");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error moving components into '{folderName}': {ex.Message}");
            }
        }

        /// <summary>
        /// Builds a list of HardpointInfo entries for left and right wheels from the
        /// Wheels JSON section, mirroring draw_suspension.py InsertWheel logic.
        /// Positions are stored (for pose step) but parts are inserted at origin.
        /// </summary>
        private static List<HardpointInfo> BuildWheelHardpoints(
            Dictionary<string, object> wheelsData,
            bool isRear,
            double referenceDistanceMM)
        {
            var list = new List<HardpointInfo>();
            try
            {
                double halfTrack        = GetWheelValue(wheelsData, "Half Track",           "left");
                double tireDiameter     = GetWheelValue(wheelsData, "Tire Diameter",        "left");
                double lateralOffset    = GetWheelValue(wheelsData, "Lateral Offset",       "left", 0);
                double verticalOffset   = GetWheelValue(wheelsData, "Vertical Offset",      "left", 0);
                double longOffset       = GetWheelValue(wheelsData, "Longitudinal Offset",  "left", 0);
                double camber           = GetWheelValue(wheelsData, "Static Camber",        "left", 0);
                double toe              = GetWheelValue(wheelsData, "Static Toe",           "left", 0);

                double xBase = referenceDistanceMM + longOffset;
                double yBase = halfTrack + lateralOffset;
                double z     = tireDiameter / 2.0 + verticalOffset;

                string prefix = isRear ? "R" : "F";
                string suffix = isRear ? "_REAR" : "_FRONT";

                // Left wheel (positive Y, positive camber/toe)
                list.Add(new HardpointInfo
                {
                    BaseName = $"{prefix}L_wheel",
                    Suffix   = suffix,
                    X        = xBase,
                    Y        = yBase,
                    Z        = z,
                    AngleX   = camber,
                    AngleY   = 0,
                    AngleZ   = toe
                });

                // Right wheel (negative Y, mirrored camber/toe)
                list.Add(new HardpointInfo
                {
                    BaseName = $"{prefix}R_wheel",
                    Suffix   = suffix,
                    X        = xBase,
                    Y        = -yBase,
                    Z        = z,
                    AngleX   = -camber,
                    AngleY   = 0,
                    AngleZ   = -toe
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error building wheel hardpoints: {ex.Message}");
            }
            return list;
        }

        private static double GetWheelValue(
            Dictionary<string, object> wheelsData,
            string key, string side,
            double fallback = 0.0)
        {
            if (wheelsData.TryGetValue(key, out object sectionObj) &&
                sectionObj is Dictionary<string, object> section &&
                section.TryGetValue(side, out object val))
            {
                return Convert.ToDouble(val);
            }
            return fallback;
        }

        private static bool IsWheelHardpoint(string baseName)
        {
            return !string.IsNullOrWhiteSpace(baseName) &&
                   baseName.IndexOf("wheel", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Reads Vehicle_Setup.json and returns rear offset in mm as
        /// -Reference distance (matching draw_suspension.py behavior).
        /// </summary>
        private static double LoadRearReferenceOffsetMm(string jsonDirectory)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jsonDirectory))
                {
                    return 0.0;
                }

                string vehicleSetupPath = Path.Combine(jsonDirectory, "Vehicle_Setup.json");
                if (!File.Exists(vehicleSetupPath))
                {
                    return 0.0;
                }

                var vehicleData = LoadJsonData(vehicleSetupPath);
                if (vehicleData == null)
                {
                    return 0.0;
                }

                if (vehicleData.TryGetValue("Reference distance", out object refDistanceObj))
                {
                    return -Convert.ToDouble(refDistanceObj);
                }

                return 0.0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read rear reference offset: {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Uses EditPart2 to enter the virtual wheel part, rename its CoordSys,
        /// and update its rotation to reflect toe/camber angles.
        /// </summary>
        private static void RenameAndOrientWheelCoordinateSystem(
            SldWorks swApp, AssemblyDoc swAssy, ModelDoc2 swModel, VirtualPartInfo partInfo)
        {
            string newCsName = string.IsNullOrWhiteSpace(partInfo.RenamedName)
                ? $"{partInfo.BaseName}{partInfo.Suffix}"
                : partInfo.RenamedName;
            try
            {
                swModel.ClearSelection2(true);
                partInfo.Component.Select4(false, null, false);

                int editInfo   = 0;
                int editResult = swAssy.EditPart2(true, false, ref editInfo);

                if (editResult == (int)swEditPartCommandStatus_e.swEditPartSuccessful)
                {
                    ModelDoc2 partDoc = (ModelDoc2)partInfo.Component.GetModelDoc2();
                    if (partDoc == null)
                    {
                        partDoc = (ModelDoc2)swApp.ActiveDoc;
                    }

                    if (partDoc != null)
                    {
                        Feature targetCs = GetInternalCoordinateSystemFeature(partDoc);

                        if (targetCs != null && targetCs.GetTypeName2() == "CoordSys")
                        {
                            // Rename to exactly hardpoint name
                            targetCs.Name = newCsName;

                            // Apply wheel orientation via ModifyDefinition if angles present
                            if (partInfo.AngleX != 0 || partInfo.AngleY != 0 || partInfo.AngleZ != 0)
                            {
                                CoordinateSystemFeatureData csData =
                                    (CoordinateSystemFeatureData)targetCs.GetDefinition();
                                if (csData != null && csData.AccessSelections(partDoc, null))
                                {
                                    // Leave origin at part origin (null entity)
                                    csData.OriginEntity = null;
                                    bool ok = targetCs.ModifyDefinition(csData, partDoc, null);
                                    csData.ReleaseSelectionAccess();
                                    Console.WriteLine(ok
                                        ? $"Renamed/oriented wheel CS '{newCsName}' (camber={partInfo.AngleX}, toe={partInfo.AngleZ})"
                                        : $"Warning: ModifyDefinition failed for wheel CS '{newCsName}'");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Renamed wheel CS to '{newCsName}'");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Could not find wheel coordinate system to rename for {partInfo.RenamedName}");
                        }

                        partDoc.EditRebuild3();
                    }
                    swAssy.EditAssembly();
                }
                else
                {
                    Console.WriteLine($"Warning: EditPart2 returned {editResult} for {partInfo.RenamedName}");
                }

                swModel.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error orienting wheel CS for {partInfo.RenamedName}: {ex.Message}");
                try { swAssy.EditAssembly(); } catch { }
            }
        }

        private static Feature GetInternalCoordinateSystemFeature(ModelDoc2 partDoc)
        {
            if (partDoc == null)
            {
                return null;
            }

            try
            {
                // Best case: default marker coordinate system name
                PartDoc partDocAsPart = partDoc as PartDoc;
                Feature byDefaultName = partDocAsPart != null
                    ? (Feature)partDocAsPart.FeatureByName("Coordinate System1")
                    : null;
                if (byDefaultName != null && IsCoordinateSystemFeature(byDefaultName))
                {
                    return byDefaultName;
                }
            }
            catch { }

            try
            {
                // Selection-based fallback by name
                bool selected = partDoc.Extension.SelectByID2(
                    "Coordinate System1",
                    "COORDSYS",
                    0, 0, 0,
                    false,
                    0,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (selected)
                {
                    SelectionMgr selMgr = (SelectionMgr)partDoc.SelectionManager;
                    Feature selectedFeature = (Feature)selMgr.GetSelectedObject6(1, -1);
                    partDoc.ClearSelection2(true);
                    if (selectedFeature != null && IsCoordinateSystemFeature(selectedFeature))
                    {
                        return selectedFeature;
                    }
                }
            }
            catch
            {
                try { partDoc.ClearSelection2(true); } catch { }
            }

            // Last fallback: traverse full feature tree including subfeatures
            return FindFirstCoordinateSystemFeature(partDoc);
        }

        private static Feature FindFirstCoordinateSystemFeature(ModelDoc2 partDoc)
        {
            Feature feat = (Feature)partDoc.FirstFeature();
            while (feat != null)
            {
                if (IsCoordinateSystemFeature(feat))
                {
                    return feat;
                }

                Feature inSub = FindCoordinateSystemInSubFeatures((Feature)feat.GetFirstSubFeature());
                if (inSub != null)
                {
                    return inSub;
                }

                feat = (Feature)feat.GetNextFeature();
            }

            return null;
        }

        private static Feature FindCoordinateSystemInSubFeatures(Feature subFeature)
        {
            Feature current = subFeature;
            while (current != null)
            {
                if (IsCoordinateSystemFeature(current))
                {
                    return current;
                }

                Feature nested = FindCoordinateSystemInSubFeatures((Feature)current.GetFirstSubFeature());
                if (nested != null)
                {
                    return nested;
                }

                current = (Feature)current.GetNextSubFeature();
            }

            return null;
        }

        private static bool IsCoordinateSystemFeature(Feature feature)
        {
            if (feature == null)
            {
                return false;
            }

            string typeName = string.Empty;
            try
            {
                typeName = feature.GetTypeName2();
            }
            catch
            {
                return false;
            }

            return string.Equals(typeName, "CoordSys", StringComparison.OrdinalIgnoreCase);
        }
    }
}
