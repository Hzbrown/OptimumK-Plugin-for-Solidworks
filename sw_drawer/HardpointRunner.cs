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
                
                var hardpoints = new List<HardpointInfo>();
                
                // Load front suspension
                if (File.Exists(frontJsonPath))
                {
                    var frontData = LoadJsonData(frontJsonPath);
                    ExtractHardpointsWithSuffix(frontData, "_FRONT", hardpoints);
                    Console.WriteLine($"Loaded {hardpoints.Count} hardpoints from Front_Suspension.json");
                }
                
                // Load rear suspension
                int frontCount = hardpoints.Count;
                if (File.Exists(rearJsonPath))
                {
                    var rearData = LoadJsonData(rearJsonPath);
                    ExtractHardpointsWithSuffix(rearData, "_REAR", hardpoints);
                    Console.WriteLine($"Loaded {hardpoints.Count - frontCount} hardpoints from Rear_Suspension.json");
                }
                
                if (hardpoints.Count == 0)
                {
                    Console.WriteLine("Error: No hardpoints found in JSON files");
                    return false;
                }
                
                // Precompute total steps (insert + make virtual + rename + rename CS + color + folder + rebuild)
                int totalSteps = hardpoints.Count * 5 + 2;
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
                        
                        // Step 4: Rename internal coordinate system
                        RenameInternalCoordinateSystem(swModel, partInfo);
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        // Step 5: Apply color
                        ApplyColorToVirtualPart(swAssy, partInfo);
                        progressCount++;
                        ReportProgress(progressCount);
                        
                        insertedParts.Add(partInfo);
                    }
                }

                // Step 6: Create Hardpoints folder
                ReportState(HardpointState.CreatingHardpointsFolder);
                CreateHardpointsFolder(swAssy, insertedParts);
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
                // Convert mm to meters for SolidWorks
                double x = hardpoint.X / 1000.0;
                double y = hardpoint.Y / 1000.0;
                double z = hardpoint.Z / 1000.0;

                // Use InsertPart2 to insert the marker as a virtual part
                Component2 newComp = (Component2)swAssy.AddComponent4(markerDoc.GetPathName(), "", x, y, z);
                
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
                    X = x,
                    Y = y,
                    Z = z
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
            // Extract from all sections except Wheels
            foreach (var section in jsonData)
            {
                if (section.Key == "Wheels")
                {
                    continue; // Wheels handled separately if needed
                }
                else
                {
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
                                    X = Convert.ToDouble(coords[0]),
                                    Y = Convert.ToDouble(coords[1]),
                                    Z = Convert.ToDouble(coords[2])
                                });
                            }
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

        private static void RenameInternalCoordinateSystem(ModelDoc2 swModel, VirtualPartInfo partInfo)
        {
            try
            {
                // The internal coordinate system in Marker.SLDPRT is called "Coordinate System1"
                // We need to rename it in the context of this virtual component
                string oldCsName = "Coordinate System1@" + partInfo.RenamedName;
                string newCsName = $"CS_{partInfo.RenamedName}";
                
                // Try to select and rename the coordinate system feature
                bool selected = swModel.Extension.SelectByID2(
                    oldCsName, "COORDSYS", 0, 0, 0, false, 0, null,
                    (int)swSelectOption_e.swSelectOptionDefault);
                
                if (selected)
                {
                    SelectionMgr selMgr = (SelectionMgr)swModel.SelectionManager;
                    Feature feat = (Feature)selMgr.GetSelectedObject6(1, -1);
                    if (feat != null)
                    {
                        feat.Name = newCsName;
                        Console.WriteLine($"Renamed internal coordinate system to: {newCsName}");
                    }
                    swModel.ClearSelection2(true);
                }
                else
                {
                    Console.WriteLine($"Warning: Could not find internal coordinate system for {partInfo.RenamedName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming internal coordinate system: {ex.Message}");
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
                bool sel2 = swModel.Extension.SelectByID2(
                    "Origin@" + partInfo.RenamedName + "@" + swModel.GetTitle(),
                    "ORIGINFOLDER", 0, 0, 0, true, 2, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

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

        private static void CreateHardpointsFolder(AssemblyDoc swAssy, List<VirtualPartInfo> parts)
        {
            try
            {
                // Get the feature manager and create a folder
                ModelDoc2 swModel = (ModelDoc2)swAssy;
                FeatureManager featMgr = swModel.FeatureManager;
                
                // Select all the virtual parts to group them using SelectByID2
                swModel.ClearSelection2(true);
                int selectedCount = 0;
                foreach (var partInfo in parts)
                {
                    // Select component by name in assembly
                    string compName = partInfo.RenamedName + "@" + swModel.GetTitle();
                    bool selected = swModel.Extension.SelectByID2(
                        compName, "COMPONENT", 0, 0, 0, true, 0, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                    
                    if (selected)
                    {
                        selectedCount++;
                    }
                }

                Console.WriteLine($"Selected {selectedCount} components for folder");

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
                bool sel2 = swModel.Extension.SelectByID2(
                    "Origin@" + partTransform.Component.Name2 + "@" + swModel.GetTitle(),
                    "ORIGINFOLDER", 0, 0, 0, true, 2, null,
                    (int)swSelectOption_e.swSelectOptionDefault);

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
        }

        private class VirtualPartTransformInfo
        {
            public Component2 Component { get; set; }
            public HardpointInfo Hardpoint { get; set; }
            public string CoordSystemName { get; set; }
        }
    }
}