using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public class InsertMarker
    {
        // Predefined colors for different groups (RGB values 0-255)
        // Matches the actual OptimumK naming conventions
        private static readonly int[,] GroupColors = new int[,]
        {
            { 255, 0, 0 },      // Red - Chassis (CHAS)
            { 0, 255, 0 },      // Green - Upright (UPRI)
            { 0, 0, 255 },      // Blue - Rocker (ROCK)
            { 255, 255, 0 },    // Yellow - Steering/Tierod (STEE/TIE)
            { 255, 0, 255 },    // Magenta - Wheel
            { 0, 255, 255 },    // Cyan - Damper (DAMP)
            { 255, 128, 0 },    // Orange - ARB
            { 128, 0, 255 },    // Purple - Pushrod (PUSH)
            { 128, 128, 128 }   // Gray - Other
        };

        // Group names matching actual OptimumK JSON naming
        private static readonly string[] GroupNames = new string[]
        {
            "Chassis",      // _Chassis suffix
            "Upright",      // _Upright suffix  
            "Rocker",       // _Rocker suffix
            "Steering",     // Tierod_, Steering
            "Wheel",        // _wheel suffix
            "Damper",       // Damper_
            "ARB",          // ARB_
            "Pushrod",      // Pushrod_
            "Other"         // Anything else
        };

        public static int GetGroupIndex(string featureName)
        {
            string upper = featureName.ToUpperInvariant();
            
            // Match based on actual OptimumK naming conventions from JSON
            // Chassis points (suffix _Chassis or _CHAS)
            if (upper.Contains("_CHASSIS") || upper.Contains("_CHAS") || upper.Contains("CHASSIS_"))
                return 0;
            
            // Upright points (suffix _Upright or _UPRI)
            if (upper.Contains("_UPRIGHT") || upper.Contains("_UPRI") || upper.Contains("UPRIGHT_"))
                return 1;
            
            // Rocker points (suffix _Rocker or _ROCK)
            if (upper.Contains("_ROCKER") || upper.Contains("_ROCK") || upper.Contains("ROCKER_"))
                return 2;
            
            // Steering/Tierod points
            if (upper.Contains("TIEROD") || upper.Contains("TIE_ROD") || upper.Contains("STEERING") || upper.Contains("_STEE"))
                return 3;
            
            // Wheel coordinate systems
            if (upper.Contains("_WHEEL") || upper.Contains("WHEEL_"))
                return 4;
            
            // Damper points
            if (upper.Contains("DAMPER") || upper.Contains("_DAMP") || upper.Contains("SHOCK"))
                return 5;
            
            // ARB points
            if (upper.Contains("ARB") || upper.Contains("ANTIROLL") || upper.Contains("SWAY"))
                return 6;
            
            // Pushrod points
            if (upper.Contains("PUSHROD") || upper.Contains("_PUSH") || upper.Contains("PUSH_"))
                return 7;
            
            return 8; // Other
        }

        /// <summary>
        /// Gets a feature by name from the model by iterating through all features.
        /// </summary>
        private static Feature GetFeatureByName(ModelDoc2 swModel, string featureName)
        {
            Feature swFeature = (Feature)swModel.FirstFeature();
            while (swFeature != null)
            {
                if (swFeature.Name == featureName)
                {
                    return swFeature;
                }
                swFeature = (Feature)swFeature.GetNextFeature();
            }
            return null;
        }

        /// <summary>
        /// Gets the path to the Marker.STEP file.
        /// </summary>
        private static string GetMarkerStepPath()
        {
            // Look for Marker.STEP in the sw_drawer directory or parent directory
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] searchPaths = new string[]
            {
                Path.Combine(exeDir, "Marker.STEP"),
                Path.Combine(exeDir, "..", "..", "..", "..", "Marker.STEP"),
                Path.Combine(exeDir, "..", "..", "..", "..", "..", "Marker.STEP"),
                Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), "Marker.STEP")
            };

            foreach (string path in searchPaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            // Default location
            return Path.Combine(Path.GetDirectoryName(exeDir.TrimEnd('\\')), "Marker.STEP");
        }

        public static bool Run(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SuspensionTools.exe marker <create|vis> <args...>");
                Console.WriteLine("  marker create <name> <radius_mm>");
                Console.WriteLine("  marker createall <radius_mm>");
                Console.WriteLine("  marker deleteall");
                Console.WriteLine("  marker vis <all|group|name|front|rear> <show|hide> [param]");
                return false;
            }

            string subCommand = args[1].ToLower();

            if (subCommand == "create" && args.Length >= 4)
            {
                string csName = args[2];
                double scale = double.Parse(args[3]) / 5.0; // Scale factor based on 5mm default
                return CreateMarkerAtCoordSystem(csName, scale);
            }
            else if (subCommand == "vis" && args.Length >= 4)
            {
                string target = args[2].ToLower();
                bool visible = args[3].ToLower() == "show";
                string param = args.Length > 4 ? args[4] : null;
                return SetMarkerVisibility(target, visible, param);
            }
            else if (subCommand == "createall" && args.Length >= 3)
            {
                double scale = double.Parse(args[2]) / 5.0;
                return CreateAllMarkers(scale);
            }
            else if (subCommand == "deleteall")
            {
                return DeleteAllMarkers();
            }

            Console.WriteLine("Invalid marker command");
            return false;
        }

        private static bool CreateMarkerAtCoordSystem(string csName, double scale)
        {
            SldWorks swApp = null;
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

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                Console.WriteLine("Error: Active document must be a part");
                return false;
            }

            PartDoc swPart = (PartDoc)swModel;
            
            // Find the coordinate system
            Feature csFeature = GetFeatureByName(swModel, csName);
            if (csFeature == null)
            {
                Console.WriteLine($"Error: Coordinate system '{csName}' not found");
                return false;
            }

            // Get coordinate system transform
            CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)csFeature.GetDefinition();
            MathTransform csTransform = csData.Transform;
            double[] transformData = (double[])csTransform.ArrayData;
            
            // Extract origin (positions are at indices 9, 10, 11)
            double x = transformData[9];
            double y = transformData[10];
            double z = transformData[11];

            string markerName = csName + "_Marker";
            
            // Check if marker already exists
            Feature existingMarker = GetFeatureByName(swModel, markerName);
            if (existingMarker != null)
            {
                Console.WriteLine($"Marker '{markerName}' already exists, skipping");
                return true;
            }

            // Import STEP file as marker
            bool success = ImportMarkerStep(swApp, swModel, swPart, markerName, x, y, z, scale, csName);
            
            if (success)
            {
                MoveToMarkersFolder(swModel, markerName);
                Console.WriteLine($"Created marker: {markerName} at ({x * 1000:F1}, {y * 1000:F1}, {z * 1000:F1}) mm");
            }
            
            return success;
        }

        private static bool ImportMarkerStep(SldWorks swApp, ModelDoc2 swModel, PartDoc swPart,
            string markerName, double x, double y, double z, double scale, string csName)
        {
            try
            {
                // Create sphere body using IModeler (matching VBA example from codestack.net)
                Modeler modeler = (Modeler)swApp.GetModeler();
                if (modeler == null)
                {
                    Console.WriteLine($"Error: Could not get Modeler for {markerName}");
                    return CreateSketchPointMarker(swModel, markerName, x, y, z, csName);
                }

                // Radius based on scale (5mm default * scale factor)
                double radius = 0.005 * scale; // 5mm in meters * scale
                
                Console.WriteLine($"Creating sphere at ({x*1000:F1}, {y*1000:F1}, {z*1000:F1}) mm, radius={radius*1000:F1}mm");
                
                // Create sphere center point and axis arrays (matching VBA)
                double[] dCenter = new double[] { x, y, z };
                double[] dAxis = new double[] { 0, 0, 1 };
                double[] dRef = new double[] { 1, 0, 0 };

                // Create spherical surface
                Surface swSurf = (Surface)modeler.CreateSphericalSurface2(dCenter, dAxis, dRef, radius);
                
                if (swSurf == null)
                {
                    Console.WriteLine($"Error: CreateSphericalSurface2 returned null for {markerName}");
                    return CreateSketchPointMarker(swModel, markerName, x, y, z, csName);
                }
                
                Console.WriteLine($"Spherical surface created for {markerName}");

                // Create full sphere - pass null for UV bounds to get full sphere
                // In VBA this is "Empty", in C# we pass null
                Body2 swBody = (Body2)swSurf.CreateTrimmedSheet4(null, true);
                
                if (swBody == null)
                {
                    Console.WriteLine($"Error: CreateTrimmedSheet4 returned null for {markerName}");
                    
                    // Try alternative: create with explicit UV range for full sphere
                    // U: 0 to 2*PI, V: -PI/2 to PI/2
                    double[] uvRange = new double[] { 0, 2 * Math.PI, -Math.PI / 2, Math.PI / 2 };
                    swBody = (Body2)swSurf.CreateTrimmedSheet4(uvRange, true);
                    
                    if (swBody == null)
                    {
                        Console.WriteLine($"Error: CreateTrimmedSheet4 with UV range also failed for {markerName}");
                        return CreateSketchPointMarker(swModel, markerName, x, y, z, csName);
                    }
                }
                
                Console.WriteLine($"Sheet body created for {markerName}, type={swBody.GetType()}");

                // Set color on body BEFORE adding to part
                SetMarkerColorOnBody(swBody, csName);

                // Add the body to the part as a feature
                Feature feat = (Feature)swPart.CreateFeatureFromBody3(
                    swBody, 
                    false, 
                    (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodyCheck
                );

                if (feat != null)
                {
                    feat.Name = markerName;
                    Console.WriteLine($"Feature created: {markerName}");
                    return true;
                }
                else
                {
                    // Try with different options
                    Console.WriteLine($"CreateFeatureFromBody3 with Check failed, trying Simplify...");
                    feat = (Feature)swPart.CreateFeatureFromBody3(
                        swBody, 
                        false, 
                        (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify
                    );
                    
                    if (feat != null)
                    {
                        feat.Name = markerName;
                        Console.WriteLine($"Feature created with Simplify: {markerName}");
                        return true;
                    }
                    
                    Console.WriteLine($"Error: CreateFeatureFromBody3 failed for {markerName}");
                    return CreateSketchPointMarker(swModel, markerName, x, y, z, csName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception creating marker {markerName}: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return CreateSketchPointMarker(swModel, markerName, x, y, z, csName);
            }
        }

        /// <summary>
        /// Sets color directly on a body object.
        /// </summary>
        private static void SetMarkerColorOnBody(Body2 body, string csName)
        {
            if (body == null) return;
            
            int groupIdx = GetGroupIndex(csName);
            int red = GroupColors[groupIdx, 0];
            int green = GroupColors[groupIdx, 1];
            int blue = GroupColors[groupIdx, 2];
            
            try
            {
                double[] matProps = new double[9];
                matProps[0] = red / 255.0;
                matProps[1] = green / 255.0;
                matProps[2] = blue / 255.0;
                matProps[3] = 1.0;   // Ambient
                matProps[4] = 1.0;   // Diffuse
                matProps[5] = 0.2;   // Specular
                matProps[6] = 0.3;   // Shininess
                matProps[7] = 0.0;   // Transparency
                matProps[8] = 0.0;   // Emission
                
                body.MaterialPropertyValues2 = matProps;
            }
            catch { }
        }

        private static bool CreateSketchPointMarker(ModelDoc2 swModel, string markerName, double x, double y, double z, string csName)
        {
            // Fallback: create a 3D sketch with a point as the marker
            swModel.ClearSelection2(true);
            swModel.SketchManager.Insert3DSketch(true);
            swModel.SketchManager.CreatePoint(x, y, z);
            swModel.SketchManager.Insert3DSketch(true);
            
            Feature sketchFeat = (Feature)swModel.Extension.GetLastFeatureAdded();
            if (sketchFeat != null)
            {
                sketchFeat.Name = markerName;
                MoveToMarkersFolder(swModel, markerName);
                Console.WriteLine($"Created sketch point marker: {markerName}");
                return true;
            }
            return false;
        }

        private static int Get3DSketchCount(ModelDoc2 swModel)
        {
            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "3DProfileFeature")
                    count++;
                feat = (Feature)feat.GetNextFeature();
            }
            return count;
        }

        private static void EnsureMarkersFolder(ModelDoc2 swModel)
        {
            Feature folder = GetFeatureByName(swModel, "Markers");
            if (folder == null)
            {
                swModel.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore);
                Feature newFolder = (Feature)swModel.Extension.GetLastFeatureAdded();
                if (newFolder != null)
                {
                    newFolder.Name = "Markers";
                }
            }
        }

        /// <summary>
        /// Moves a feature into the "Markers" folder, creating the folder if it doesn't exist.
        /// </summary>
        private static void MoveToMarkersFolder(ModelDoc2 swModel, string featureName)
        {
            const string folderName = "Markers";
            FeatureManager featMgr = swModel.FeatureManager;

            Feature feat = GetFeatureByName(swModel, featureName);
            if (feat == null) return;

            // Try to find existing "Markers" folder
            Feature folder = null;
            Feature swFeat = (Feature)swModel.FirstFeature();
            
            while (swFeat != null)
            {
                if (swFeat.GetTypeName2() == "FtrFolder" && swFeat.Name == folderName)
                {
                    folder = swFeat;
                    break;
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }

            // Create folder if it doesn't exist
            if (folder == null)
            {
                feat.Select2(false, 0);
                folder = featMgr.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                
                if (folder != null)
                {
                    folder.Name = folderName;
                }
                swModel.ClearSelection2(true);
            }
            else
            {
                feat.Select2(false, 0);
                swModel.Extension.ReorderFeature(feat.Name, folder.Name, (int)swMoveLocation_e.swMoveToFolder);
                swModel.ClearSelection2(true);
            }
        }

        private static void SetMarkerColor(ModelDoc2 swModel, PartDoc swPart, string markerName, string csName, Body2 body = null)
        {
            int groupIdx = GetGroupIndex(csName);
            int red = GroupColors[groupIdx, 0];
            int green = GroupColors[groupIdx, 1];
            int blue = GroupColors[groupIdx, 2];
            
            // Find body by feature name
            object[] bodies = (object[])swPart.GetBodies2((int)swBodyType_e.swAllBodies, true);
            if (bodies != null)
            {
                foreach (Body2 bodyItem in bodies)
                {
                    if (bodyItem != null && bodyItem.Name != null && bodyItem.Name.Contains(markerName))
                    {
                        try
                        {
                            double[] matProps = new double[9];
                            matProps[0] = red / 255.0;
                            matProps[1] = green / 255.0;
                            matProps[2] = blue / 255.0;
                            matProps[3] = 1.0;
                            matProps[4] = 1.0;
                            matProps[5] = 0.2;
                            matProps[6] = 0.3;
                            matProps[7] = 0.0;
                            matProps[8] = 0.0;
                            
                            bodyItem.MaterialPropertyValues2 = matProps;
                            break;
                        }
                        catch { }
                    }
                }
            }
        }

        private static bool CreateAllMarkers(double scale)
        {
            SldWorks swApp = null;
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

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                Console.WriteLine("Error: Active document must be a part");
                return false;
            }

            PartDoc swPart = (PartDoc)swModel;
            
            // Collect all coordinate system features first
            System.Collections.Generic.List<Feature> coordSystems = new System.Collections.Generic.List<Feature>();
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "CoordSys" && feat.Name != "Origin")
                {
                    coordSystems.Add(feat);
                }
                feat = (Feature)feat.GetNextFeature();
            }

            int count = 0;
            foreach (Feature csFeat in coordSystems)
            {
                string csName = csFeat.Name;
                Feature existingMarker = GetFeatureByName(swModel, csName + "_Marker");
                if (existingMarker == null)
                {
                    if (CreateMarkerFromFeature(swApp, swModel, swPart, csFeat, scale))
                        count++;
                }
            }

            Console.WriteLine($"Created {count} markers");
            swModel.EditRebuild3();
            return true;
        }

        private static bool CreateMarkerFromFeature(SldWorks swApp, ModelDoc2 swModel, PartDoc swPart, Feature csFeature, double scale)
        {
            string csName = csFeature.Name;
            string markerName = csName + "_Marker";

            try
            {
                CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)csFeature.GetDefinition();
                MathTransform csTransform = csData.Transform;
                double[] transformData = (double[])csTransform.ArrayData;

                double x = transformData[9];
                double y = transformData[10];
                double z = transformData[11];

                bool success = ImportMarkerStep(swApp, swModel, swPart, markerName, x, y, z, scale, csName);
                
                if (success)
                {
                    MoveToMarkersFolder(swModel, markerName);
                }
                
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating marker {markerName}: {ex.Message}");
                return false;
            }
        }

        private static bool DeleteAllMarkers()
        {
            SldWorks swApp = null;
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

            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            System.Collections.Generic.List<string> toDelete = new System.Collections.Generic.List<string>();
            
            while (feat != null)
            {
                if (feat.Name.EndsWith("_Marker") || feat.Name.EndsWith("_Marker_Sketch"))
                {
                    toDelete.Add(feat.Name);
                }
                feat = (Feature)feat.GetNextFeature();
            }

            foreach (string name in toDelete)
            {
                Feature f = GetFeatureByName(swModel, name);
                if (f != null)
                {
                    swModel.ClearSelection2(true);
                    f.Select2(false, 0);
                    swModel.EditDelete();
                    count++;
                }
            }

            Console.WriteLine($"Deleted {count} markers");
            return true;
        }

        private static bool SetMarkerVisibility(string target, bool visible, string param)
        {
            SldWorks swApp = null;
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

            int count = 0;
            Feature feat = (Feature)swModel.FirstFeature();
            
            while (feat != null)
            {
                bool shouldModify = false;
                string featName = feat.Name;
                
                if (featName.EndsWith("_Marker") || featName.EndsWith("_Marker_Sketch"))
                {
                    string baseName = featName.Replace("_Marker_Sketch", "").Replace("_Marker", "");
                    
                    switch (target)
                    {
                        case "all":
                            shouldModify = true;
                            break;
                        case "group":
                            if (param != null)
                            {
                                int targetGroup = Array.FindIndex(GroupNames, g => 
                                    g.Equals(param, StringComparison.OrdinalIgnoreCase));
                                shouldModify = (targetGroup >= 0 && GetGroupIndex(baseName) == targetGroup);
                            }
                            break;
                        case "name":
                            shouldModify = param != null && baseName.IndexOf(param, StringComparison.OrdinalIgnoreCase) >= 0;
                            break;
                        case "front":
                            shouldModify = baseName.ToUpperInvariant().Contains("_FRONT") || 
                                          baseName.StartsWith("F") && (baseName.Contains("L_") || baseName.Contains("R_"));
                            break;
                        case "rear":
                            shouldModify = baseName.ToUpperInvariant().Contains("_REAR") ||
                                          (baseName.StartsWith("R") && !baseName.Contains("Rocker") && (baseName.Contains("L_") || baseName.Contains("R_")));
                            break;
                    }
                }
                
                if (shouldModify)
                {
                    feat.SetSuppression2(
                        visible ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature 
                                : (int)swFeatureSuppressionAction_e.swSuppressFeature,
                        (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    count++;
                }
                
                feat = (Feature)feat.GetNextFeature();
            }

            string action = visible ? "shown" : "hidden";
            Console.WriteLine($"{count} markers {action}");
            return true;
        }
    }
}
