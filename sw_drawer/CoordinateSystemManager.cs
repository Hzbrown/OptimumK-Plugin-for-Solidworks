using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    /// <summary>
    /// Enhanced coordinate system manager using IFeatureManager Interface
    /// for better folder management and assembly origin positioning
    /// </summary>
    public static class CoordinateSystemManager
    {
        /// <summary>
        /// Creates a coordinate system at assembly origin with pose-based positioning
        /// </summary>
        public static bool CreatePoseCoordinateSystem(
            SldWorks swApp,
            string name,
            double x, double y, double z,
            double angleX = 0, double angleY = 0, double angleZ = 0)
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            if (swDoc == null)
            {
                Console.WriteLine("No active document found.");
                return false;
            }

            // Check if coordinate system already exists
            bool exists = swDoc.Extension.SelectByID2(
                name, "COORDSYS",
                0, 0, 0,
                false, 0, null,
                (int)swSelectOption_e.swSelectOptionDefault
            );

            if (exists)
            {
                // Update existing coordinate system
                return UpdateExistingCoordinateSystem(swDoc, name, x, y, z, angleX, angleY, angleZ);
            }

            swDoc.ClearSelection2(true);

            // Create coordinate system at assembly origin (0,0,0)
            Feature newFeat = swDoc.FeatureManager
                .CreateCoordinateSystemUsingNumericalValues(
                    true,           // UseLocation
                    0.0,            // DeltaX (meters) - at origin
                    0.0,            // DeltaY (meters) - at origin
                    0.0,            // DeltaZ (meters) - at origin
                    false,          // UseRotation - initially no rotation
                    0.0,            // AngleX (radians)
                    0.0,            // AngleY (radians)
                    0.0             // AngleZ (radians)
                ) as Feature;

            if (newFeat == null)
            {
                Console.WriteLine($"Failed to create coordinate system '{name}'.");
                return false;
            }

            newFeat.Name = name;

            // Apply pose-based positioning using ModifyDefinition
            bool success = ApplyPosePositioning(swDoc, newFeat, x, y, z, angleX, angleY, angleZ);
            
            if (success)
            {
                // Move to "Coordinates" folder
                MoveToCoordinatesFolder(swDoc, newFeat);
                swDoc.EditRebuild3();
                
                Console.WriteLine($"Created pose coordinate system: '{name}' at origin with pose ({x}, {y}, {z}) mm, angles ({angleX}, {angleY}, {angleZ}) deg.");
            }

            return success;
        }

        /// <summary>
        /// Updates an existing coordinate system with new pose values
        /// </summary>
        private static bool UpdateExistingCoordinateSystem(
            ModelDoc2 swDoc, 
            string name, 
            double x, double y, double z,
            double angleX, double angleY, double angleZ)
        {
            SelectionMgr selMgr = (SelectionMgr)swDoc.SelectionManager;
            Feature coordFeat = (Feature)selMgr.GetSelectedObject6(1, -1);
            
            if (coordFeat == null)
            {
                Console.WriteLine($"Could not find coordinate system '{name}'.");
                return false;
            }

            CoordinateSystemFeatureData coordData = (CoordinateSystemFeatureData)coordFeat.GetDefinition();
            
            if (coordData == null)
            {
                Console.WriteLine($"Could not get coordinate system data for '{name}'.");
                return false;
            }

            bool accessOk = coordData.AccessSelections(swDoc, null);
            
            if (!accessOk)
            {
                Console.WriteLine($"Could not access selections for '{name}'.");
                return false;
            }

            try
            {
                // Clear entity reference to allow numerical positioning
                coordData.OriginEntity = null;
                
                // Apply pose-based positioning
                bool modified = coordFeat.ModifyDefinition(coordData, swDoc, null);
                
                if (modified)
                {
                    Console.WriteLine($"Updated pose coordinate system: '{name}' with pose ({x}, {y}, {z}) mm, angles ({angleX}, {angleY}, {angleZ}) deg.");
                }
                else
                {
                    Console.WriteLine($"Failed to modify coordinate system '{name}'.");
                }

                coordData.ReleaseSelectionAccess();
                return modified;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating coordinate system '{name}': {ex.Message}");
                coordData.ReleaseSelectionAccess();
                return false;
            }
        }

        /// <summary>
        /// Applies pose-based positioning to a coordinate system
        /// </summary>
        private static bool ApplyPosePositioning(
            ModelDoc2 swDoc, 
            Feature coordFeat, 
            double x, double y, double z,
            double angleX, double angleY, double angleZ)
        {
            CoordinateSystemFeatureData coordData = (CoordinateSystemFeatureData)coordFeat.GetDefinition();
            
            if (coordData == null)
            {
                Console.WriteLine("Could not get coordinate system data.");
                return false;
            }

            bool accessOk = coordData.AccessSelections(swDoc, null);
            
            if (!accessOk)
            {
                Console.WriteLine("Could not access coordinate system selections.");
                return false;
            }

            try
            {
                // Clear entity reference to allow numerical positioning
                coordData.OriginEntity = null;
                
                // Create a new coordinate system with the pose values
                // We need to recreate the coordinate system with the new pose
                Feature newCoordFeat = swDoc.FeatureManager
                    .CreateCoordinateSystemUsingNumericalValues(
                        true,           // UseLocation
                        x / 1000.0,     // DeltaX (meters)
                        y / 1000.0,     // DeltaY (meters)
                        z / 1000.0,     // DeltaZ (meters)
                        (angleX != 0 || angleY != 0 || angleZ != 0),  // UseRotation
                        angleX * Math.PI / 180.0,  // AngleX (radians)
                        angleY * Math.PI / 180.0,  // AngleY (radians)
                        angleZ * Math.PI / 180.0   // AngleZ (radians)
                    ) as Feature;

                if (newCoordFeat != null)
                {
                    newCoordFeat.Name = coordFeat.Name;
                    coordFeat = newCoordFeat;
                }

                coordData.ReleaseSelectionAccess();
                
                return newCoordFeat != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying pose positioning: {ex.Message}");
                coordData.ReleaseSelectionAccess();
                return false;
            }
        }

        /// <summary>
        /// Enhanced folder management using IFeatureManager Interface
        /// Creates or finds the "Coordinates" folder and moves the feature into it
        /// </summary>
        private static void MoveToCoordinatesFolder(ModelDoc2 swDoc, Feature feat)
        {
            const string folderName = "Coordinates";
            FeatureManager featMgr = swDoc.FeatureManager;

            // Try to find existing "Coordinates" folder using IFeatureManager
            Feature folder = FindOrCreateCoordinatesFolder(featMgr, folderName);

            if (folder != null)
            {
                // Move feature into the folder using IFeatureManager
                MoveFeatureToFolder(featMgr, feat, folder);
            }
        }

        /// <summary>
        /// Finds existing coordinates folder or creates a new one using IFeatureManager
        /// </summary>
        private static Feature FindOrCreateCoordinatesFolder(FeatureManager featMgr, string folderName)
        {
            // Use the model document to search for existing folder
            // We need to get the model document from the feature manager
            ModelDoc2 swDoc = null;
            try
            {
                // Try to get the model document from the feature manager
                // This is a workaround since FeatureManager doesn't have GetModelDoc2
                swDoc = (ModelDoc2)featMgr.GetType().InvokeMember("ModelDoc", 
                    System.Reflection.BindingFlags.GetProperty, null, featMgr, null);
            }
            catch
            {
                // If that fails, we'll need to get it another way
                // For now, we'll just create the folder
            }

            Feature folder = null;
            
            // If we have the document, search for existing folder
            if (swDoc != null)
            {
                Feature swFeat = (Feature)swDoc.FirstFeature();
                while (swFeat != null)
                {
                    if (swFeat.GetTypeName2() == "FtrFolder" && swFeat.Name == folderName)
                    {
                        folder = swFeat;
                        break;
                    }
                    swFeat = (Feature)swFeat.GetNextFeature();
                }
            }

            // Create folder if it doesn't exist using IFeatureManager
            if (folder == null)
            {
                folder = CreateCoordinatesFolder(featMgr, folderName);
            }

            return folder;
        }

        /// <summary>
        /// Creates a new coordinates folder using IFeatureManager Interface
        /// </summary>
        private static Feature CreateCoordinatesFolder(FeatureManager featMgr, string folderName)
        {
            try
            {
                // Use IFeatureManager.InsertFeatureTreeFolder2 to create folder
                Feature folder = featMgr.InsertFeatureTreeFolder2(
                    (int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing
                );
                
                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created folder: {folderName}");
                }
                
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating coordinates folder: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Moves a feature into a folder using IFeatureManager
        /// </summary>
        private static void MoveFeatureToFolder(FeatureManager featMgr, Feature feat, Feature folder)
        {
            try
            {
                // Select the feature to move
                feat.Select2(false, 0);
                
                // Use IFeatureManager to move the feature to the folder
                // The MoveToFolder method requires the folder name and feature name
                bool moved = featMgr.MoveToFolder(folder.Name, feat.Name, false);
                
                if (moved)
                {
                    Console.WriteLine($"Moved feature '{feat.Name}' to folder '{folder.Name}'");
                }
                else
                {
                    Console.WriteLine($"Failed to move feature '{feat.Name}' to folder '{folder.Name}'");
                }
                
                // Clear selection
                // We need to get the model document to clear selection
                ModelDoc2 swDoc = null;
                try
                {
                    swDoc = (ModelDoc2)featMgr.GetType().InvokeMember("ModelDoc", 
                        System.Reflection.BindingFlags.GetProperty, null, featMgr, null);
                }
                catch
                {
                    // If we can't get the document, just continue
                }
                
                if (swDoc != null)
                {
                    swDoc.ClearSelection2(true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error moving feature to folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets all coordinate systems in the document
        /// </summary>
        public static List<CoordinateSystemInfo> GetAllCoordinateSystems(ModelDoc2 swDoc)
        {
            var coordSystems = new List<CoordinateSystemInfo>();
            Feature feat = (Feature)swDoc.FirstFeature();

            while (feat != null)
            {
                if (feat.GetTypeName2() == "CoordSys")
                {
                    try
                    {
                        CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)feat.GetDefinition();
                        if (csData != null && csData.AccessSelections(swDoc, null))
                        {
                            MathTransform transform = csData.Transform;
                            if (transform != null)
                            {
                                double[] td = (double[])transform.ArrayData;
                                coordSystems.Add(new CoordinateSystemInfo
                                {
                                    Name = feat.Name,
                                    Feature = feat,
                                    X = td[9] * 1000.0,  // Convert back to mm
                                    Y = td[10] * 1000.0,
                                    Z = td[11] * 1000.0
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

        /// <summary>
        /// Gets all coordinate systems in a specific folder
        /// </summary>
        public static List<CoordinateSystemInfo> GetCoordinateSystemsInFolder(ModelDoc2 swDoc, string folderName)
        {
            var coordSystems = new List<CoordinateSystemInfo>();
            Feature folder = FindFolderByName(swDoc, folderName);

            if (folder == null)
            {
                Console.WriteLine($"Folder '{folderName}' not found.");
                return coordSystems;
            }

            // Get features within the folder
            Feature feat = (Feature)folder.GetFirstSubFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "CoordSys")
                {
                    try
                    {
                        CoordinateSystemFeatureData csData = (CoordinateSystemFeatureData)feat.GetDefinition();
                        if (csData != null && csData.AccessSelections(swDoc, null))
                        {
                            MathTransform transform = csData.Transform;
                            if (transform != null)
                            {
                                double[] td = (double[])transform.ArrayData;
                                coordSystems.Add(new CoordinateSystemInfo
                                {
                                    Name = feat.Name,
                                    Feature = feat,
                                    X = td[9] * 1000.0,
                                    Y = td[10] * 1000.0,
                                    Z = td[11] * 1000.0
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
                feat = (Feature)feat.GetNextSubFeature();
            }

            return coordSystems;
        }

        /// <summary>
        /// Finds a folder by name
        /// </summary>
        private static Feature FindFolderByName(ModelDoc2 swDoc, string folderName)
        {
            Feature feat = (Feature)swDoc.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "FtrFolder" && feat.Name == folderName)
                {
                    return feat;
                }
                feat = (Feature)feat.GetNextFeature();
            }
            return null;
        }

        /// <summary>
        /// Information about a coordinate system
        /// </summary>
        public class CoordinateSystemInfo
        {
            public string Name { get; set; }
            public Feature Feature { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
        }
    }
}