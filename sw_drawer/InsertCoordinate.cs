// https://help.solidworks.com/2021/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~CreateCoordinateSystemUsingNumericalValues.html
using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public static class InsertCoordinate
    {
        /// <summary>
        /// Inserts or updates a coordinate system in the active SolidWorks document.
        /// Creates at assembly origin (0,0,0) - position comes in pose step.
        /// Rotation angles are applied immediately (for wheel toe/camber).
        /// </summary>
        public static bool InsertCoordinateSystem(
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

            bool useRotation = (angleX != 0 || angleY != 0 || angleZ != 0);

            // Degrees -> radians for rotation
            double radX = angleX * Math.PI / 180.0;
            double radY = angleY * Math.PI / 180.0;
            double radZ = angleZ * Math.PI / 180.0;

            // Check if coordinate system already exists
            bool exists = swDoc.Extension.SelectByID2(
                name, "COORDSYS",
                0, 0, 0,
                false, 0, null,
                (int)swSelectOption_e.swSelectOptionDefault
            );

            if (exists)
            {
                // Preserve the existing feature - just clear its entity reference
                SelectionMgr selMgr = (SelectionMgr)swDoc.SelectionManager;
                Feature coordFeat = (Feature)selMgr.GetSelectedObject6(1, -1);

                if (coordFeat != null)
                {
                    CoordinateSystemFeatureData coordData = (CoordinateSystemFeatureData)coordFeat.GetDefinition();
                    if (coordData != null)
                    {
                        bool accessOk = coordData.AccessSelections(swDoc, null);
                        if (accessOk)
                        {
                            coordData.OriginEntity = null;
                            bool modified = coordFeat.ModifyDefinition(coordData, swDoc, null);
                            coordData.ReleaseSelectionAccess();

                            if (modified)
                            {
                                swDoc.ClearSelection2(true);
                                swDoc.EditRebuild3();
                                Console.WriteLine($"Already exists, kept: '{name}'");
                                return true;
                            }
                        }
                    }
                }

                swDoc.ClearSelection2(true);
                return false;
            }

            swDoc.ClearSelection2(true);

            // Create coordinate system at assembly origin (0, 0, 0).
            // Position comes in the pose step via mates/transforms.
            Feature newFeat = swDoc.FeatureManager
                .CreateCoordinateSystemUsingNumericalValues(
                    true,           // UseLocation
                    0.0,            // DeltaX = 0 (at assembly origin)
                    0.0,            // DeltaY = 0 (at assembly origin)
                    0.0,            // DeltaZ = 0 (at assembly origin)
                    useRotation,    // UseRotation (toe/camber for wheels)
                    radX,           // AngleX (radians)
                    radY,           // AngleY (radians)
                    radZ            // AngleZ (radians)
                ) as Feature;

            if (newFeat == null)
            {
                Console.WriteLine($"Failed to create coordinate system '{name}'.");
                return false;
            }

            newFeat.Name = name;

            // Organise into "Coordinates" folder
            MoveToCoordinatesFolder(swDoc, newFeat);

            swDoc.EditRebuild3();

            Console.WriteLine(useRotation
                ? $"Created: '{name}' at origin, angles ({angleX}, {angleY}, {angleZ}) deg."
                : $"Created: '{name}' at origin.");

            return true;
        }

        /// <summary>
        /// Moves a feature into the "Coordinates" folder, creating it if it doesn't exist.
        /// Uses IFeatureManager.InsertFeatureTreeFolder2 and MoveToFolder.
        /// </summary>
        private static void MoveToCoordinatesFolder(ModelDoc2 swDoc, Feature feat)
        {
            const string folderName = "Coordinates";
            FeatureManager featMgr = swDoc.FeatureManager;

            // Search for existing folder
            Feature folder = FindFolder(swDoc, folderName);

            if (folder == null)
            {
                // Select the new feature so the folder is created containing it
                swDoc.ClearSelection2(true);
                feat.Select2(false, 0);

                folder = featMgr.InsertFeatureTreeFolder2(
                    (int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing) as Feature;

                if (folder != null)
                {
                    folder.Name = folderName;
                    Console.WriteLine($"Created folder '{folderName}'");
                }
                else
                {
                    Console.WriteLine($"Warning: could not create folder '{folderName}'");
                }

                swDoc.ClearSelection2(true);
            }
            else
            {
                // Select the feature and move it into the existing folder
                swDoc.ClearSelection2(true);
                feat.Select2(false, 0);

                // IFeatureManager.MoveToFolder(DestFolderName, MoveFromFeat, IncludeChildren)
                bool moved = featMgr.MoveToFolder(folderName, feat.Name, false);
                if (!moved)
                {
                    Console.WriteLine($"Warning: could not move '{feat.Name}' into '{folderName}'");
                }

                swDoc.ClearSelection2(true);
            }
        }

        /// <summary>Finds a FtrFolder feature by name, or returns null.</summary>
        private static Feature FindFolder(ModelDoc2 swDoc, string folderName)
        {
            Feature f = (Feature)swDoc.FirstFeature();
            while (f != null)
            {
                if (f.GetTypeName2() == "FtrFolder" && f.Name == folderName)
                    return f;
                f = (Feature)f.GetNextFeature();
            }
            return null;
        }
    }
}
