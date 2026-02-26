// https://help.solidworks.com/2021/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~CreateCoordinateSystemUsingNumericalValues.html
using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public static class InsertCoordinate
    {
        public static Feature InsertCoordinateSystemFeature(
            ModelDoc2 swDoc,
            string name,
            double x, double y, double z,
            double angleX = 0, double angleY = 0, double angleZ = 0,
            bool createAtOrigin = true,
            string folderName = "Coordinates",
            bool hideInGui = false)
        {
            if (swDoc == null)
            {
                Console.WriteLine("No active document found.");
                return null;
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
                SelectionMgr selMgr = (SelectionMgr)swDoc.SelectionManager;
                Feature existingFeat = (Feature)selMgr.GetSelectedObject6(1, -1);
                swDoc.ClearSelection2(true);

                if (existingFeat != null)
                {
                    if (!string.IsNullOrWhiteSpace(folderName))
                    {
                        MoveFeatureToFolder(swDoc, existingFeat, folderName);
                    }

                    if (hideInGui)
                    {
                        HideReferenceGeometry(swDoc, existingFeat);
                    }

                    Console.WriteLine($"Coordinate system '{name}' already exists.");
                    return existingFeat;
                }
            }

            swDoc.ClearSelection2(true);

            double deltaX = createAtOrigin ? 0.0 : x / 1000.0; // mm -> m
            double deltaY = createAtOrigin ? 0.0 : y / 1000.0;
            double deltaZ = createAtOrigin ? 0.0 : z / 1000.0;

            Feature newFeat = swDoc.FeatureManager
                .CreateCoordinateSystemUsingNumericalValues(
                    true,
                    deltaX,
                    deltaY,
                    deltaZ,
                    useRotation,
                    radX,
                    radY,
                    radZ
                ) as Feature;

            if (newFeat == null)
            {
                Console.WriteLine($"Failed to create coordinate system '{name}'.");
                return null;
            }

            newFeat.Name = name;

            if (!string.IsNullOrWhiteSpace(folderName))
            {
                MoveFeatureToFolder(swDoc, newFeat, folderName);
            }

            if (hideInGui)
            {
                HideReferenceGeometry(swDoc, newFeat);
            }

            swDoc.EditRebuild3();

            if (createAtOrigin)
            {
                Console.WriteLine(useRotation
                    ? $"Coordinate system '{name}' created at origin, angles ({angleX}, {angleY}, {angleZ}) deg."
                    : $"Coordinate system '{name}' created at origin.");
            }
            else
            {
                Console.WriteLine(useRotation
                    ? $"Coordinate system '{name}' created at ({x}, {y}, {z}) mm, angles ({angleX}, {angleY}, {angleZ}) deg."
                    : $"Coordinate system '{name}' created at ({x}, {y}, {z}) mm.");
            }

            return newFeat;
        }

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
            Feature feat = InsertCoordinateSystemFeature(
                swDoc,
                name,
                x, y, z,
                angleX, angleY, angleZ,
                createAtOrigin: true,
                folderName: "Coordinates",
                hideInGui: false);
            return feat != null;
        }

        /// <summary>
        /// Moves a feature into a folder, creating it if it doesn't exist.
        /// Uses IFeatureManager.InsertFeatureTreeFolder2 and MoveToFolder.
        /// </summary>
        private static void MoveFeatureToFolder(ModelDoc2 swDoc, Feature feat, string folderName)
        {
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

        private static void HideReferenceGeometry(ModelDoc2 swDoc, Feature feat)
        {
            try
            {
                swDoc.ClearSelection2(true);
                bool selected = feat.Select2(false, 0);
                if (selected)
                {
                    swDoc.BlankRefGeom();
                }
                swDoc.ClearSelection2(true);
            }
            catch
            {
                try { swDoc.ClearSelection2(true); } catch { }
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
