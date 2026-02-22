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
        /// Uses ModifyDefinition to preserve feature ID when updating.
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

            // Convert mm -> meters, degrees -> radians
            double deltaX = x / 1000.0;
            double deltaY = y / 1000.0;
            double deltaZ = z / 1000.0;
            double radX = angleX * Math.PI / 180.0;
            double radY = angleY * Math.PI / 180.0;
            double radZ = angleZ * Math.PI / 180.0;

            // Check if coordinate system exists
            bool exists = swDoc.Extension.SelectByID2(
                name, "COORDSYS",
                0, 0, 0,
                false, 0, null,
                (int)swSelectOption_e.swSelectOptionDefault
            );

            if (exists)
            {
                // Get selected feature and modify it in place
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
                            // Clear entity reference to allow numerical positioning
                            coordData.OriginEntity = null;
                            
                            // Apply the changes
                            bool modified = coordFeat.ModifyDefinition(coordData, swDoc, null);
                            coordData.ReleaseSelectionAccess();
                            
                            if (modified)
                            {
                                swDoc.ClearSelection2(true);
                                swDoc.EditRebuild3();
                                
                                Console.WriteLine(useRotation
                                    ? $"Updated: '{name}' at ({x}, {y}, {z}) mm, angles ({angleX}, {angleY}, {angleZ}) deg."
                                    : $"Updated: '{name}' at ({x}, {y}, {z}) mm.");
                                return true;
                            }
                            else
                            {
                                Console.WriteLine($"ModifyDefinition failed for '{name}'.");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"AccessSelections failed for '{name}'.");
                        }
                    }
                }
                
                swDoc.ClearSelection2(true);
                return false;
            }

            swDoc.ClearSelection2(true);

            // Create new coordinate system using numerical values
            Feature newFeat = swDoc.FeatureManager
                .CreateCoordinateSystemUsingNumericalValues(
                    true,           // UseLocation
                    deltaX,         // DeltaX (meters)
                    deltaY,         // DeltaY (meters)
                    deltaZ,         // DeltaZ (meters)
                    useRotation,    // UseRotation
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
            swDoc.EditRebuild3();

            Console.WriteLine(useRotation
                ? $"Created: '{name}' at ({x}, {y}, {z}) mm, angles ({angleX}, {angleY}, {angleZ}) deg."
                : $"Created: '{name}' at ({x}, {y}, {z}) mm.");

            return true;
        }
    }
}