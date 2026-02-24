using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    public static class FeatureVisibility
    {
        /// <summary>
        /// Gets a feature by name from the active document.
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
        /// Shows or hides a feature by name in the active document.
        /// </summary>
        /// <param name="swApp">SolidWorks application instance</param>
        /// <param name="featureName">Name of the feature to show/hide</param>
        /// <param name="visible">True to show, false to hide</param>
        /// <returns>True if successful, false otherwise</returns>
        public static bool SetFeatureVisibility(SldWorks swApp, string featureName, bool visible)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document.");
                return false;
            }

            Feature swFeature = GetFeatureByName(swModel, featureName);
            if (swFeature == null)
            {
                Console.WriteLine($"Warning: Feature '{featureName}' not found.");
                return false;
            }

            try
            {
                int suppressState = visible 
                    ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature 
                    : (int)swFeatureSuppressionAction_e.swSuppressFeature;

                bool result = swFeature.SetSuppression2(
                    suppressState,
                    (int)swInConfigurationOpts_e.swThisConfiguration,
                    null
                );

                if (result)
                {
                    Console.WriteLine($"Feature '{featureName}' {(visible ? "shown" : "hidden")} successfully.");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not {(visible ? "show" : "hide")} feature '{featureName}'.");
                }

                swModel.GraphicsRedraw2();
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting visibility for '{featureName}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Shows or hides multiple features matching a suffix pattern.
        /// </summary>
        public static int SetFeatureVisibilityBySuffix(SldWorks swApp, string suffix, bool visible)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document.");
                return 0;
            }

            int count = 0;
            Feature swFeature = (Feature)swModel.FirstFeature();

            while (swFeature != null)
            {
                string name = swFeature.Name;
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    if (SetFeatureVisibilityDirect(swFeature, visible, swModel))
                    {
                        count++;
                    }
                }
                swFeature = (Feature)swFeature.GetNextFeature();
            }

            swModel.GraphicsRedraw2();
            Console.WriteLine($"{(visible ? "Shown" : "Hidden")} {count} features with suffix '{suffix}'.");
            return count;
        }

        /// <summary>
        /// Shows or hides features containing a specific substring.
        /// </summary>
        public static int SetFeatureVisibilityBySubstring(SldWorks swApp, string substring, bool visible)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document.");
                return 0;
            }

            int count = 0;
            Feature swFeature = (Feature)swModel.FirstFeature();

            while (swFeature != null)
            {
                string name = swFeature.Name;
                if (name.IndexOf(substring, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (SetFeatureVisibilityDirect(swFeature, visible, swModel))
                    {
                        count++;
                    }
                }
                swFeature = (Feature)swFeature.GetNextFeature();
            }

            swModel.GraphicsRedraw2();
            Console.WriteLine($"{(visible ? "Shown" : "Hidden")} {count} features containing '{substring}'.");
            return count;
        }

        /// <summary>
        /// Shows or hides wheel features.
        /// </summary>
        public static int SetWheelVisibility(SldWorks swApp, bool visible, bool frontOnly = false, bool rearOnly = false)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Console.WriteLine("Error: No active document.");
                return 0;
            }

            int count = 0;
            Feature swFeature = (Feature)swModel.FirstFeature();

            while (swFeature != null)
            {
                string name = swFeature.Name;
                bool isWheel = name.IndexOf("_wheel", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isFront = name.StartsWith("F", StringComparison.OrdinalIgnoreCase);
                bool isRear = name.StartsWith("R", StringComparison.OrdinalIgnoreCase);

                if (isWheel)
                {
                    bool shouldProcess = (!frontOnly && !rearOnly) || 
                                         (frontOnly && isFront) || 
                                         (rearOnly && isRear);
                    if (shouldProcess && SetFeatureVisibilityDirect(swFeature, visible, swModel))
                    {
                        count++;
                    }
                }
                swFeature = (Feature)swFeature.GetNextFeature();
            }

            swModel.GraphicsRedraw2();
            string scope = frontOnly ? "front " : (rearOnly ? "rear " : "");
            Console.WriteLine($"{(visible ? "Shown" : "Hidden")} {count} {scope}wheel features.");
            return count;
        }

        /// <summary>
        /// Internal method to set feature visibility directly on a Feature object.
        /// </summary>
        private static bool SetFeatureVisibilityDirect(Feature swFeature, bool visible, ModelDoc2 swModel)
        {
            try
            {
                int suppressState = visible 
                    ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature 
                    : (int)swFeatureSuppressionAction_e.swSuppressFeature;

                return swFeature.SetSuppression2(
                    suppressState,
                    (int)swInConfigurationOpts_e.swThisConfiguration,
                    null
                );
            }
            catch
            {
                return false;
            }
        }
    }
}
