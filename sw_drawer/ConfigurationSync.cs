using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    /// <summary>
    /// Synchronises configurations across an assembly hierarchy.
    /// When the top-level Kinematic Skeleton is set to a configuration (e.g. "Static"),
    /// all subassemblies (Chassis, FL Corner, …) are switched to the matching config.
    /// Features (coordinate systems) are scoped so they only appear in their target config.
    /// </summary>
    internal static class ConfigurationSync
    {
        /// <summary>
        /// Ensure every direct subassembly references <paramref name="targetConfigName"/>.
        /// If a subassembly does not yet contain that configuration it is created first.
        /// </summary>
        internal static int SyncSubassemblyConfigurations(AssemblyDoc swAssy, string targetConfigName)
        {
            if (swAssy == null || string.IsNullOrWhiteSpace(targetConfigName))
                return 0;

            object[] components = swAssy.GetComponents(false) as object[];
            if (components == null)
                return 0;

            int synced = 0;
            foreach (object obj in components)
            {
                Component2 comp = obj as Component2;
                if (comp == null)
                    continue;

                ModelDoc2 compDoc = comp.GetModelDoc2() as ModelDoc2;
                if (compDoc == null || compDoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                    continue;

                GetOrCreateConfiguration(compDoc, targetConfigName);
                comp.ReferencedConfiguration = targetConfigName;
                synced++;
                Console.WriteLine($"ConfigSync: {comp.Name2} → {targetConfigName}");
            }

            return synced;
        }

        /// <summary>
        /// Create <paramref name="configName"/> in the top-level assembly and every direct
        /// subassembly. If the config is newly created, every existing pose-scoped feature
        /// (coordinate systems inside "* Transforms" folders, and mates named with a pose
        /// prefix) is explicitly suppressed in it so future poses inherit suppression of
        /// earlier poses' features without manual intervention.
        /// </summary>
        internal static void CreateConfigurationInHierarchy(
            AssemblyDoc swAssy, ModelDoc2 swModel, string configName)
        {
            if (swModel == null || string.IsNullOrWhiteSpace(configName))
                return;

            GetOrCreateConfiguration(swModel, configName, out bool wasNewTop);
            if (wasNewTop)
                SuppressExistingPoseFeaturesInConfig(swModel, configName);

            object[] components = swAssy?.GetComponents(false) as object[];
            if (components == null)
                return;

            foreach (object obj in components)
            {
                Component2 comp = obj as Component2;
                ModelDoc2 compDoc = comp?.GetModelDoc2() as ModelDoc2;
                if (compDoc == null || compDoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                    continue;

                GetOrCreateConfiguration(compDoc, configName, out bool wasNewSub);
                if (wasNewSub)
                    SuppressExistingPoseFeaturesInConfig(compDoc, configName);
            }
        }

        /// <summary>
        /// Scope <paramref name="feat"/> exclusively to <paramref name="targetConfigName"/>:
        /// unsuppressed in the target, suppressed in every other currently-existing config.
        /// Future configs will auto-inherit suppression if created via
        /// <see cref="CreateConfigurationInHierarchy"/> after the feature exists.
        /// </summary>
        internal static void ScopeFeatureExclusivelyToConfiguration(
            ModelDoc2 swModel, Feature feat, string targetConfigName)
        {
            if (swModel == null || feat == null || string.IsNullOrWhiteSpace(targetConfigName))
                return;

            string[] configNames = swModel.GetConfigurationNames() as string[];
            if (configNames == null || configNames.Length == 0)
                return;

            foreach (string cfgName in configNames)
            {
                int state = string.Equals(cfgName, targetConfigName, StringComparison.OrdinalIgnoreCase)
                    ? (int)swFeatureSuppressionAction_e.swUnSuppressFeature
                    : (int)swFeatureSuppressionAction_e.swSuppressFeature;

                feat.SetSuppression2(
                    state,
                    (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                    new string[] { cfgName });
            }
        }

        /// <summary>
        /// Find all features belonging to any existing pose in <paramref name="swModel"/>
        /// and suppress them in <paramref name="newConfigName"/>. Pose features are:
        ///   • CoordSys features whose owning folder ends with " Transforms"
        ///   • Mate features whose name starts with "{poseName} " for any discovered pose
        /// </summary>
        private static void SuppressExistingPoseFeaturesInConfig(ModelDoc2 swModel, string newConfigName)
        {
            if (swModel == null) return;

            const string folderSuffix = " Transforms";
            string[] onlyNew = new string[] { newConfigName };

            // Collect pose names from existing "* Transforms" folders
            var poseNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var poseFeatures = new List<Feature>();

            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();
                string name = feat.Name;

                if (typeName == "FtrFolder" && name.EndsWith(folderSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    string poseName = name.Substring(0, name.Length - folderSuffix.Length);
                    if (!string.IsNullOrWhiteSpace(poseName))
                    {
                        poseNames.Add(poseName);
                        poseFeatures.Add(feat);
                    }
                }
                feat = (Feature)feat.GetNextFeature();
            }

            // Second pass: CoordSys features whose name starts with a known pose prefix
            // (coordinate systems may live inside folders; also catch any that escaped)
            feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();
                string name = feat.Name;

                if (typeName == "CoordSys")
                {
                    foreach (string poseName in poseNames)
                    {
                        if (name.StartsWith(poseName + " ", StringComparison.OrdinalIgnoreCase))
                        {
                            poseFeatures.Add(feat);
                            break;
                        }
                    }
                }
                feat = (Feature)feat.GetNextFeature();
            }

            // Third pass: mates inside the MateGroup whose name matches a pose prefix
            Feature mateGroup = FindMateGroup(swModel);
            if (mateGroup != null)
            {
                Feature subFeat = (Feature)mateGroup.GetFirstSubFeature();
                while (subFeat != null)
                {
                    foreach (string poseName in poseNames)
                    {
                        if (subFeat.Name.StartsWith(poseName + " ", StringComparison.OrdinalIgnoreCase))
                        {
                            poseFeatures.Add(subFeat);
                            break;
                        }
                    }
                    subFeat = (Feature)subFeat.GetNextSubFeature();
                }
            }

            // Suppress all collected pose features in the newly-created config
            foreach (Feature poseFeat in poseFeatures)
            {
                try
                {
                    poseFeat.SetSuppression2(
                        (int)swFeatureSuppressionAction_e.swSuppressFeature,
                        (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                        onlyNew);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ConfigSync: Warning – could not suppress '{poseFeat.Name}' in '{newConfigName}': {ex.Message}");
                }
            }

            if (poseFeatures.Count > 0)
                Console.WriteLine($"ConfigSync: Suppressed {poseFeatures.Count} existing pose features in new config '{newConfigName}'");
        }

        /// <summary>Find the MateGroup feature in a model (assembly).</summary>
        private static Feature FindMateGroup(ModelDoc2 swModel)
        {
            Feature feat = (Feature)swModel.FirstFeature();
            while (feat != null)
            {
                if (feat.GetTypeName2() == "MateGroup")
                    return feat;
                feat = (Feature)feat.GetNextFeature();
            }
            return null;
        }

        /// <summary>
        /// Return the configuration named <paramref name="configName"/>, creating it if absent.
        /// </summary>
        internal static Configuration GetOrCreateConfiguration(ModelDoc2 swModel, string configName)
        {
            GetOrCreateConfiguration(swModel, configName, out _);
            return swModel?.GetConfigurationByName(configName) as Configuration;
        }

        /// <summary>
        /// Return true if the config exists or was created.  <paramref name="wasCreated"/>
        /// reports whether a new configuration was added (vs reusing an existing one).
        /// </summary>
        private static bool GetOrCreateConfiguration(ModelDoc2 swModel, string configName, out bool wasCreated)
        {
            wasCreated = false;
            if (swModel == null || string.IsNullOrWhiteSpace(configName))
                return false;

            string[] existing = swModel.GetConfigurationNames() as string[];
            if (existing != null)
            {
                foreach (string name in existing)
                {
                    if (string.Equals(name, configName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }

            Configuration newCfg = swModel.AddConfiguration3(
                configName,
                "",
                "",
                (int)swConfigurationOptions2_e.swConfigOption_DontActivate
            ) as Configuration;

            if (newCfg != null)
            {
                wasCreated = true;
                Console.WriteLine($"ConfigSync: Created configuration '{configName}' in {swModel.GetTitle()}");
                return true;
            }

            Console.WriteLine($"ConfigSync: Warning – failed to create '{configName}' in {swModel.GetTitle()}");
            return false;
        }
    }
}
