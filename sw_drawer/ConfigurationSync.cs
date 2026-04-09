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
        /// Create <paramref name="configName"/> in both the top-level assembly and every
        /// direct subassembly.  Existing configurations with the same name are left untouched.
        /// </summary>
        internal static void CreateConfigurationInHierarchy(
            AssemblyDoc swAssy, ModelDoc2 swModel, string configName)
        {
            if (swModel == null || string.IsNullOrWhiteSpace(configName))
                return;

            GetOrCreateConfiguration(swModel, configName);

            object[] components = swAssy?.GetComponents(false) as object[];
            if (components == null)
                return;

            foreach (object obj in components)
            {
                Component2 comp = obj as Component2;
                ModelDoc2 compDoc = comp?.GetModelDoc2() as ModelDoc2;
                if (compDoc != null && compDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                    GetOrCreateConfiguration(compDoc, configName);
            }
        }

        /// <summary>
        /// Make <paramref name="feat"/> unsuppressed in <paramref name="targetConfigName"/>
        /// and suppressed in every other configuration of <paramref name="swModel"/>.
        /// </summary>
        internal static void ScopeFeatureToConfiguration(
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
        /// Future-proof scoping: suppress <paramref name="feat"/> in ALL configurations
        /// (including any created later), then unsuppress ONLY in <paramref name="targetConfigName"/>.
        /// </summary>
        internal static void ScopeFeatureExclusivelyToConfiguration(
            ModelDoc2 swModel, Feature feat, string targetConfigName)
        {
            if (swModel == null || feat == null || string.IsNullOrWhiteSpace(targetConfigName))
                return;

            // Suppress in all configurations (covers future ones too)
            feat.SetSuppression2(
                (int)swFeatureSuppressionAction_e.swSuppressFeature,
                (int)swInConfigurationOpts_e.swAllConfiguration,
                null);

            // Unsuppress only in the target configuration
            feat.SetSuppression2(
                (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                new string[] { targetConfigName });
        }

        /// <summary>
        /// Return the configuration named <paramref name="configName"/>, creating it if absent.
        /// </summary>
        internal static Configuration GetOrCreateConfiguration(ModelDoc2 swModel, string configName)
        {
            if (swModel == null || string.IsNullOrWhiteSpace(configName))
                return null;

            string[] existing = swModel.GetConfigurationNames() as string[];
            if (existing != null)
            {
                foreach (string name in existing)
                {
                    if (string.Equals(name, configName, StringComparison.OrdinalIgnoreCase))
                        return swModel.GetConfigurationByName(configName) as Configuration;
                }
            }

            // Derive from the active configuration so properties are inherited
            Configuration active = swModel.ConfigurationManager?.ActiveConfiguration;
            string parentName = active?.Name ?? "";

            Configuration newCfg = swModel.AddConfiguration3(
                configName,
                "",   // comment
                "",   // alternate name
                (int)swConfigurationOptions2_e.swConfigOption_DontActivate
            ) as Configuration;

            if (newCfg != null)
                Console.WriteLine($"ConfigSync: Created configuration '{configName}' in {swModel.GetTitle()}");
            else
                Console.WriteLine($"ConfigSync: Warning – failed to create '{configName}' in {swModel.GetTitle()}");

            return newCfg;
        }
    }
}
