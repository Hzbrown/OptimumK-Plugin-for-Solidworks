using System;

namespace sw_drawer
{
    /// <summary>
    /// State enum for hardpoint operations progress reporting
    /// </summary>
    public enum HardpointState
    {
        Initializing,
        LoadingJson,
        LoadingMarkerPart,
        InsertingBodies,
        RenamingBodies,
        ApplyingColors,
        CreatingCoordinateSystems,
        CreatingHardpointsFolder,
        CreatingTransformsFolder,
        CreatingTransforms,
        UpdatingSuppression,
        Rebuilding,
        Complete
    }

    /// <summary>
    /// Helper class for state descriptions
    /// </summary>
    public static class HardpointStateHelper
    {
        public static string GetDescription(HardpointState state)
        {
            switch (state)
            {
                case HardpointState.Initializing:
                    return "Starting...";
                case HardpointState.LoadingJson:
                    return "Loading JSON data...";
                case HardpointState.LoadingMarkerPart:
                    return "Loading marker part...";
                case HardpointState.InsertingBodies:
                    return "Inserting marker bodies...";
                case HardpointState.RenamingBodies:
                    return "Renaming bodies...";
                case HardpointState.ApplyingColors:
                    return "Applying colors...";
                case HardpointState.CreatingCoordinateSystems:
                    return "Creating coordinate systems...";
                case HardpointState.CreatingHardpointsFolder:
                    return "Creating Hardpoints folder...";
                case HardpointState.CreatingTransformsFolder:
                    return "Creating Transforms folder...";
                case HardpointState.CreatingTransforms:
                    return "Creating transform features...";
                case HardpointState.UpdatingSuppression:
                    return "Updating suppression...";
                case HardpointState.Rebuilding:
                    return "Rebuilding model...";
                case HardpointState.Complete:
                    return "Done";
                default:
                    return state.ToString();
            }
        }
    }
}