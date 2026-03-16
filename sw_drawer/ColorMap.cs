using System.Collections.Generic;

namespace sw_drawer
{
    /// <summary>
    /// Canonical color mapping for suspension hardpoint naming prefixes (RGB 0-255).
    /// Used by both HardpointRunner and InsertMarker so colors stay consistent.
    /// </summary>
    internal static class ColorMap
    {
        /// <summary>
        /// Maps name token → RGB array. Checked in dictionary order; first match wins.
        /// Tie-rod tokens are checked explicitly before this table in GetColorForName.
        /// </summary>
        public static readonly Dictionary<string, int[]> Colors = new Dictionary<string, int[]>
        {
            { "CHAS_", new[] { 255,   0,   0 } },   // Red           – Chassis
            { "UPRI_", new[] {   0,   0, 255 } },   // Blue          – Upright
            { "ROCK_", new[] {   0, 128, 255 } },   // Light Blue    – Rocker
            { "NSMA_", new[] { 255, 192, 203 } },   // Pink          – Non-Sprung Mass
            { "PUSH_", new[] {   0, 255,   0 } },   // Green         – Pushrod
            { "TIER_", new[] { 255, 165,   0 } },   // Orange        – Tie Rod
            { "DAMP_", new[] { 128,   0, 128 } },   // Purple        – Damper
            { "ARBA_", new[] { 255, 255,   0 } },   // Yellow        – ARB
            { "_FRONT", new[] {  0, 200, 200 } },   // Cyan          – Front (fallback)
            { "_REAR",  new[] { 200, 100,   0 } },  // Brown         – Rear  (fallback)
            { "wheel",  new[] {  64,  64,  64 } },  // Dark Gray     – Wheels
        };

        /// <summary>Shared tie-rod colour used before the table is consulted.</summary>
        public static readonly int[] TieRodColor = { 255, 165, 0 };

        /// <summary>Fallback colour when no prefix matches.</summary>
        public static readonly int[] DefaultColor = { 128, 128, 128 };
    }
}
