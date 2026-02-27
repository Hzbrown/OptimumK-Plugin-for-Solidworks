# Enhanced Coordinate Insertion for SolidWorks Plugin

## Overview

This document describes the enhanced coordinate insertion functionality for the SolidWorks Plugin for SolidWorks. The enhancements focus on three key areas:

1. **Assembly Origin Positioning** - All coordinate systems are created at assembly origin (0,0,0)
2. **Pose-Based Positioning** - Position comes from pose data rather than direct coordinate insertion
3. **Enhanced Folder Management** - Using IFeatureManager Interface for better organization
4. **Virtual Part Editing** - Capabilities to edit virtual parts using SOLIDWORKS API

## Key Features

### 1. Assembly Origin Positioning

**Before:** Coordinate systems were created at their actual 3D positions in the assembly.

**After:** All coordinate systems are created at the assembly origin (0,0,0) and positioned using pose data.

**Benefits:**
- Maintains floating position at assembly origin
- Easier to manage and organize coordinate systems
- Better integration with assembly workflows
- Simplified coordinate system management

### 2. Pose-Based Positioning

**Implementation:** Uses `ModifyDefinition` method to apply pose values to coordinate systems.

**Features:**
- Position values (X, Y, Z) in millimeters
- Rotation angles (AngleX, AngleY, AngleZ) in degrees
- Automatic conversion to meters and radians for SolidWorks API
- Updates existing coordinate systems with new pose values

**API Usage:**
```csharp
// Create coordinate system at origin
Feature newFeat = swDoc.FeatureManager.CreateCoordinateSystemUsingNumericalValues(
    true,           // UseLocation
    0.0,            // DeltaX (meters) - at origin
    0.0,            // DeltaY (meters) - at origin  
    0.0,            // DeltaZ (meters) - at origin
    false,          // UseRotation - initially no rotation
    0.0,            // AngleX (radians)
    0.0,            // AngleY (radians)
    0.0             // AngleZ (radians)
);

// Apply pose-based positioning
bool success = ApplyPosePositioning(swDoc, newFeat, x, y, z, angleX, angleY, angleZ);
```

### 3. Enhanced Folder Management

**Implementation:** Uses `IFeatureManager` Interface for advanced folder operations.

**Key Methods:**
- `InsertFeatureTreeFolder2()` - Creates new folders
- `MoveToFolder()` - Moves features to folders
- `GetTypeName2()` - Identifies feature types (e.g., "FtrFolder")
- `Name` property - Sets folder names

**Features:**
- Automatic creation of "Coordinates" folder
- Organizes coordinate systems by type
- Better integration with existing folder structure
- Enhanced error handling and recovery

**API Usage:**
```csharp
// Create coordinates folder
Feature folder = featMgr.InsertFeatureTreeFolder2(
    (int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing
);
folder.Name = "Coordinates";

// Move coordinate system to folder
feat.Select2(false, 0);
bool moved = featMgr.MoveToFolder(folder.Name, feat.Name, false);
```

### 4. Virtual Part Editing

**Implementation:** Accesses virtual parts' underlying documents for modifications.

**Features:**
- Enter "edit component" mode programmatically
- Modify geometry in virtual parts
- Modify properties (custom properties, material)
- Add coordinate systems to virtual parts
- Batch edit multiple virtual parts
- Get virtual part information and properties

**API Usage:**
```csharp
// Edit virtual part
bool success = VirtualPartEditor.EditVirtualPart(swApp, virtualComponent, (partDoc) =>
{
    // Modify the part document
    // Add coordinate systems, modify geometry, etc.
});

// Add coordinate system to virtual part
bool added = VirtualPartEditor.AddCoordinateSystemToVirtualPart(
    swApp, virtualComponent, "CS_Name", x, y, z);
```

## File Structure

### New Files

1. **`sw_drawer/CoordinateSystemManager.cs`**
   - Enhanced coordinate system management
   - Assembly origin positioning
   - Pose-based positioning
   - Enhanced folder management
   - Coordinate system querying

2. **`sw_drawer/VirtualPartEditor.cs`**
   - Virtual part editing capabilities
   - Edit component mode access
   - Geometry and property modification
   - Coordinate system addition to virtual parts

3. **`test_coordinate_insertion.py`**
   - Comprehensive test suite
   - Integration testing
   - Feature validation

### Modified Files

1. **`sw_drawer/InsertCoordinate.cs`** (existing)
   - Original coordinate insertion logic
   - Can be enhanced with new features

2. **`sw_drawer/HardpointRunner.cs`** (existing)
   - Hardpoint coordinate system creation
   - Can integrate with enhanced features

## Command Line Interface

### Coordinate System Creation

```bash
# Basic coordinate system creation
SuspensionTools.exe <name> <x> <y> <z> [angleX] [angleY] [angleZ]

# Example
SuspensionTools.exe Test_CS 100 200 300 0 0 0
```

**Features:**
- Creates coordinate system at assembly origin
- Applies pose-based positioning
- Automatically moves to "Coordinates" folder
- Updates existing coordinate systems if they exist

### Hardpoints Command

```bash
# Create hardpoints from JSON
SuspensionTools.exe hardpoints add <jsonPath> <markerPath>

# Example
SuspensionTools.exe hardpoints add coordinates.json Marker.SLDPRT
```

**Features:**
- Creates virtual marker parts at coordinate locations
- Applies color coding based on component type
- Creates coordinate systems in virtual parts
- Organizes components in "Hardpoints" folder
- Supports pose-based positioning for configuration updates

## Integration with Existing System

### Hardpoint System Integration

The enhanced coordinate insertion integrates seamlessly with the existing hardpoint system:

1. **JSON Data Format:** Compatible with existing JSON structure
2. **Color Coding:** Maintains existing color mapping
3. **Folder Organization:** Creates separate "Coordinates" and "Hardpoints" folders
4. **Progress Reporting:** Enhanced progress and state reporting
5. **Error Handling:** Improved error handling and recovery

### Backward Compatibility

- Existing coordinate insertion commands continue to work
- Enhanced features are additive, not replacing
- Can be enabled/disabled as needed
- Maintains compatibility with existing workflows

## Usage Examples

### Creating Individual Coordinate Systems

```csharp
// Using the enhanced coordinate system manager
bool success = CoordinateSystemManager.CreatePoseCoordinateSystem(
    swApp, "Test_CS", 100.0, 200.0, 300.0, 0.0, 0.0, 0.0);
```

### Batch Coordinate Creation from JSON

```csharp
// Load JSON data
var jsonData = LoadJson("coordinates.json");
var hardpoints = ExtractHardpoints(jsonData);

// Create coordinate systems
foreach (var hardpoint in hardpoints)
{
    CoordinateSystemManager.CreatePoseCoordinateSystem(
        swApp, hardpoint.Name, hardpoint.X, hardpoint.Y, hardpoint.Z);
}
```

### Virtual Part Editing

```csharp
// Get virtual parts
var virtualParts = VirtualPartEditor.GetVirtualParts(assemblyDoc);

// Batch edit virtual parts
VirtualPartEditor.BatchEditVirtualParts(swApp, virtualParts, (partDoc, component) =>
{
    // Add coordinate system to each virtual part
    VirtualPartEditor.AddCoordinateSystemToVirtualPart(
        swApp, component, "CS_Local", 0, 0, 0);
});
```

## Testing and Validation

### Test Suite

The `test_coordinate_insertion.py` script provides comprehensive testing:

1. **Enhanced Features Test:** Validates all new functionality
2. **Coordinate System Manager Test:** Tests pose-based positioning
3. **Virtual Part Editor Test:** Tests virtual part editing capabilities
4. **Command Line Interface Test:** Tests CLI functionality
5. **Integration Test:** End-to-end workflow testing

### Running Tests

```bash
# Run the test suite
python test_coordinate_insertion.py

# Expected output:
# ✓ All enhanced feature files are present
# ✓ CoordinateSystemManager imported successfully
# ✓ VirtualPartEditor imported successfully
# ✓ Found SuspensionTools.exe
# Results: 5/5 tests passed
```

## Build and Deployment

### Building the Project

```bash
# Build the project
cd sw_drawer
dotnet build -c Release

# Verify the executable
ls bin/Release/net48/SuspensionTools.exe
```

### Deployment

1. **Copy Executable:** Copy `SuspensionTools.exe` to the plugin directory
2. **Verify Dependencies:** Ensure SolidWorks API references are correct
3. **Test Integration:** Run the test suite to verify functionality
4. **Update Documentation:** Update user documentation with new features

## Troubleshooting

### Common Issues

1. **SolidWorks API Not Found:**
   - Ensure SolidWorks is installed
   - Check API path configuration in `.csproj`
   - Verify registry entries for SolidWorks installation

2. **Coordinate Systems Not Created:**
   - Check that assembly document is active
   - Verify coordinate system names are unique
   - Check for sufficient permissions

3. **Virtual Part Editing Fails:**
   - Ensure virtual parts exist in assembly
   - Check that edit component mode can be entered
   - Verify part document access

### Error Messages

- **"No active document found":** Ensure SolidWorks is running with an active document
- **"Active document must be an assembly":** Coordinate insertion requires assembly documents
- **"Failed to create coordinate system":** Check coordinate system name and parameters
- **"Could not access part document":** Virtual part editing requires proper permissions

## Future Enhancements

### Planned Features

1. **Configuration Management:** Support for multiple configurations with different poses
2. **Coordinate System Templates:** Pre-defined coordinate system templates
3. **Advanced Querying:** Enhanced coordinate system querying and filtering
4. **Integration with Simulation:** Direct integration with simulation workflows
5. **User Interface:** Enhanced UI for coordinate system management

### API Extensions

1. **Coordinate System Validation:** Validate coordinate system parameters
2. **Batch Operations:** Enhanced batch operations for coordinate systems
3. **Export/Import:** Export coordinate systems to various formats
4. **Version Control:** Integration with version control systems

## Conclusion

The enhanced coordinate insertion functionality provides significant improvements to the SolidWorks Plugin:

- **Better Organization:** Coordinate systems at assembly origin with proper folder management
- **Flexible Positioning:** Pose-based positioning for dynamic configurations
- **Enhanced Editing:** Virtual part editing capabilities
- **Improved Integration:** Better integration with existing workflows
- **Comprehensive Testing:** Full test suite for validation

These enhancements make the coordinate insertion system more robust, flexible, and user-friendly while maintaining backward compatibility with existing functionality.