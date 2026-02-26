# Task Progress Checklist

## Coordinate Insertion
- [x] Create "coordinates" folder (already exists)
- [x] Create virtual Marker.sldprt part (via InsertMarker.cs)
- [x] Rename virtual part to match JSON name schema 
- [x] Apply color coding based on name schema
- [ ] Rename internal coordinate system to match
- [x] Implement progress bar with state display (see InsertMarker.cs states)
- [x] Implement abort functionality (worker thread support)

## Pose Creation
- [ ] Create UI for pose name input
- [ ] Create "<posename> Transforms" folder
- [ ] Create coordinate systems named "<posename> <coordinate name>"
- [ ] Create coincident mates to align virtual parts
- [ ] Hide pose coordinate systems by default
- [ ] Implement progress bar with state display
- [ ] Implement abort functionality

## Visualization Controls
- [ ] Create "Show/Hide All" toggle
- [ ] Create "Show/Hide Front" toggle
- [ ] Create "Show/Hide Rear" toggle
- [ ] Implement color-coded category buttons
- [ ] Connect toggle functionality to SolidWorks API