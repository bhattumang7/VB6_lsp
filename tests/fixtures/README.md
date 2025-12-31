# Test Fixtures

This directory contains test data files used for integration testing.

## Files

### Game2048.RES
Real VB6 resource file from a Game 2048 project containing:
- 16 bitmap resources (IDs 101-113, 199-201)
- US English language (0x0409)
- Total size: ~200KB

### sheep.res
Real VB6 resource file containing:
- 3 Icon resources
- 1 GroupIcon resource
- 10 WAVE (audio) resources
- 8 CUSTOM resources
- 24 Bitmap resources
- Total: 46 resource entries
- Total size: ~759KB

## Usage

Both files are used to test:
- Win32 .res file parsing
- Resource header reading
- DWORD alignment handling
- Binary data preservation
- Round-trip read/write accuracy
- Different resource type handling

## Generated Files

Test runs may create temporary files:
- `*_copy.res` - Round-trip test outputs (ignored by git)
