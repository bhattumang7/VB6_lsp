//! VB6 Control Property Definitions
//!
//! Comprehensive property definitions for all VB6 form controls.
//! Based on vb6parse-master property definitions.

/// Property definition
#[derive(Debug, Clone)]
pub struct PropertyDef {
    /// Property name
    pub name: &'static str,
    /// Property description
    pub description: &'static str,
    /// Property type
    pub property_type: PropertyType,
    /// Whether the property is read-only at runtime
    pub read_only: bool,
    /// Default value (if known)
    pub default_value: Option<&'static str>,
    /// Valid values for enumerated properties
    pub valid_values: &'static [PropertyValue],
}

/// Property type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropertyType {
    String,
    Integer,
    Long,
    Single,
    Double,
    Boolean,
    Color,
    Font,
    Picture,
    Enum,
    Object,
    Variant,
    Currency,
    Date,
}

impl PropertyType {
    /// Get VB6 type name
    pub fn vb6_type(&self) -> &'static str {
        match self {
            PropertyType::String => "String",
            PropertyType::Integer => "Integer",
            PropertyType::Long => "Long",
            PropertyType::Single => "Single",
            PropertyType::Double => "Double",
            PropertyType::Boolean => "Boolean",
            PropertyType::Color => "OLE_COLOR",
            PropertyType::Font => "StdFont",
            PropertyType::Picture => "StdPicture",
            PropertyType::Enum => "Integer",
            PropertyType::Object => "Object",
            PropertyType::Variant => "Variant",
            PropertyType::Currency => "Currency",
            PropertyType::Date => "Date",
        }
    }
}

/// Valid value for enumerated properties
#[derive(Debug, Clone)]
pub struct PropertyValue {
    /// Numeric value
    pub value: i32,
    /// Symbolic name (e.g., "vbLeftJustify")
    pub name: &'static str,
    /// Description
    pub description: &'static str,
}

// =============================================================================
// Common Properties (shared across multiple controls)
// =============================================================================

/// Common appearance property values
pub static APPEARANCE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "Flat", description: "Flat appearance" },
    PropertyValue { value: 1, name: "3D", description: "3D appearance (default)" },
];

/// Common border style values (for most controls)
pub static BORDERSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "None", description: "No border" },
    PropertyValue { value: 1, name: "Fixed Single", description: "Fixed single-line border" },
];

/// Form border style values
pub static FORM_BORDERSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "None", description: "No border or title bar" },
    PropertyValue { value: 1, name: "Fixed Single", description: "Fixed single-line border, not resizable" },
    PropertyValue { value: 2, name: "Sizable", description: "Sizable border (default)" },
    PropertyValue { value: 3, name: "Fixed Dialog", description: "Fixed dialog border, not resizable" },
    PropertyValue { value: 4, name: "Fixed ToolWindow", description: "Fixed tool window border" },
    PropertyValue { value: 5, name: "Sizable ToolWindow", description: "Sizable tool window border" },
];

/// Alignment values
pub static ALIGNMENT_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbLeftJustify", description: "Left align (default)" },
    PropertyValue { value: 1, name: "vbRightJustify", description: "Right align" },
    PropertyValue { value: 2, name: "vbCenter", description: "Center" },
];

/// Mouse pointer values
pub static MOUSEPOINTER_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbDefault", description: "Default cursor" },
    PropertyValue { value: 1, name: "vbArrow", description: "Arrow cursor" },
    PropertyValue { value: 2, name: "vbCrosshair", description: "Crosshair cursor" },
    PropertyValue { value: 3, name: "vbIBeam", description: "I-beam cursor" },
    PropertyValue { value: 4, name: "vbIconPointer", description: "Icon cursor" },
    PropertyValue { value: 5, name: "vbSizePointer", description: "Size cursor" },
    PropertyValue { value: 6, name: "vbSizeNESW", description: "NE-SW size cursor" },
    PropertyValue { value: 7, name: "vbSizeNS", description: "N-S size cursor" },
    PropertyValue { value: 8, name: "vbSizeNWSE", description: "NW-SE size cursor" },
    PropertyValue { value: 9, name: "vbSizeWE", description: "W-E size cursor" },
    PropertyValue { value: 10, name: "vbUpArrow", description: "Up arrow cursor" },
    PropertyValue { value: 11, name: "vbHourglass", description: "Hourglass cursor" },
    PropertyValue { value: 12, name: "vbNoDrop", description: "No drop cursor" },
    PropertyValue { value: 13, name: "vbArrowHourglass", description: "Arrow with hourglass" },
    PropertyValue { value: 14, name: "vbArrowQuestion", description: "Arrow with question mark" },
    PropertyValue { value: 15, name: "vbSizeAll", description: "Size all cursor" },
    PropertyValue { value: 99, name: "vbCustom", description: "Custom cursor (use MouseIcon)" },
];

/// Scroll bars values
pub static SCROLLBARS_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "None", description: "No scroll bars" },
    PropertyValue { value: 1, name: "Horizontal", description: "Horizontal scroll bar only" },
    PropertyValue { value: 2, name: "Vertical", description: "Vertical scroll bar only" },
    PropertyValue { value: 3, name: "Both", description: "Both horizontal and vertical scroll bars" },
];

/// Window state values
pub static WINDOWSTATE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbNormal", description: "Normal window" },
    PropertyValue { value: 1, name: "vbMinimized", description: "Minimized" },
    PropertyValue { value: 2, name: "vbMaximized", description: "Maximized" },
];

/// Start up position values
pub static STARTUPPOSITION_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbStartUpManual", description: "No initial setting" },
    PropertyValue { value: 1, name: "vbStartUpOwner", description: "Center on owner form" },
    PropertyValue { value: 2, name: "vbStartUpScreen", description: "Center on screen" },
    PropertyValue { value: 3, name: "vbStartUpWindowsDefault", description: "Windows default position" },
];

/// Drag mode values
pub static DRAGMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbManual", description: "Manual drag mode" },
    PropertyValue { value: 1, name: "vbAutomatic", description: "Automatic drag mode" },
];

/// OLE drag mode values
pub static OLEDRAGMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbOLEDragManual", description: "Manual OLE drag" },
    PropertyValue { value: 1, name: "vbOLEDragAutomatic", description: "Automatic OLE drag" },
];

/// OLE drop mode values
pub static OLEDROPMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbOLEDropNone", description: "No OLE drop" },
    PropertyValue { value: 1, name: "vbOLEDropManual", description: "Manual OLE drop" },
    PropertyValue { value: 2, name: "vbOLEDropAutomatic", description: "Automatic OLE drop" },
];

/// Checkbox/option button value
pub static CHECKVALUE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbUnchecked", description: "Unchecked" },
    PropertyValue { value: 1, name: "vbChecked", description: "Checked" },
    PropertyValue { value: 2, name: "vbGrayed", description: "Grayed (indeterminate)" },
];

/// ComboBox style values
pub static COMBOSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbComboDropdown", description: "Dropdown combo (default)" },
    PropertyValue { value: 1, name: "vbComboSimple", description: "Simple combo with edit" },
    PropertyValue { value: 2, name: "vbComboDropdownList", description: "Dropdown list only" },
];

/// ListBox style values
pub static LISTBOXSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "Standard", description: "Standard list box" },
    PropertyValue { value: 1, name: "Checkbox", description: "List box with checkboxes" },
];

/// Draw mode values
pub static DRAWMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 1, name: "vbBlackness", description: "Blackness" },
    PropertyValue { value: 2, name: "vbNotMergePen", description: "Not merge pen" },
    PropertyValue { value: 3, name: "vbMaskNotPen", description: "Mask not pen" },
    PropertyValue { value: 4, name: "vbNotCopyPen", description: "Not copy pen" },
    PropertyValue { value: 5, name: "vbMaskPenNot", description: "Mask pen not" },
    PropertyValue { value: 6, name: "vbInvert", description: "Invert" },
    PropertyValue { value: 7, name: "vbXorPen", description: "XOR pen" },
    PropertyValue { value: 8, name: "vbNotMaskPen", description: "Not mask pen" },
    PropertyValue { value: 9, name: "vbMaskPen", description: "Mask pen" },
    PropertyValue { value: 10, name: "vbNotXorPen", description: "Not XOR pen" },
    PropertyValue { value: 11, name: "vbNop", description: "No operation" },
    PropertyValue { value: 12, name: "vbMergeNotPen", description: "Merge not pen" },
    PropertyValue { value: 13, name: "vbCopyPen", description: "Copy pen (default)" },
    PropertyValue { value: 14, name: "vbMergePenNot", description: "Merge pen not" },
    PropertyValue { value: 15, name: "vbMergePen", description: "Merge pen" },
    PropertyValue { value: 16, name: "vbWhiteness", description: "Whiteness" },
];

/// Draw style values
pub static DRAWSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbSolid", description: "Solid line" },
    PropertyValue { value: 1, name: "vbDash", description: "Dashed line" },
    PropertyValue { value: 2, name: "vbDot", description: "Dotted line" },
    PropertyValue { value: 3, name: "vbDashDot", description: "Dash-dot line" },
    PropertyValue { value: 4, name: "vbDashDotDot", description: "Dash-dot-dot line" },
    PropertyValue { value: 5, name: "vbInvisible", description: "Invisible line" },
    PropertyValue { value: 6, name: "vbInsideSolid", description: "Inside solid line" },
];

/// Fill style values
pub static FILLSTYLE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbFSSolid", description: "Solid fill" },
    PropertyValue { value: 1, name: "vbFSTransparent", description: "Transparent (default)" },
    PropertyValue { value: 2, name: "vbHorizontalLine", description: "Horizontal lines" },
    PropertyValue { value: 3, name: "vbVerticalLine", description: "Vertical lines" },
    PropertyValue { value: 4, name: "vbUpwardDiagonal", description: "Upward diagonal lines" },
    PropertyValue { value: 5, name: "vbDownwardDiagonal", description: "Downward diagonal lines" },
    PropertyValue { value: 6, name: "vbCross", description: "Cross pattern" },
    PropertyValue { value: 7, name: "vbDiagonalCross", description: "Diagonal cross pattern" },
];

/// Scale mode values
pub static SCALEMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbUser", description: "User-defined scale" },
    PropertyValue { value: 1, name: "vbTwips", description: "Twips (default)" },
    PropertyValue { value: 2, name: "vbPoints", description: "Points" },
    PropertyValue { value: 3, name: "vbPixels", description: "Pixels" },
    PropertyValue { value: 4, name: "vbCharacters", description: "Characters" },
    PropertyValue { value: 5, name: "vbInches", description: "Inches" },
    PropertyValue { value: 6, name: "vbMillimeters", description: "Millimeters" },
    PropertyValue { value: 7, name: "vbCentimeters", description: "Centimeters" },
];

/// Shape control shape values
pub static SHAPE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbShapeRectangle", description: "Rectangle" },
    PropertyValue { value: 1, name: "vbShapeSquare", description: "Square" },
    PropertyValue { value: 2, name: "vbShapeOval", description: "Oval" },
    PropertyValue { value: 3, name: "vbShapeCircle", description: "Circle" },
    PropertyValue { value: 4, name: "vbShapeRoundedRectangle", description: "Rounded rectangle" },
    PropertyValue { value: 5, name: "vbShapeRoundedSquare", description: "Rounded square" },
];

/// Picture size mode values
pub static STRETCHMODE_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "vbClip", description: "Clip image" },
    PropertyValue { value: 1, name: "vbStretch", description: "Stretch image to fit" },
];

/// Boolean values (for reference)
pub static BOOLEAN_VALUES: &[PropertyValue] = &[
    PropertyValue { value: 0, name: "False", description: "False" },
    PropertyValue { value: -1, name: "True", description: "True" },
];

// =============================================================================
// Form Properties
// =============================================================================

pub static FORM_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the form", property_type: PropertyType::String, read_only: true, default_value: Some("Form1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed in the title bar", property_type: PropertyType::String, read_only: false, default_value: Some("Form1"), valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets whether the form appears flat or 3D", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets the background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets the border style", property_type: PropertyType::Enum, read_only: true, default_value: Some("2"), valid_values: FORM_BORDERSTYLE_VALUES },
    PropertyDef { name: "ControlBox", description: "Returns/sets whether the control box is displayed", property_type: PropertyType::Boolean, read_only: true, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the form responds to user events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font used for text", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets the foreground color for text", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets the height of the form", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Icon", description: "Returns/sets the icon displayed when minimized", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "KeyPreview", description: "Returns/sets whether form receives key events before controls", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Left", description: "Returns/sets the left edge position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MaxButton", description: "Returns/sets whether the maximize button is displayed", property_type: PropertyType::Boolean, read_only: true, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "MDIChild", description: "Returns/sets whether the form is an MDI child", property_type: PropertyType::Boolean, read_only: true, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "MinButton", description: "Returns/sets whether the minimize button is displayed", property_type: PropertyType::Boolean, read_only: true, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "MousePointer", description: "Returns/sets the mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Moveable", description: "Returns/sets whether the form can be moved", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Picture", description: "Returns/sets the background picture", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ScaleMode", description: "Returns/sets the scale mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: SCALEMODE_VALUES },
    PropertyDef { name: "ShowInTaskbar", description: "Returns/sets whether the form appears in the taskbar", property_type: PropertyType::Boolean, read_only: true, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "StartUpPosition", description: "Returns/sets the initial position", property_type: PropertyType::Enum, read_only: false, default_value: Some("3"), valid_values: STARTUPPOSITION_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets a user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets the top edge position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets whether the form is visible", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets the width of the form", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "WindowState", description: "Returns/sets the window state", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: WINDOWSTATE_VALUES },
    PropertyDef { name: "AutoRedraw", description: "Returns/sets whether graphics are redrawn automatically", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "ClipControls", description: "Returns/sets whether graphics methods repaint entire form", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "DrawMode", description: "Returns/sets the drawing mode for graphics methods", property_type: PropertyType::Enum, read_only: false, default_value: Some("13"), valid_values: DRAWMODE_VALUES },
    PropertyDef { name: "DrawStyle", description: "Returns/sets the line style for graphics methods", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: DRAWSTYLE_VALUES },
    PropertyDef { name: "DrawWidth", description: "Returns/sets the line width for graphics methods", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "FillColor", description: "Returns/sets the fill color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "FillStyle", description: "Returns/sets the fill style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: FILLSTYLE_VALUES },
    PropertyDef { name: "hWnd", description: "Returns the window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "hDC", description: "Returns the device context handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// MDIForm Properties
// =============================================================================

pub static MDIFORM_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the MDI form", property_type: PropertyType::String, read_only: true, default_value: Some("MDIForm1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed in the title bar", property_type: PropertyType::String, read_only: false, default_value: Some("MDIForm1"), valid_values: &[] },
    PropertyDef { name: "BackColor", description: "Returns/sets the background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000C&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the form responds to user events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Height", description: "Returns/sets the height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Icon", description: "Returns/sets the icon", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets the left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Moveable", description: "Returns/sets whether the form can be moved", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Picture", description: "Returns/sets the background picture", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ScrollBars", description: "Returns/sets the scroll bar visibility", property_type: PropertyType::Boolean, read_only: true, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets a user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets the top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets the width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "WindowState", description: "Returns/sets the window state", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: WINDOWSTATE_VALUES },
    PropertyDef { name: "hWnd", description: "Returns the window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// TextBox Properties
// =============================================================================

pub static TEXTBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Text1"), valid_values: &[] },
    PropertyDef { name: "Text", description: "Returns/sets the text content", property_type: PropertyType::String, read_only: false, default_value: Some("Text1"), valid_values: &[] },
    PropertyDef { name: "Alignment", description: "Returns/sets text alignment", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: ALIGNMENT_VALUES },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: true, default_value: Some("1"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "HideSelection", description: "Returns/sets whether selection is hidden when control loses focus", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Locked", description: "Returns/sets whether the text can be edited", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "MaxLength", description: "Returns/sets maximum text length (0 = no limit)", property_type: PropertyType::Long, read_only: false, default_value: Some("0"), valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "MultiLine", description: "Returns/sets whether control supports multiple lines", property_type: PropertyType::Boolean, read_only: true, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "PasswordChar", description: "Returns/sets character to display for passwords", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "ScrollBars", description: "Returns/sets scroll bar visibility", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: SCROLLBARS_VALUES },
    PropertyDef { name: "SelLength", description: "Returns/sets length of selected text", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "SelStart", description: "Returns/sets starting position of selected text", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "SelText", description: "Returns/sets selected text", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// Label Properties
// =============================================================================

pub static LABEL_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Label1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed", property_type: PropertyType::String, read_only: false, default_value: Some("Label1"), valid_values: &[] },
    PropertyDef { name: "Alignment", description: "Returns/sets text alignment", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: ALIGNMENT_VALUES },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "AutoSize", description: "Returns/sets whether the control resizes to fit its contents", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BackStyle", description: "Returns/sets background style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: &[
        PropertyValue { value: 0, name: "Transparent", description: "Transparent background" },
        PropertyValue { value: 1, name: "Opaque", description: "Opaque background" },
    ] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "UseMnemonic", description: "Returns/sets whether & creates access key", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "WordWrap", description: "Returns/sets whether text wraps", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
];

// =============================================================================
// CommandButton Properties
// =============================================================================

pub static COMMANDBUTTON_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Command1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed on the button", property_type: PropertyType::String, read_only: false, default_value: Some("Command1"), valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color (Style must be Graphical)", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "Cancel", description: "Returns/sets whether Escape key triggers Click", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Default", description: "Returns/sets whether Enter key triggers Click", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the button responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Picture", description: "Returns/sets the picture (Style must be Graphical)", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Style", description: "Returns/sets button style", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "Standard", description: "Standard Windows button" },
        PropertyValue { value: 1, name: "Graphical", description: "Button with picture" },
    ] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// CheckBox Properties
// =============================================================================

pub static CHECKBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Check1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed", property_type: PropertyType::String, read_only: false, default_value: Some("Check1"), valid_values: &[] },
    PropertyDef { name: "Value", description: "Returns/sets the check state", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: CHECKVALUE_VALUES },
    PropertyDef { name: "Alignment", description: "Returns/sets text alignment", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "LeftJustify", description: "Checkbox left, caption right" },
        PropertyValue { value: 1, name: "RightJustify", description: "Checkbox right, caption left" },
    ] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Style", description: "Returns/sets checkbox style", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "Standard", description: "Standard checkbox" },
        PropertyValue { value: 1, name: "Graphical", description: "Button-style checkbox" },
    ] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// OptionButton Properties
// =============================================================================

pub static OPTIONBUTTON_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Option1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed", property_type: PropertyType::String, read_only: false, default_value: Some("Option1"), valid_values: &[] },
    PropertyDef { name: "Value", description: "Returns/sets whether the option is selected", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Alignment", description: "Returns/sets text alignment", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "LeftJustify", description: "Button left, caption right" },
        PropertyValue { value: 1, name: "RightJustify", description: "Button right, caption left" },
    ] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Style", description: "Returns/sets option button style", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "Standard", description: "Standard option button" },
        PropertyValue { value: 1, name: "Graphical", description: "Button-style option" },
    ] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// Frame Properties
// =============================================================================

pub static FRAME_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Frame1"), valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the text displayed in the frame border", property_type: PropertyType::String, read_only: false, default_value: Some("Frame1"), valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the frame and its controls respond to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font for the caption", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// ListBox Properties
// =============================================================================

pub static LISTBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("List1"), valid_values: &[] },
    PropertyDef { name: "List", description: "Returns/sets the items in the list", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListIndex", description: "Returns/sets the index of the selected item (-1 = none)", property_type: PropertyType::Integer, read_only: false, default_value: Some("-1"), valid_values: &[] },
    PropertyDef { name: "ListCount", description: "Returns the number of items", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "Text", description: "Returns the text of the selected item", property_type: PropertyType::String, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "Columns", description: "Returns/sets the number of columns", property_type: PropertyType::Integer, read_only: false, default_value: Some("0"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "IntegralHeight", description: "Returns/sets whether to show only complete rows", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "ItemData", description: "Returns/sets data associated with each item", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "MultiSelect", description: "Returns/sets multi-selection mode", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "None", description: "Single selection only" },
        PropertyValue { value: 1, name: "Simple", description: "Click toggles selection" },
        PropertyValue { value: 2, name: "Extended", description: "Shift+Click for range selection" },
    ] },
    PropertyDef { name: "SelCount", description: "Returns the number of selected items", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "Selected", description: "Returns/sets whether an item is selected", property_type: PropertyType::Boolean, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Sorted", description: "Returns/sets whether items are sorted", property_type: PropertyType::Boolean, read_only: true, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Style", description: "Returns/sets list box style", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: LISTBOXSTYLE_VALUES },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TopIndex", description: "Returns/sets index of topmost visible item", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// ComboBox Properties
// =============================================================================

pub static COMBOBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Combo1"), valid_values: &[] },
    PropertyDef { name: "Text", description: "Returns/sets the text in the edit area", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "List", description: "Returns/sets the items in the list", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListIndex", description: "Returns/sets the index of the selected item", property_type: PropertyType::Integer, read_only: false, default_value: Some("-1"), valid_values: &[] },
    PropertyDef { name: "ListCount", description: "Returns the number of items", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "Style", description: "Returns/sets combo box style", property_type: PropertyType::Enum, read_only: true, default_value: Some("0"), valid_values: COMBOSTYLE_VALUES },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "IntegralHeight", description: "Returns/sets whether to show only complete rows", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "ItemData", description: "Returns/sets data associated with each item", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Locked", description: "Returns/sets whether the text can be edited", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "SelLength", description: "Returns/sets length of selected text", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "SelStart", description: "Returns/sets starting position of selected text", property_type: PropertyType::Long, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "SelText", description: "Returns/sets selected text", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Sorted", description: "Returns/sets whether items are sorted", property_type: PropertyType::Boolean, read_only: true, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TopIndex", description: "Returns/sets index of topmost visible item", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// PictureBox Properties
// =============================================================================

pub static PICTUREBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Picture1"), valid_values: &[] },
    PropertyDef { name: "Picture", description: "Returns/sets the picture displayed", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "AutoRedraw", description: "Returns/sets whether graphics are persistent", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "AutoSize", description: "Returns/sets whether control resizes to fit picture", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "DrawMode", description: "Returns/sets drawing mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("13"), valid_values: DRAWMODE_VALUES },
    PropertyDef { name: "DrawStyle", description: "Returns/sets line style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: DRAWSTYLE_VALUES },
    PropertyDef { name: "DrawWidth", description: "Returns/sets line width", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "FillColor", description: "Returns/sets fill color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "FillStyle", description: "Returns/sets fill style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: FILLSTYLE_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000012&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "ScaleMode", description: "Returns/sets scale mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: SCALEMODE_VALUES },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "hDC", description: "Returns device context handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// Image Properties
// =============================================================================

pub static IMAGE_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Image1"), valid_values: &[] },
    PropertyDef { name: "Picture", description: "Returns/sets the picture displayed", property_type: PropertyType::Picture, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "Stretch", description: "Returns/sets whether picture stretches to fit", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
];

// =============================================================================
// Timer Properties
// =============================================================================

pub static TIMER_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Timer1"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the timer is running", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Interval", description: "Returns/sets the interval in milliseconds (0 = disabled)", property_type: PropertyType::Long, read_only: false, default_value: Some("0"), valid_values: &[] },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
];

// =============================================================================
// ScrollBar Properties (shared by HScrollBar and VScrollBar)
// =============================================================================

pub static SCROLLBAR_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("HScroll1"), valid_values: &[] },
    PropertyDef { name: "Value", description: "Returns/sets the current position", property_type: PropertyType::Integer, read_only: false, default_value: Some("0"), valid_values: &[] },
    PropertyDef { name: "Min", description: "Returns/sets the minimum value", property_type: PropertyType::Integer, read_only: false, default_value: Some("0"), valid_values: &[] },
    PropertyDef { name: "Max", description: "Returns/sets the maximum value", property_type: PropertyType::Integer, read_only: false, default_value: Some("32767"), valid_values: &[] },
    PropertyDef { name: "SmallChange", description: "Returns/sets the amount changed when clicking arrows", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "LargeChange", description: "Returns/sets the amount changed when clicking track", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MousePointer", description: "Returns/sets mouse pointer style", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: MOUSEPOINTER_VALUES },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// Shape Properties
// =============================================================================

pub static SHAPE_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Shape1"), valid_values: &[] },
    PropertyDef { name: "Shape", description: "Returns/sets the shape type", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: SHAPE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "BackStyle", description: "Returns/sets background style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: &[
        PropertyValue { value: 0, name: "Transparent", description: "Transparent background" },
        PropertyValue { value: 1, name: "Opaque", description: "Opaque background" },
    ] },
    PropertyDef { name: "BorderColor", description: "Returns/sets border color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: &[
        PropertyValue { value: 0, name: "Transparent", description: "No border" },
        PropertyValue { value: 1, name: "Solid", description: "Solid border" },
        PropertyValue { value: 2, name: "Dash", description: "Dashed border" },
        PropertyValue { value: 3, name: "Dot", description: "Dotted border" },
        PropertyValue { value: 4, name: "DashDot", description: "Dash-dot border" },
        PropertyValue { value: 5, name: "DashDotDot", description: "Dash-dot-dot border" },
        PropertyValue { value: 6, name: "InsideSolid", description: "Inside solid border" },
    ] },
    PropertyDef { name: "BorderWidth", description: "Returns/sets border width", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "DrawMode", description: "Returns/sets drawing mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("13"), valid_values: DRAWMODE_VALUES },
    PropertyDef { name: "FillColor", description: "Returns/sets fill color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "FillStyle", description: "Returns/sets fill style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: FILLSTYLE_VALUES },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
];

// =============================================================================
// Line Properties
// =============================================================================

pub static LINE_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Line1"), valid_values: &[] },
    PropertyDef { name: "BorderColor", description: "Returns/sets line color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets line style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: &[
        PropertyValue { value: 0, name: "Transparent", description: "No line" },
        PropertyValue { value: 1, name: "Solid", description: "Solid line" },
        PropertyValue { value: 2, name: "Dash", description: "Dashed line" },
        PropertyValue { value: 3, name: "Dot", description: "Dotted line" },
        PropertyValue { value: 4, name: "DashDot", description: "Dash-dot line" },
        PropertyValue { value: 5, name: "DashDotDot", description: "Dash-dot-dot line" },
        PropertyValue { value: 6, name: "InsideSolid", description: "Inside solid line" },
    ] },
    PropertyDef { name: "BorderWidth", description: "Returns/sets line width", property_type: PropertyType::Integer, read_only: false, default_value: Some("1"), valid_values: &[] },
    PropertyDef { name: "DrawMode", description: "Returns/sets drawing mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("13"), valid_values: DRAWMODE_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "X1", description: "Returns/sets X coordinate of start point", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "X2", description: "Returns/sets X coordinate of end point", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Y1", description: "Returns/sets Y coordinate of start point", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Y2", description: "Returns/sets Y coordinate of end point", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
];

// =============================================================================
// Data Properties
// =============================================================================

pub static DATA_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Data1"), valid_values: &[] },
    PropertyDef { name: "DatabaseName", description: "Returns/sets the database path", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "RecordSource", description: "Returns/sets the table or query", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the caption text", property_type: PropertyType::String, read_only: false, default_value: Some("Data1"), valid_values: &[] },
    PropertyDef { name: "Connect", description: "Returns/sets the database connect string", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BOFAction", description: "Returns/sets action at beginning of file", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "vbMoveFirst", description: "Move to first record" },
        PropertyValue { value: 1, name: "vbBOF", description: "Stay at BOF" },
    ] },
    PropertyDef { name: "EOFAction", description: "Returns/sets action at end of file", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "vbMoveLast", description: "Move to last record" },
        PropertyValue { value: 1, name: "vbEOF", description: "Stay at EOF" },
        PropertyValue { value: 2, name: "vbAddNew", description: "Add new record" },
    ] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Exclusive", description: "Returns/sets exclusive database access", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H00000000&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Options", description: "Returns/sets database options", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ReadOnly", description: "Returns/sets read-only mode", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "RecordsetType", description: "Returns/sets the recordset type", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: &[
        PropertyValue { value: 0, name: "vbRSTypeTable", description: "Table-type recordset" },
        PropertyValue { value: 1, name: "vbRSTypeDynaset", description: "Dynaset-type recordset" },
        PropertyValue { value: 2, name: "vbRSTypeSnapShot", description: "Snapshot-type recordset" },
    ] },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// OLE Properties
// =============================================================================

pub static OLE_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("OLE1"), valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "AutoActivate", description: "Returns/sets when object activates", property_type: PropertyType::Enum, read_only: false, default_value: Some("2"), valid_values: &[
        PropertyValue { value: 0, name: "vbOLEActivateManual", description: "Manual activation" },
        PropertyValue { value: 1, name: "vbOLEActivateGetFocus", description: "Activate on focus" },
        PropertyValue { value: 2, name: "vbOLEActivateDoubleclick", description: "Activate on double-click" },
        PropertyValue { value: 3, name: "vbOLEActivateAuto", description: "Automatic activation" },
    ] },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H8000000F&"), valid_values: &[] },
    PropertyDef { name: "BorderStyle", description: "Returns/sets border style", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: BORDERSTYLE_VALUES },
    PropertyDef { name: "Class", description: "Returns the OLE class name", property_type: PropertyType::String, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "DisplayType", description: "Returns/sets display type", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "vbOLEDisplayContent", description: "Display content" },
        PropertyValue { value: 1, name: "vbOLEDisplayIcon", description: "Display as icon" },
    ] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "OLEType", description: "Returns the OLE object type", property_type: PropertyType::Enum, read_only: true, default_value: None, valid_values: &[
        PropertyValue { value: 0, name: "vbOLELinked", description: "Linked object" },
        PropertyValue { value: 1, name: "vbOLEEmbedded", description: "Embedded object" },
        PropertyValue { value: 3, name: "vbOLENone", description: "No object" },
    ] },
    PropertyDef { name: "OLETypeAllowed", description: "Returns/sets allowed OLE type", property_type: PropertyType::Enum, read_only: false, default_value: Some("2"), valid_values: &[
        PropertyValue { value: 0, name: "vbOLELinked", description: "Linked only" },
        PropertyValue { value: 1, name: "vbOLEEmbedded", description: "Embedded only" },
        PropertyValue { value: 2, name: "vbOLEEither", description: "Either linked or embedded" },
    ] },
    PropertyDef { name: "SizeMode", description: "Returns/sets how object is displayed", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "vbOLESizeClip", description: "Clip to fit" },
        PropertyValue { value: 1, name: "vbOLESizeStretch", description: "Stretch to fit" },
        PropertyValue { value: 2, name: "vbOLESizeAutoSize", description: "Auto-size control" },
        PropertyValue { value: 3, name: "vbOLESizeZoom", description: "Zoom proportionally" },
    ] },
    PropertyDef { name: "SourceDoc", description: "Returns/sets source document path", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "SourceItem", description: "Returns/sets source item within document", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// FileListBox Properties
// =============================================================================

pub static FILELISTBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("File1"), valid_values: &[] },
    PropertyDef { name: "FileName", description: "Returns/sets the selected filename", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Path", description: "Returns/sets the current path", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Pattern", description: "Returns/sets the file filter pattern", property_type: PropertyType::String, read_only: false, default_value: Some("*.*"), valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "Archive", description: "Returns/sets whether to show archive files", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Hidden", description: "Returns/sets whether to show hidden files", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListCount", description: "Returns the number of files", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListIndex", description: "Returns/sets the selected index", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "MultiSelect", description: "Returns/sets multi-selection mode", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "None", description: "Single selection" },
        PropertyValue { value: 1, name: "Simple", description: "Simple multi-select" },
        PropertyValue { value: 2, name: "Extended", description: "Extended multi-select" },
    ] },
    PropertyDef { name: "Normal", description: "Returns/sets whether to show normal files", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "ReadOnly", description: "Returns/sets whether to show read-only files", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "System", description: "Returns/sets whether to show system files", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// DirListBox Properties
// =============================================================================

pub static DIRLISTBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Dir1"), valid_values: &[] },
    PropertyDef { name: "Path", description: "Returns/sets the current path", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListCount", description: "Returns the number of directories", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListIndex", description: "Returns/sets the selected index", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// DriveListBox Properties
// =============================================================================

pub static DRIVELISTBOX_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the control", property_type: PropertyType::String, read_only: true, default_value: Some("Drive1"), valid_values: &[] },
    PropertyDef { name: "Drive", description: "Returns/sets the current drive", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Appearance", description: "Returns/sets flat or 3D appearance", property_type: PropertyType::Enum, read_only: false, default_value: Some("1"), valid_values: APPEARANCE_VALUES },
    PropertyDef { name: "BackColor", description: "Returns/sets background color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000005&"), valid_values: &[] },
    PropertyDef { name: "Enabled", description: "Returns/sets whether the control responds to events", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Font", description: "Returns/sets the font", property_type: PropertyType::Font, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ForeColor", description: "Returns/sets foreground color", property_type: PropertyType::Color, read_only: false, default_value: Some("&H80000008&"), valid_values: &[] },
    PropertyDef { name: "Height", description: "Returns/sets height", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Left", description: "Returns/sets left position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListCount", description: "Returns the number of drives", property_type: PropertyType::Integer, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "ListIndex", description: "Returns/sets the selected index", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabIndex", description: "Returns/sets tab order", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "TabStop", description: "Returns/sets whether Tab stops on this control", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Top", description: "Returns/sets top position", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Width", description: "Returns/sets width", property_type: PropertyType::Single, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "hWnd", description: "Returns window handle", property_type: PropertyType::Long, read_only: true, default_value: None, valid_values: &[] },
];

// =============================================================================
// Menu Properties
// =============================================================================

pub static MENU_PROPERTIES: &[PropertyDef] = &[
    PropertyDef { name: "Name", description: "Returns the name of the menu item", property_type: PropertyType::String, read_only: true, default_value: None, valid_values: &[] },
    PropertyDef { name: "Caption", description: "Returns/sets the menu text (use & for access key, - for separator)", property_type: PropertyType::String, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Checked", description: "Returns/sets whether menu has checkmark", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Enabled", description: "Returns/sets whether menu is enabled", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "Index", description: "Returns/sets the index in control array", property_type: PropertyType::Integer, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "NegotiatePosition", description: "Returns/sets OLE menu position", property_type: PropertyType::Enum, read_only: false, default_value: Some("0"), valid_values: &[
        PropertyValue { value: 0, name: "vbMenuNegotiateNone", description: "Menu does not appear" },
        PropertyValue { value: 1, name: "vbMenuNegotiateLeft", description: "Menu on left" },
        PropertyValue { value: 2, name: "vbMenuNegotiateMiddle", description: "Menu in middle" },
        PropertyValue { value: 3, name: "vbMenuNegotiateRight", description: "Menu on right" },
    ] },
    PropertyDef { name: "Shortcut", description: "Returns/sets the keyboard shortcut", property_type: PropertyType::Enum, read_only: false, default_value: None, valid_values: &[] },
    PropertyDef { name: "Tag", description: "Returns/sets user-defined value", property_type: PropertyType::String, read_only: false, default_value: Some(""), valid_values: &[] },
    PropertyDef { name: "Visible", description: "Returns/sets visibility", property_type: PropertyType::Boolean, read_only: false, default_value: Some("True"), valid_values: BOOLEAN_VALUES },
    PropertyDef { name: "WindowList", description: "Returns/sets whether menu shows MDI child window list", property_type: PropertyType::Boolean, read_only: false, default_value: Some("False"), valid_values: BOOLEAN_VALUES },
];
