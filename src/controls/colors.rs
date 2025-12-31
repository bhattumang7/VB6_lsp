//! VB6 System Color Constants
//!
//! Provides parsing and display for VB6 color values including system colors.
//! Based on vb6parse-master color definitions.

use std::fmt;

/// VB6 system color constants
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SystemColor {
    /// &H80000000& - Scroll bar gray area
    ScrollBar,
    /// &H80000001& - Desktop background
    Background,
    /// &H80000002& - Active window title bar
    ActiveCaption,
    /// &H80000003& - Inactive window title bar
    InactiveCaption,
    /// &H80000004& - Menu background
    Menu,
    /// &H80000005& - Window background
    Window,
    /// &H80000006& - Window frame
    WindowFrame,
    /// &H80000007& - Menu text
    MenuText,
    /// &H80000008& - Window text
    WindowText,
    /// &H80000009& - Active title bar text
    CaptionText,
    /// &H8000000A& - Active window border
    ActiveBorder,
    /// &H8000000B& - Inactive window border
    InactiveBorder,
    /// &H8000000C& - MDI background
    AppWorkspace,
    /// &H8000000D& - Selected item background
    Highlight,
    /// &H8000000E& - Selected item text
    HighlightText,
    /// &H8000000F& - Button face (3D face)
    ButtonFace,
    /// &H80000010& - Button shadow (3D shadow)
    ButtonShadow,
    /// &H80000011& - Disabled text
    GrayText,
    /// &H80000012& - Button text
    ButtonText,
    /// &H80000013& - Inactive title bar text
    InactiveCaptionText,
    /// &H80000014& - Button highlight (3D highlight)
    ButtonHighlight,
    /// &H80000015& - 3D dark shadow
    DarkShadow3D,
    /// &H80000016& - 3D light
    Light3D,
    /// &H80000017& - Info (tooltip) text
    InfoText,
    /// &H80000018& - Info (tooltip) background
    InfoBackground,
}

impl SystemColor {
    /// Get the numeric value for this system color
    pub fn value(&self) -> u32 {
        match self {
            SystemColor::ScrollBar => 0x80000000,
            SystemColor::Background => 0x80000001,
            SystemColor::ActiveCaption => 0x80000002,
            SystemColor::InactiveCaption => 0x80000003,
            SystemColor::Menu => 0x80000004,
            SystemColor::Window => 0x80000005,
            SystemColor::WindowFrame => 0x80000006,
            SystemColor::MenuText => 0x80000007,
            SystemColor::WindowText => 0x80000008,
            SystemColor::CaptionText => 0x80000009,
            SystemColor::ActiveBorder => 0x8000000A,
            SystemColor::InactiveBorder => 0x8000000B,
            SystemColor::AppWorkspace => 0x8000000C,
            SystemColor::Highlight => 0x8000000D,
            SystemColor::HighlightText => 0x8000000E,
            SystemColor::ButtonFace => 0x8000000F,
            SystemColor::ButtonShadow => 0x80000010,
            SystemColor::GrayText => 0x80000011,
            SystemColor::ButtonText => 0x80000012,
            SystemColor::InactiveCaptionText => 0x80000013,
            SystemColor::ButtonHighlight => 0x80000014,
            SystemColor::DarkShadow3D => 0x80000015,
            SystemColor::Light3D => 0x80000016,
            SystemColor::InfoText => 0x80000017,
            SystemColor::InfoBackground => 0x80000018,
        }
    }

    /// Get the VB6 constant name (e.g., "vbScrollBars")
    pub fn vb6_name(&self) -> &'static str {
        match self {
            SystemColor::ScrollBar => "vbScrollBars",
            SystemColor::Background => "vbDesktop",
            SystemColor::ActiveCaption => "vbActiveTitleBar",
            SystemColor::InactiveCaption => "vbInactiveTitleBar",
            SystemColor::Menu => "vbMenuBar",
            SystemColor::Window => "vbWindowBackground",
            SystemColor::WindowFrame => "vbWindowFrame",
            SystemColor::MenuText => "vbMenuText",
            SystemColor::WindowText => "vbWindowText",
            SystemColor::CaptionText => "vbTitleBarText",
            SystemColor::ActiveBorder => "vbActiveBorder",
            SystemColor::InactiveBorder => "vbInactiveBorder",
            SystemColor::AppWorkspace => "vbApplicationWorkspace",
            SystemColor::Highlight => "vbHighlight",
            SystemColor::HighlightText => "vbHighlightText",
            SystemColor::ButtonFace => "vbButtonFace",
            SystemColor::ButtonShadow => "vbButtonShadow",
            SystemColor::GrayText => "vbGrayText",
            SystemColor::ButtonText => "vbButtonText",
            SystemColor::InactiveCaptionText => "vbInactiveCaptionText",
            SystemColor::ButtonHighlight => "vb3DHighlight",
            SystemColor::DarkShadow3D => "vb3DDKShadow",
            SystemColor::Light3D => "vb3DLight",
            SystemColor::InfoText => "vbInfoText",
            SystemColor::InfoBackground => "vbInfoBackground",
        }
    }

    /// Get a human-readable description
    pub fn description(&self) -> &'static str {
        match self {
            SystemColor::ScrollBar => "Scroll bar gray area",
            SystemColor::Background => "Desktop background",
            SystemColor::ActiveCaption => "Active window title bar",
            SystemColor::InactiveCaption => "Inactive window title bar",
            SystemColor::Menu => "Menu background",
            SystemColor::Window => "Window background",
            SystemColor::WindowFrame => "Window frame",
            SystemColor::MenuText => "Menu text",
            SystemColor::WindowText => "Window text",
            SystemColor::CaptionText => "Active title bar text",
            SystemColor::ActiveBorder => "Active window border",
            SystemColor::InactiveBorder => "Inactive window border",
            SystemColor::AppWorkspace => "MDI application background",
            SystemColor::Highlight => "Selected item background",
            SystemColor::HighlightText => "Selected item text",
            SystemColor::ButtonFace => "Button face / 3D surface",
            SystemColor::ButtonShadow => "Button shadow / 3D shadow",
            SystemColor::GrayText => "Disabled (grayed) text",
            SystemColor::ButtonText => "Button text",
            SystemColor::InactiveCaptionText => "Inactive title bar text",
            SystemColor::ButtonHighlight => "Button highlight / 3D highlight",
            SystemColor::DarkShadow3D => "3D dark shadow",
            SystemColor::Light3D => "3D light color",
            SystemColor::InfoText => "Tooltip text",
            SystemColor::InfoBackground => "Tooltip background",
        }
    }

    /// Parse a system color from its numeric value
    pub fn from_value(value: u32) -> Option<Self> {
        match value {
            0x80000000 => Some(SystemColor::ScrollBar),
            0x80000001 => Some(SystemColor::Background),
            0x80000002 => Some(SystemColor::ActiveCaption),
            0x80000003 => Some(SystemColor::InactiveCaption),
            0x80000004 => Some(SystemColor::Menu),
            0x80000005 => Some(SystemColor::Window),
            0x80000006 => Some(SystemColor::WindowFrame),
            0x80000007 => Some(SystemColor::MenuText),
            0x80000008 => Some(SystemColor::WindowText),
            0x80000009 => Some(SystemColor::CaptionText),
            0x8000000A => Some(SystemColor::ActiveBorder),
            0x8000000B => Some(SystemColor::InactiveBorder),
            0x8000000C => Some(SystemColor::AppWorkspace),
            0x8000000D => Some(SystemColor::Highlight),
            0x8000000E => Some(SystemColor::HighlightText),
            0x8000000F => Some(SystemColor::ButtonFace),
            0x80000010 => Some(SystemColor::ButtonShadow),
            0x80000011 => Some(SystemColor::GrayText),
            0x80000012 => Some(SystemColor::ButtonText),
            0x80000013 => Some(SystemColor::InactiveCaptionText),
            0x80000014 => Some(SystemColor::ButtonHighlight),
            0x80000015 => Some(SystemColor::DarkShadow3D),
            0x80000016 => Some(SystemColor::Light3D),
            0x80000017 => Some(SystemColor::InfoText),
            0x80000018 => Some(SystemColor::InfoBackground),
            _ => None,
        }
    }

    /// Get all system colors
    pub fn all() -> &'static [SystemColor] {
        &[
            SystemColor::ScrollBar,
            SystemColor::Background,
            SystemColor::ActiveCaption,
            SystemColor::InactiveCaption,
            SystemColor::Menu,
            SystemColor::Window,
            SystemColor::WindowFrame,
            SystemColor::MenuText,
            SystemColor::WindowText,
            SystemColor::CaptionText,
            SystemColor::ActiveBorder,
            SystemColor::InactiveBorder,
            SystemColor::AppWorkspace,
            SystemColor::Highlight,
            SystemColor::HighlightText,
            SystemColor::ButtonFace,
            SystemColor::ButtonShadow,
            SystemColor::GrayText,
            SystemColor::ButtonText,
            SystemColor::InactiveCaptionText,
            SystemColor::ButtonHighlight,
            SystemColor::DarkShadow3D,
            SystemColor::Light3D,
            SystemColor::InfoText,
            SystemColor::InfoBackground,
        ]
    }
}

impl fmt::Display for SystemColor {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{} ({})", self.vb6_name(), self.description())
    }
}

/// Represents a VB6 color value (either RGB or system color)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VB6Color {
    /// An RGB color value (0x00BBGGRR format - BGR, not RGB!)
    Rgb { red: u8, green: u8, blue: u8 },
    /// A system color constant
    System(SystemColor),
}

impl VB6Color {
    /// Parse a VB6 color string (e.g., "&H00FF0000&" or "&H80000005&")
    pub fn parse(s: &str) -> Option<Self> {
        let s = s.trim();

        // Handle VB6 hex format: &Hxxxxxxxx& or &Hxxxxxx
        let hex_str = if s.starts_with("&H") || s.starts_with("&h") {
            let s = s.trim_start_matches("&H").trim_start_matches("&h");
            s.trim_end_matches('&')
        } else if s.starts_with("0x") || s.starts_with("0X") {
            s.trim_start_matches("0x").trim_start_matches("0X")
        } else {
            // Try parsing as decimal
            if let Ok(value) = s.parse::<u32>() {
                return Self::from_u32(value);
            }
            return None;
        };

        // Parse hex value
        let value = u32::from_str_radix(hex_str, 16).ok()?;
        Self::from_u32(value)
    }

    /// Create from a u32 value
    pub fn from_u32(value: u32) -> Option<Self> {
        // Check if it's a system color (high bit set)
        if value & 0x80000000 != 0 {
            SystemColor::from_value(value).map(VB6Color::System)
        } else {
            // It's an RGB value (VB6 uses BGR format: 0x00BBGGRR)
            let red = (value & 0xFF) as u8;
            let green = ((value >> 8) & 0xFF) as u8;
            let blue = ((value >> 16) & 0xFF) as u8;
            Some(VB6Color::Rgb { red, green, blue })
        }
    }

    /// Convert to u32 value
    pub fn to_u32(&self) -> u32 {
        match self {
            VB6Color::Rgb { red, green, blue } => {
                (*red as u32) | ((*green as u32) << 8) | ((*blue as u32) << 16)
            }
            VB6Color::System(sys) => sys.value(),
        }
    }

    /// Get a description of this color
    pub fn description(&self) -> String {
        match self {
            VB6Color::Rgb { red, green, blue } => {
                format!("RGB({}, {}, {})", red, green, blue)
            }
            VB6Color::System(sys) => sys.to_string(),
        }
    }

    /// Format as VB6 hex string
    pub fn to_vb6_string(&self) -> String {
        format!("&H{:08X}&", self.to_u32())
    }

    /// Check if this is a system color
    pub fn is_system_color(&self) -> bool {
        matches!(self, VB6Color::System(_))
    }

    /// Get the system color if this is one
    pub fn as_system_color(&self) -> Option<&SystemColor> {
        match self {
            VB6Color::System(sys) => Some(sys),
            _ => None,
        }
    }

    /// Common VB6 color constants
    pub fn black() -> Self {
        VB6Color::Rgb { red: 0, green: 0, blue: 0 }
    }

    pub fn white() -> Self {
        VB6Color::Rgb { red: 255, green: 255, blue: 255 }
    }

    pub fn red() -> Self {
        VB6Color::Rgb { red: 255, green: 0, blue: 0 }
    }

    pub fn green() -> Self {
        VB6Color::Rgb { red: 0, green: 255, blue: 0 }
    }

    pub fn blue() -> Self {
        VB6Color::Rgb { red: 0, green: 0, blue: 255 }
    }

    pub fn yellow() -> Self {
        VB6Color::Rgb { red: 255, green: 255, blue: 0 }
    }

    pub fn cyan() -> Self {
        VB6Color::Rgb { red: 0, green: 255, blue: 255 }
    }

    pub fn magenta() -> Self {
        VB6Color::Rgb { red: 255, green: 0, blue: 255 }
    }
}

impl fmt::Display for VB6Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.description())
    }
}

/// VB6 color constants (for completion)
pub static VB6_COLOR_CONSTANTS: &[(&str, u32)] = &[
    ("vbBlack", 0x00000000),
    ("vbRed", 0x000000FF),
    ("vbGreen", 0x0000FF00),
    ("vbYellow", 0x0000FFFF),
    ("vbBlue", 0x00FF0000),
    ("vbMagenta", 0x00FF00FF),
    ("vbCyan", 0x00FFFF00),
    ("vbWhite", 0x00FFFFFF),
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_rgb_color() {
        let color = VB6Color::parse("&H00FF0000&").unwrap();
        match color {
            VB6Color::Rgb { red, green, blue } => {
                assert_eq!(red, 0);
                assert_eq!(green, 0);
                assert_eq!(blue, 255); // Blue in VB6 BGR format
            }
            _ => panic!("Expected RGB color"),
        }
    }

    #[test]
    fn test_parse_system_color() {
        let color = VB6Color::parse("&H80000005&").unwrap();
        match color {
            VB6Color::System(sys) => {
                assert_eq!(sys, SystemColor::Window);
            }
            _ => panic!("Expected system color"),
        }
    }

    #[test]
    fn test_system_color_description() {
        let sys = SystemColor::ButtonFace;
        assert_eq!(sys.vb6_name(), "vbButtonFace");
        assert!(sys.description().contains("Button"));
    }

    #[test]
    fn test_color_round_trip() {
        let original = "&H8000000F&";
        let color = VB6Color::parse(original).unwrap();
        assert_eq!(color.to_vb6_string(), original);
    }
}
