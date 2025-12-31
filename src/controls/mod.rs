//! VB6 Form Control Definitions
//!
//! Provides comprehensive property definitions for VB6 form controls,
//! enabling rich LSP features like hover info, completion, and validation.
//!
//! Based on vb6parse-master control definitions.

mod properties;
mod colors;
pub mod frx;

pub use colors::{SystemColor, VB6Color};
pub use properties::{PropertyDef, PropertyType, PropertyValue};

use std::collections::HashMap;
use once_cell::sync::Lazy;

/// Control definition with all its properties
#[derive(Debug, Clone)]
pub struct ControlDef {
    /// Control type name (e.g., "TextBox", "Label")
    pub name: &'static str,
    /// Full VB6 type name (e.g., "VB.TextBox")
    pub full_name: &'static str,
    /// Description for hover info
    pub description: &'static str,
    /// Properties available on this control
    pub properties: &'static [PropertyDef],
    /// Events available on this control
    pub events: &'static [EventDef],
    /// Methods available on this control
    pub methods: &'static [MethodDef],
    /// Whether this control can contain other controls
    pub is_container: bool,
}

/// Event definition
#[derive(Debug, Clone)]
pub struct EventDef {
    pub name: &'static str,
    pub description: &'static str,
    pub parameters: &'static str,
}

/// Method definition
#[derive(Debug, Clone)]
pub struct MethodDef {
    pub name: &'static str,
    pub description: &'static str,
    pub signature: &'static str,
    pub return_type: Option<&'static str>,
}

/// Menu keyboard shortcut
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MenuShortcut {
    None,
    CtrlA, CtrlB, CtrlC, CtrlD, CtrlE, CtrlF, CtrlG, CtrlH, CtrlI, CtrlJ,
    CtrlK, CtrlL, CtrlM, CtrlN, CtrlO, CtrlP, CtrlQ, CtrlR, CtrlS, CtrlT,
    CtrlU, CtrlV, CtrlW, CtrlX, CtrlY, CtrlZ,
    F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12,
    CtrlF1, CtrlF2, CtrlF3, CtrlF4, CtrlF5, CtrlF6, CtrlF7, CtrlF8, CtrlF9, CtrlF10, CtrlF11, CtrlF12,
    ShiftF1, ShiftF2, ShiftF3, ShiftF4, ShiftF5, ShiftF6, ShiftF7, ShiftF8, ShiftF9, ShiftF10, ShiftF11, ShiftF12,
    CtrlShiftF1, CtrlShiftF2, CtrlShiftF3, CtrlShiftF4, CtrlShiftF5, CtrlShiftF6,
    CtrlShiftF7, CtrlShiftF8, CtrlShiftF9, CtrlShiftF10, CtrlShiftF11, CtrlShiftF12,
    CtrlIns, ShiftIns, Del, ShiftDel,
    AltBksp, CtrlBksp,
}

impl MenuShortcut {
    /// Parse a VB6 shortcut string
    pub fn from_str(s: &str) -> Option<Self> {
        match s.trim() {
            "^A" => Some(Self::CtrlA), "^B" => Some(Self::CtrlB), "^C" => Some(Self::CtrlC),
            "^D" => Some(Self::CtrlD), "^E" => Some(Self::CtrlE), "^F" => Some(Self::CtrlF),
            "^G" => Some(Self::CtrlG), "^H" => Some(Self::CtrlH), "^I" => Some(Self::CtrlI),
            "^J" => Some(Self::CtrlJ), "^K" => Some(Self::CtrlK), "^L" => Some(Self::CtrlL),
            "^M" => Some(Self::CtrlM), "^N" => Some(Self::CtrlN), "^O" => Some(Self::CtrlO),
            "^P" => Some(Self::CtrlP), "^Q" => Some(Self::CtrlQ), "^R" => Some(Self::CtrlR),
            "^S" => Some(Self::CtrlS), "^T" => Some(Self::CtrlT), "^U" => Some(Self::CtrlU),
            "^V" => Some(Self::CtrlV), "^W" => Some(Self::CtrlW), "^X" => Some(Self::CtrlX),
            "^Y" => Some(Self::CtrlY), "^Z" => Some(Self::CtrlZ),
            "{F1}" => Some(Self::F1), "{F2}" => Some(Self::F2), "{F3}" => Some(Self::F3),
            "{F4}" => Some(Self::F4), "{F5}" => Some(Self::F5), "{F6}" => Some(Self::F6),
            "{F7}" => Some(Self::F7), "{F8}" => Some(Self::F8), "{F9}" => Some(Self::F9),
            "{F10}" => Some(Self::F10), "{F11}" => Some(Self::F11), "{F12}" => Some(Self::F12),
            _ => None,
        }
    }

    /// Get display string for the shortcut
    pub fn display(&self) -> &'static str {
        match self {
            Self::None => "",
            Self::CtrlA => "Ctrl+A", Self::CtrlB => "Ctrl+B", Self::CtrlC => "Ctrl+C",
            Self::CtrlD => "Ctrl+D", Self::CtrlE => "Ctrl+E", Self::CtrlF => "Ctrl+F",
            Self::CtrlG => "Ctrl+G", Self::CtrlH => "Ctrl+H", Self::CtrlI => "Ctrl+I",
            Self::CtrlJ => "Ctrl+J", Self::CtrlK => "Ctrl+K", Self::CtrlL => "Ctrl+L",
            Self::CtrlM => "Ctrl+M", Self::CtrlN => "Ctrl+N", Self::CtrlO => "Ctrl+O",
            Self::CtrlP => "Ctrl+P", Self::CtrlQ => "Ctrl+Q", Self::CtrlR => "Ctrl+R",
            Self::CtrlS => "Ctrl+S", Self::CtrlT => "Ctrl+T", Self::CtrlU => "Ctrl+U",
            Self::CtrlV => "Ctrl+V", Self::CtrlW => "Ctrl+W", Self::CtrlX => "Ctrl+X",
            Self::CtrlY => "Ctrl+Y", Self::CtrlZ => "Ctrl+Z",
            Self::F1 => "F1", Self::F2 => "F2", Self::F3 => "F3",
            Self::F4 => "F4", Self::F5 => "F5", Self::F6 => "F6",
            Self::F7 => "F7", Self::F8 => "F8", Self::F9 => "F9",
            Self::F10 => "F10", Self::F11 => "F11", Self::F12 => "F12",
            Self::CtrlF1 => "Ctrl+F1", Self::CtrlF2 => "Ctrl+F2", Self::CtrlF3 => "Ctrl+F3",
            Self::CtrlF4 => "Ctrl+F4", Self::CtrlF5 => "Ctrl+F5", Self::CtrlF6 => "Ctrl+F6",
            Self::CtrlF7 => "Ctrl+F7", Self::CtrlF8 => "Ctrl+F8", Self::CtrlF9 => "Ctrl+F9",
            Self::CtrlF10 => "Ctrl+F10", Self::CtrlF11 => "Ctrl+F11", Self::CtrlF12 => "Ctrl+F12",
            Self::ShiftF1 => "Shift+F1", Self::ShiftF2 => "Shift+F2", Self::ShiftF3 => "Shift+F3",
            Self::ShiftF4 => "Shift+F4", Self::ShiftF5 => "Shift+F5", Self::ShiftF6 => "Shift+F6",
            Self::ShiftF7 => "Shift+F7", Self::ShiftF8 => "Shift+F8", Self::ShiftF9 => "Shift+F9",
            Self::ShiftF10 => "Shift+F10", Self::ShiftF11 => "Shift+F11", Self::ShiftF12 => "Shift+F12",
            Self::CtrlShiftF1 => "Ctrl+Shift+F1", Self::CtrlShiftF2 => "Ctrl+Shift+F2",
            Self::CtrlShiftF3 => "Ctrl+Shift+F3", Self::CtrlShiftF4 => "Ctrl+Shift+F4",
            Self::CtrlShiftF5 => "Ctrl+Shift+F5", Self::CtrlShiftF6 => "Ctrl+Shift+F6",
            Self::CtrlShiftF7 => "Ctrl+Shift+F7", Self::CtrlShiftF8 => "Ctrl+Shift+F8",
            Self::CtrlShiftF9 => "Ctrl+Shift+F9", Self::CtrlShiftF10 => "Ctrl+Shift+F10",
            Self::CtrlShiftF11 => "Ctrl+Shift+F11", Self::CtrlShiftF12 => "Ctrl+Shift+F12",
            Self::CtrlIns => "Ctrl+Ins", Self::ShiftIns => "Shift+Ins",
            Self::Del => "Del", Self::ShiftDel => "Shift+Del",
            Self::AltBksp => "Alt+Backspace", Self::CtrlBksp => "Ctrl+Backspace",
        }
    }
}

// =============================================================================
// Control Definitions
// =============================================================================

/// Form control
pub static FORM_DEF: ControlDef = ControlDef {
    name: "Form",
    full_name: "VB.Form",
    description: "A window or dialog box that forms the main interface for an application",
    properties: &properties::FORM_PROPERTIES,
    events: &[
        EventDef { name: "Load", description: "Occurs when a form is loaded into memory", parameters: "" },
        EventDef { name: "Unload", description: "Occurs when a form is about to be removed from memory", parameters: "Cancel As Integer" },
        EventDef { name: "Initialize", description: "Occurs when an instance of a form is created", parameters: "" },
        EventDef { name: "Terminate", description: "Occurs when all references to an instance are removed", parameters: "" },
        EventDef { name: "Activate", description: "Occurs when the form becomes the active window", parameters: "" },
        EventDef { name: "Deactivate", description: "Occurs when the form is no longer the active window", parameters: "" },
        EventDef { name: "Resize", description: "Occurs when the form is resized", parameters: "" },
        EventDef { name: "Paint", description: "Occurs when the form needs repainting", parameters: "" },
        EventDef { name: "QueryUnload", description: "Occurs before the form unloads", parameters: "Cancel As Integer, UnloadMode As Integer" },
        EventDef { name: "Click", description: "Occurs when the user clicks the form", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks the form", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Show", description: "Displays the form", signature: "Show [Modal], [OwnerForm]", return_type: None },
        MethodDef { name: "Hide", description: "Hides the form without unloading it", signature: "Hide", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the form immediately", signature: "Refresh", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the form", signature: "SetFocus", return_type: None },
        MethodDef { name: "PrintForm", description: "Sends a form image to the printer", signature: "PrintForm", return_type: None },
        MethodDef { name: "Move", description: "Moves the form", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
        MethodDef { name: "ZOrder", description: "Sets the z-order of the form", signature: "ZOrder [Position]", return_type: None },
    ],
    is_container: true,
};

/// MDI Form control
pub static MDIFORM_DEF: ControlDef = ControlDef {
    name: "MDIForm",
    full_name: "VB.MDIForm",
    description: "A Multiple Document Interface form that can contain child forms",
    properties: &properties::MDIFORM_PROPERTIES,
    events: &[
        EventDef { name: "Load", description: "Occurs when the MDI form is loaded", parameters: "" },
        EventDef { name: "Unload", description: "Occurs when the MDI form is unloaded", parameters: "Cancel As Integer" },
        EventDef { name: "Activate", description: "Occurs when the MDI form becomes active", parameters: "" },
        EventDef { name: "Deactivate", description: "Occurs when the MDI form becomes inactive", parameters: "" },
        EventDef { name: "Resize", description: "Occurs when the MDI form is resized", parameters: "" },
        EventDef { name: "QueryUnload", description: "Occurs before unloading", parameters: "Cancel As Integer, UnloadMode As Integer" },
    ],
    methods: &[
        MethodDef { name: "Arrange", description: "Arranges child forms or icons", signature: "Arrange Arrangement", return_type: None },
        MethodDef { name: "Show", description: "Displays the MDI form", signature: "Show", return_type: None },
        MethodDef { name: "Hide", description: "Hides the MDI form", signature: "Hide", return_type: None },
    ],
    is_container: true,
};

/// TextBox control
pub static TEXTBOX_DEF: ControlDef = ControlDef {
    name: "TextBox",
    full_name: "VB.TextBox",
    description: "A control that displays text and allows the user to enter or edit text",
    properties: &properties::TEXTBOX_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the text content changes", parameters: "" },
        EventDef { name: "Click", description: "Occurs when the user clicks the control", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "Validate", description: "Occurs before the control loses focus", parameters: "Cancel As Boolean" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Label control
pub static LABEL_DEF: ControlDef = ControlDef {
    name: "Label",
    full_name: "VB.Label",
    description: "A control that displays text that the user cannot edit",
    properties: &properties::LABEL_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the user clicks the control", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "Change", description: "Occurs when the Caption changes", parameters: "" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// CommandButton control
pub static COMMANDBUTTON_DEF: ControlDef = ControlDef {
    name: "CommandButton",
    full_name: "VB.CommandButton",
    description: "A button control that triggers an action when clicked",
    properties: &properties::COMMANDBUTTON_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the button is clicked", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// CheckBox control
pub static CHECKBOX_DEF: ControlDef = ControlDef {
    name: "CheckBox",
    full_name: "VB.CheckBox",
    description: "A control that allows the user to select or deselect an option",
    properties: &properties::CHECKBOX_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the checkbox is clicked", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// OptionButton control
pub static OPTIONBUTTON_DEF: ControlDef = ControlDef {
    name: "OptionButton",
    full_name: "VB.OptionButton",
    description: "A radio button control for mutually exclusive options",
    properties: &properties::OPTIONBUTTON_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the option button is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Frame control
pub static FRAME_DEF: ControlDef = ControlDef {
    name: "Frame",
    full_name: "VB.Frame",
    description: "A container control that groups related controls",
    properties: &properties::FRAME_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the frame is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: true,
};

/// ListBox control
pub static LISTBOX_DEF: ControlDef = ControlDef {
    name: "ListBox",
    full_name: "VB.ListBox",
    description: "A control that displays a list of items from which the user can select",
    properties: &properties::LISTBOX_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when an item is selected", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when an item is double-clicked", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "Scroll", description: "Occurs when the list is scrolled", parameters: "" },
        EventDef { name: "ItemCheck", description: "Occurs when a checkbox item changes state", parameters: "Item As Integer" },
    ],
    methods: &[
        MethodDef { name: "AddItem", description: "Adds an item to the list", signature: "AddItem Item, [Index]", return_type: None },
        MethodDef { name: "RemoveItem", description: "Removes an item from the list", signature: "RemoveItem Index", return_type: None },
        MethodDef { name: "Clear", description: "Removes all items from the list", signature: "Clear", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// ComboBox control
pub static COMBOBOX_DEF: ControlDef = ControlDef {
    name: "ComboBox",
    full_name: "VB.ComboBox",
    description: "A combination of a text box and a drop-down list",
    properties: &properties::COMBOBOX_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the text content changes", parameters: "" },
        EventDef { name: "Click", description: "Occurs when an item is selected", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "DropDown", description: "Occurs when the list drops down", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "Scroll", description: "Occurs when the list is scrolled", parameters: "" },
    ],
    methods: &[
        MethodDef { name: "AddItem", description: "Adds an item to the list", signature: "AddItem Item, [Index]", return_type: None },
        MethodDef { name: "RemoveItem", description: "Removes an item from the list", signature: "RemoveItem Index", return_type: None },
        MethodDef { name: "Clear", description: "Removes all items from the list", signature: "Clear", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// PictureBox control
pub static PICTUREBOX_DEF: ControlDef = ControlDef {
    name: "PictureBox",
    full_name: "VB.PictureBox",
    description: "A container control that can display graphics and contain other controls",
    properties: &properties::PICTUREBOX_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the control is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "Paint", description: "Occurs when the control needs repainting", parameters: "" },
        EventDef { name: "Resize", description: "Occurs when the control is resized", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Cls", description: "Clears graphics and text", signature: "Cls", return_type: None },
        MethodDef { name: "Line", description: "Draws a line or rectangle", signature: "Line [Step] (x1, y1) - [Step] (x2, y2), [Color], [B][F]", return_type: None },
        MethodDef { name: "Circle", description: "Draws a circle or ellipse", signature: "Circle [Step] (x, y), radius, [Color], [Start], [End], [Aspect]", return_type: None },
        MethodDef { name: "PSet", description: "Sets a pixel color", signature: "PSet [Step] (x, y), [Color]", return_type: None },
        MethodDef { name: "Point", description: "Returns the color of a pixel", signature: "Point(x, y)", return_type: Some("Long") },
        MethodDef { name: "Print", description: "Prints text on the control", signature: "Print [OutputList]", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: true,
};

/// Image control
pub static IMAGE_DEF: ControlDef = ControlDef {
    name: "Image",
    full_name: "VB.Image",
    description: "A lightweight control for displaying images",
    properties: &properties::IMAGE_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the control is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Timer control
pub static TIMER_DEF: ControlDef = ControlDef {
    name: "Timer",
    full_name: "VB.Timer",
    description: "A non-visible control that triggers events at specified intervals",
    properties: &properties::TIMER_PROPERTIES,
    events: &[
        EventDef { name: "Timer", description: "Occurs when the timer interval elapses", parameters: "" },
    ],
    methods: &[],
    is_container: false,
};

/// HScrollBar control
pub static HSCROLLBAR_DEF: ControlDef = ControlDef {
    name: "HScrollBar",
    full_name: "VB.HScrollBar",
    description: "A horizontal scroll bar control",
    properties: &properties::SCROLLBAR_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the scroll position changes", parameters: "" },
        EventDef { name: "Scroll", description: "Occurs while scrolling", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// VScrollBar control
pub static VSCROLLBAR_DEF: ControlDef = ControlDef {
    name: "VScrollBar",
    full_name: "VB.VScrollBar",
    description: "A vertical scroll bar control",
    properties: &properties::SCROLLBAR_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the scroll position changes", parameters: "" },
        EventDef { name: "Scroll", description: "Occurs while scrolling", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
    ],
    methods: &[
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Shape control
pub static SHAPE_DEF: ControlDef = ControlDef {
    name: "Shape",
    full_name: "VB.Shape",
    description: "A lightweight control for drawing shapes",
    properties: &properties::SHAPE_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the shape is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Line control
pub static LINE_DEF: ControlDef = ControlDef {
    name: "Line",
    full_name: "VB.Line",
    description: "A lightweight control for drawing lines",
    properties: &properties::LINE_PROPERTIES,
    events: &[],
    methods: &[
        MethodDef { name: "Refresh", description: "Repaints the control", signature: "Refresh", return_type: None },
    ],
    is_container: false,
};

/// Data control
pub static DATA_DEF: ControlDef = ControlDef {
    name: "Data",
    full_name: "VB.Data",
    description: "A control for connecting to databases and navigating records",
    properties: &properties::DATA_PROPERTIES,
    events: &[
        EventDef { name: "Reposition", description: "Occurs after moving to a new record", parameters: "" },
        EventDef { name: "Validate", description: "Occurs before moving to a different record", parameters: "Action As Integer, Save As Integer" },
        EventDef { name: "Error", description: "Occurs when a data access error happens", parameters: "DataErr As Integer, Response As Integer" },
        EventDef { name: "MouseDown", description: "Occurs when a mouse button is pressed", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseUp", description: "Occurs when a mouse button is released", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
        EventDef { name: "MouseMove", description: "Occurs when the mouse moves", parameters: "Button As Integer, Shift As Integer, X As Single, Y As Single" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Refreshes the recordset", signature: "Refresh", return_type: None },
        MethodDef { name: "UpdateControls", description: "Updates bound controls", signature: "UpdateControls", return_type: None },
        MethodDef { name: "UpdateRecord", description: "Saves changes to current record", signature: "UpdateRecord", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// OLE control
pub static OLE_DEF: ControlDef = ControlDef {
    name: "OLE",
    full_name: "VB.OLE",
    description: "A control for embedding and linking OLE objects",
    properties: &properties::OLE_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the control is clicked", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when the user double-clicks", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "Updated", description: "Occurs when the linked object is updated", parameters: "Code As Integer" },
        EventDef { name: "Resize", description: "Occurs when the control is resized", parameters: "HeightNew As Single, WidthNew As Single" },
    ],
    methods: &[
        MethodDef { name: "CreateEmbed", description: "Creates an embedded object", signature: "CreateEmbed SourceDoc, [Class]", return_type: None },
        MethodDef { name: "CreateLink", description: "Creates a linked object", signature: "CreateLink SourceDoc, [SourceItem]", return_type: None },
        MethodDef { name: "Delete", description: "Deletes the OLE object", signature: "Delete", return_type: None },
        MethodDef { name: "DoVerb", description: "Opens the object for editing", signature: "DoVerb [Verb]", return_type: None },
        MethodDef { name: "InsertObjDlg", description: "Shows Insert Object dialog", signature: "InsertObjDlg", return_type: None },
        MethodDef { name: "PasteSpecialDlg", description: "Shows Paste Special dialog", signature: "PasteSpecialDlg", return_type: None },
        MethodDef { name: "SaveToFile", description: "Saves object to file", signature: "SaveToFile FileNum", return_type: None },
        MethodDef { name: "ReadFromFile", description: "Reads object from file", signature: "ReadFromFile FileNum", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// FileListBox control
pub static FILELISTBOX_DEF: ControlDef = ControlDef {
    name: "FileListBox",
    full_name: "VB.FileListBox",
    description: "A control that displays files in the current directory",
    properties: &properties::FILELISTBOX_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when a file is selected", parameters: "" },
        EventDef { name: "DblClick", description: "Occurs when a file is double-clicked", parameters: "" },
        EventDef { name: "PathChange", description: "Occurs when the path changes", parameters: "" },
        EventDef { name: "PatternChange", description: "Occurs when the pattern changes", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "Scroll", description: "Occurs when the list is scrolled", parameters: "" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Refreshes the file list", signature: "Refresh", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// DirListBox control
pub static DIRLISTBOX_DEF: ControlDef = ControlDef {
    name: "DirListBox",
    full_name: "VB.DirListBox",
    description: "A control that displays the directory structure",
    properties: &properties::DIRLISTBOX_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the directory changes", parameters: "" },
        EventDef { name: "Click", description: "Occurs when a directory is selected", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "Scroll", description: "Occurs when the list is scrolled", parameters: "" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Refreshes the directory list", signature: "Refresh", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// DriveListBox control
pub static DRIVELISTBOX_DEF: ControlDef = ControlDef {
    name: "DriveListBox",
    full_name: "VB.DriveListBox",
    description: "A control that displays available drives",
    properties: &properties::DRIVELISTBOX_PROPERTIES,
    events: &[
        EventDef { name: "Change", description: "Occurs when the selected drive changes", parameters: "" },
        EventDef { name: "GotFocus", description: "Occurs when the control receives focus", parameters: "" },
        EventDef { name: "LostFocus", description: "Occurs when the control loses focus", parameters: "" },
        EventDef { name: "KeyDown", description: "Occurs when a key is pressed", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyUp", description: "Occurs when a key is released", parameters: "KeyCode As Integer, Shift As Integer" },
        EventDef { name: "KeyPress", description: "Occurs when a key is pressed and released", parameters: "KeyAscii As Integer" },
        EventDef { name: "Scroll", description: "Occurs when the list is scrolled", parameters: "" },
    ],
    methods: &[
        MethodDef { name: "Refresh", description: "Refreshes the drive list", signature: "Refresh", return_type: None },
        MethodDef { name: "SetFocus", description: "Gives focus to the control", signature: "SetFocus", return_type: None },
        MethodDef { name: "Move", description: "Moves the control", signature: "Move Left, [Top], [Width], [Height]", return_type: None },
    ],
    is_container: false,
};

/// Menu control
pub static MENU_DEF: ControlDef = ControlDef {
    name: "Menu",
    full_name: "VB.Menu",
    description: "A menu item that appears on a form's menu bar",
    properties: &properties::MENU_PROPERTIES,
    events: &[
        EventDef { name: "Click", description: "Occurs when the menu item is clicked", parameters: "" },
    ],
    methods: &[],
    is_container: false,
};

// =============================================================================
// Control Registry
// =============================================================================

/// All control definitions indexed by type name
static CONTROL_REGISTRY: Lazy<HashMap<&'static str, &'static ControlDef>> = Lazy::new(|| {
    let mut map = HashMap::new();

    // Standard controls
    map.insert("Form", &FORM_DEF);
    map.insert("MDIForm", &MDIFORM_DEF);
    map.insert("TextBox", &TEXTBOX_DEF);
    map.insert("Label", &LABEL_DEF);
    map.insert("CommandButton", &COMMANDBUTTON_DEF);
    map.insert("CheckBox", &CHECKBOX_DEF);
    map.insert("OptionButton", &OPTIONBUTTON_DEF);
    map.insert("Frame", &FRAME_DEF);
    map.insert("ListBox", &LISTBOX_DEF);
    map.insert("ComboBox", &COMBOBOX_DEF);
    map.insert("PictureBox", &PICTUREBOX_DEF);
    map.insert("Image", &IMAGE_DEF);
    map.insert("Timer", &TIMER_DEF);
    map.insert("HScrollBar", &HSCROLLBAR_DEF);
    map.insert("VScrollBar", &VSCROLLBAR_DEF);
    map.insert("Shape", &SHAPE_DEF);
    map.insert("Line", &LINE_DEF);
    map.insert("Data", &DATA_DEF);
    map.insert("OLE", &OLE_DEF);
    map.insert("FileListBox", &FILELISTBOX_DEF);
    map.insert("DirListBox", &DIRLISTBOX_DEF);
    map.insert("DriveListBox", &DRIVELISTBOX_DEF);
    map.insert("Menu", &MENU_DEF);

    map
});

/// Get a control definition by type name (case-insensitive)
pub fn get_control(type_name: &str) -> Option<&'static ControlDef> {
    // Try exact match first
    if let Some(def) = CONTROL_REGISTRY.get(type_name) {
        return Some(def);
    }

    // Case-insensitive search
    let type_lower = type_name.to_lowercase();
    for (name, def) in CONTROL_REGISTRY.iter() {
        if name.to_lowercase() == type_lower {
            return Some(def);
        }
    }

    None
}

/// Get all available control names
pub fn get_control_names() -> Vec<&'static str> {
    CONTROL_REGISTRY.keys().copied().collect()
}

/// Get property definition for a control
pub fn get_property(control_type: &str, property_name: &str) -> Option<&'static PropertyDef> {
    let control = get_control(control_type)?;
    let prop_lower = property_name.to_lowercase();

    control.properties.iter()
        .find(|p| p.name.to_lowercase() == prop_lower)
}

/// Get event definition for a control
pub fn get_event(control_type: &str, event_name: &str) -> Option<&'static EventDef> {
    let control = get_control(control_type)?;
    let event_lower = event_name.to_lowercase();

    control.events.iter()
        .find(|e| e.name.to_lowercase() == event_lower)
}

/// Get method definition for a control
pub fn get_method(control_type: &str, method_name: &str) -> Option<&'static MethodDef> {
    let control = get_control(control_type)?;
    let method_lower = method_name.to_lowercase();

    control.methods.iter()
        .find(|m| m.name.to_lowercase() == method_lower)
}

/// Get all property names for a control type
pub fn get_property_names(control_type: &str) -> Vec<&'static str> {
    get_control(control_type)
        .map(|c| c.properties.iter().map(|p| p.name).collect())
        .unwrap_or_default()
}

/// Get all event names for a control type
pub fn get_event_names(control_type: &str) -> Vec<&'static str> {
    get_control(control_type)
        .map(|c| c.events.iter().map(|e| e.name).collect())
        .unwrap_or_default()
}

/// Get all method names for a control type
pub fn get_method_names(control_type: &str) -> Vec<&'static str> {
    get_control(control_type)
        .map(|c| c.methods.iter().map(|m| m.name).collect())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_control() {
        let textbox = get_control("TextBox");
        assert!(textbox.is_some());
        assert_eq!(textbox.unwrap().name, "TextBox");

        // Case-insensitive
        let textbox2 = get_control("textbox");
        assert!(textbox2.is_some());
    }

    #[test]
    fn test_get_property() {
        let prop = get_property("TextBox", "Text");
        assert!(prop.is_some());
        assert_eq!(prop.unwrap().name, "Text");
    }

    #[test]
    fn test_get_event() {
        let event = get_event("CommandButton", "Click");
        assert!(event.is_some());
        assert_eq!(event.unwrap().name, "Click");
    }

    #[test]
    fn test_menu_shortcut() {
        assert_eq!(MenuShortcut::CtrlS.display(), "Ctrl+S");
        assert_eq!(MenuShortcut::from_str("^S"), Some(MenuShortcut::CtrlS));
    }
}
