/**
 * Symbol Definitions
 *
 * Defines the Symbol interface and related types for the symbol table.
 */

import type { SourceRange } from "./position.js";

/**
 * Unique identifier for a symbol within the symbol table
 */
export type SymbolId = number;

/**
 * Unique identifier for a scope
 */
export type ScopeId = number;

/**
 * The kind of symbol
 */
export enum SymbolKind {
  // Module-level declarations
  Variable = "Variable",
  Constant = "Constant",
  UserDefinedType = "UserDefinedType",
  Enum = "Enum",
  EnumMember = "EnumMember",
  TypeMember = "TypeMember",

  // Procedures
  Sub = "Sub",
  Function = "Function",
  PropertyGet = "PropertyGet",
  PropertyLet = "PropertyLet",
  PropertySet = "PropertySet",
  Event = "Event",
  DeclareFunction = "DeclareFunction",
  DeclareSub = "DeclareSub",

  // Local scope
  Parameter = "Parameter",
  LocalVariable = "LocalVariable",
  LocalConstant = "LocalConstant",

  // Special
  ForLoopVariable = "ForLoopVariable",
  ForEachVariable = "ForEachVariable",
  Label = "Label",

  // Form controls (TextBox, Label, Button, etc.)
  FormControl = "FormControl",
}

/**
 * Check if this symbol kind creates a scope
 */
export function kindCreatesScope(kind: SymbolKind): boolean {
  return [
    SymbolKind.Sub,
    SymbolKind.Function,
    SymbolKind.PropertyGet,
    SymbolKind.PropertyLet,
    SymbolKind.PropertySet,
  ].includes(kind);
}

/**
 * Check if this is a procedure-like symbol
 */
export function kindIsProcedure(kind: SymbolKind): boolean {
  return [
    SymbolKind.Sub,
    SymbolKind.Function,
    SymbolKind.PropertyGet,
    SymbolKind.PropertyLet,
    SymbolKind.PropertySet,
    SymbolKind.DeclareFunction,
    SymbolKind.DeclareSub,
  ].includes(kind);
}

/**
 * Check if this is a callable symbol
 */
export function kindIsCallable(kind: SymbolKind): boolean {
  return kindIsProcedure(kind) || kind === SymbolKind.Event;
}

/**
 * Get a display string for the kind
 */
export function kindDisplayName(kind: SymbolKind): string {
  const displayNames: Record<SymbolKind, string> = {
    [SymbolKind.Variable]: "Variable",
    [SymbolKind.LocalVariable]: "Variable",
    [SymbolKind.Constant]: "Constant",
    [SymbolKind.LocalConstant]: "Constant",
    [SymbolKind.UserDefinedType]: "Type",
    [SymbolKind.Enum]: "Enum",
    [SymbolKind.EnumMember]: "Enum Member",
    [SymbolKind.TypeMember]: "Field",
    [SymbolKind.Sub]: "Sub",
    [SymbolKind.DeclareSub]: "Sub",
    [SymbolKind.Function]: "Function",
    [SymbolKind.DeclareFunction]: "Function",
    [SymbolKind.PropertyGet]: "Property Get",
    [SymbolKind.PropertyLet]: "Property Let",
    [SymbolKind.PropertySet]: "Property Set",
    [SymbolKind.Event]: "Event",
    [SymbolKind.Parameter]: "Parameter",
    [SymbolKind.ForLoopVariable]: "Loop Variable",
    [SymbolKind.ForEachVariable]: "Loop Variable",
    [SymbolKind.Label]: "Label",
    [SymbolKind.FormControl]: "Control",
  };
  return displayNames[kind] || kind;
}

/**
 * Visibility level for symbols
 */
export enum Visibility {
  Private = "Private",
  Public = "Public",
  Friend = "Friend",
  Global = "Global",
}

/**
 * Type information for a symbol
 */
export interface TypeInfo {
  /** The type name (e.g., "Integer", "String", "MyClass") */
  name: string;
  /** Whether this is an array type */
  isArray: boolean;
  /** Whether this is a New expression type (for classes) */
  isNew: boolean;
}

export function createTypeInfo(
  name: string,
  isArray = false,
  isNew = false
): TypeInfo {
  return { name, isArray, isNew };
}

/**
 * Format type for display (e.g., "Integer()" for arrays)
 */
export function formatTypeInfo(type: TypeInfo): string {
  return type.isArray ? `${type.name}()` : type.name;
}

/**
 * Parameter information for procedures
 */
export interface ParameterInfo {
  /** Parameter name */
  name: string;
  /** Parameter type */
  typeInfo?: TypeInfo;
  /** Whether passed by reference (default in VB6) */
  byRef: boolean;
  /** Whether optional */
  optional: boolean;
  /** Default value expression (for optional params) */
  defaultValue?: string;
  /** Position range of the entire parameter declaration */
  range: SourceRange;
  /** Position range of just the name */
  nameRange: SourceRange;
}

/**
 * Format parameter for signature display
 */
export function formatParameter(param: ParameterInfo): string {
  const parts: string[] = [];

  if (param.optional) {
    parts.push("Optional");
  }

  if (param.byRef) {
    parts.push("ByRef");
  } else {
    parts.push("ByVal");
  }

  parts.push(param.name);

  if (param.typeInfo) {
    parts.push(`As ${formatTypeInfo(param.typeInfo)}`);
  }

  if (param.defaultValue) {
    parts.push(`= ${param.defaultValue}`);
  }

  return parts.join(" ");
}

/**
 * A symbol definition in the symbol table
 */
export interface Symbol {
  /** Unique identifier */
  id: SymbolId;
  /** Symbol name (case-preserved, but lookups are case-insensitive) */
  name: string;
  /** Symbol kind */
  kind: SymbolKind;
  /** Visibility */
  visibility: Visibility;
  /** Type information (if applicable) */
  typeInfo?: TypeInfo;
  /** The range of the entire declaration */
  definitionRange: SourceRange;
  /** The range of just the name (for precise go-to-definition) */
  nameRange: SourceRange;
  /** The scope this symbol belongs to */
  scopeId: ScopeId;
  /** For procedures: parameters */
  parameters: ParameterInfo[];
  /** For enums/types: member symbol IDs */
  members: SymbolId[];
  /** Documentation/comments associated with this symbol */
  documentation?: string;
  /** Value (for constants and enum members) */
  value?: string;
}

/**
 * Format the symbol as a signature for hover display
 */
export function formatSignature(symbol: Symbol): string {
  const vis = symbol.visibility;
  const typeStr = symbol.typeInfo
    ? formatTypeInfo(symbol.typeInfo)
    : "Variant";
  const params = symbol.parameters.map(formatParameter).join(", ");

  switch (symbol.kind) {
    case SymbolKind.Sub:
    case SymbolKind.DeclareSub:
      return `${vis} Sub ${symbol.name}(${params})`;

    case SymbolKind.Function:
    case SymbolKind.DeclareFunction:
      return `${vis} Function ${symbol.name}(${params}) As ${typeStr}`;

    case SymbolKind.PropertyGet:
      return `${vis} Property Get ${symbol.name}(${params}) As ${typeStr}`;

    case SymbolKind.PropertyLet:
      return `${vis} Property Let ${symbol.name}(${params})`;

    case SymbolKind.PropertySet:
      return `${vis} Property Set ${symbol.name}(${params})`;

    case SymbolKind.Variable:
    case SymbolKind.LocalVariable:
      return `${vis} ${symbol.name} As ${typeStr}`;

    case SymbolKind.Constant:
    case SymbolKind.LocalConstant:
      return `${vis} Const ${symbol.name} = ${symbol.value ?? "?"}`;

    case SymbolKind.Parameter:
      return `Parameter ${symbol.name} As ${typeStr}`;

    case SymbolKind.Event:
      return `${vis} Event ${symbol.name}(${params})`;

    case SymbolKind.UserDefinedType:
      return `${vis} Type ${symbol.name}`;

    case SymbolKind.Enum:
      return `${vis} Enum ${symbol.name}`;

    case SymbolKind.EnumMember:
      return symbol.value ? `${symbol.name} = ${symbol.value}` : symbol.name;

    case SymbolKind.TypeMember:
      return `${symbol.name} As ${typeStr}`;

    case SymbolKind.ForLoopVariable:
    case SymbolKind.ForEachVariable:
      return `(loop variable) ${symbol.name}`;

    case SymbolKind.Label:
      return `${symbol.name}:`;

    case SymbolKind.FormControl:
      return `${symbol.name} As ${typeStr}`;

    default:
      return symbol.name;
  }
}

/**
 * A reference to a symbol (usage site)
 */
export interface SymbolReference {
  /** The symbol being referenced */
  symbolId: SymbolId;
  /** The range of this reference */
  range: SourceRange;
  /** The scope where this reference occurs */
  scopeId: ScopeId;
  /** Whether this is an assignment target (LHS of =) */
  isAssignment: boolean;
}
