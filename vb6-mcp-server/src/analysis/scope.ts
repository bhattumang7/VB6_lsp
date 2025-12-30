/**
 * Scope Definitions
 *
 * Defines scope hierarchy for the symbol table.
 */

import type { SourceRange } from "./position.js";
import type { ScopeId, SymbolId } from "./types.js";

/**
 * The type of scope
 */
export enum ScopeKind {
  /** Module-level scope (file scope) */
  Module = "Module",
  /** Procedure scope (Sub, Function, Property) */
  Procedure = "Procedure",
  /** With block scope (implicit object reference) */
  WithBlock = "WithBlock",
  /** For loop scope (loop variable) */
  ForLoop = "ForLoop",
  /** For Each loop scope */
  ForEachLoop = "ForEachLoop",
}

/**
 * Check if this scope kind introduces a new variable scope
 * (where local variables can be declared)
 */
export function scopeIsVariableScope(kind: ScopeKind): boolean {
  return kind === ScopeKind.Module || kind === ScopeKind.Procedure;
}

/**
 * Get a display name for the scope kind
 */
export function scopeDisplayName(kind: ScopeKind): string {
  const displayNames: Record<ScopeKind, string> = {
    [ScopeKind.Module]: "Module",
    [ScopeKind.Procedure]: "Procedure",
    [ScopeKind.WithBlock]: "With Block",
    [ScopeKind.ForLoop]: "For Loop",
    [ScopeKind.ForEachLoop]: "For Each Loop",
  };
  return displayNames[kind];
}

/**
 * A scope in the scope hierarchy
 */
export interface Scope {
  /** Unique identifier */
  id: ScopeId;
  /** Scope kind */
  kind: ScopeKind;
  /** Parent scope (undefined for module scope) */
  parent?: ScopeId;
  /** Range covered by this scope */
  range: SourceRange;
  /** Symbols declared in this scope (lowercase name -> symbol id) */
  symbols: Map<string, SymbolId>;
  /** Child scopes (in document order) */
  children: ScopeId[];
  /** For WithBlock: the object expression being referenced */
  withObject?: string;
  /** The symbol that created this scope (for procedure scopes) */
  definingSymbol?: SymbolId;
}

/**
 * Create a new scope
 */
export function createScope(
  id: ScopeId,
  kind: ScopeKind,
  parent: ScopeId | undefined,
  range: SourceRange
): Scope {
  return {
    id,
    kind,
    parent,
    range,
    symbols: new Map(),
    children: [],
    withObject: undefined,
    definingSymbol: undefined,
  };
}

/**
 * Add a symbol to a scope
 */
export function addSymbolToScope(
  scope: Scope,
  name: string,
  symbolId: SymbolId
): void {
  // Use lowercase for case-insensitive lookup
  scope.symbols.set(name.toLowerCase(), symbolId);
}

/**
 * Look up a symbol by name in this scope only (case-insensitive)
 */
export function lookupLocalSymbol(
  scope: Scope,
  name: string
): SymbolId | undefined {
  return scope.symbols.get(name.toLowerCase());
}

/**
 * Get all symbol IDs in this scope
 */
export function getScopeSymbols(scope: Scope): SymbolId[] {
  return Array.from(scope.symbols.values());
}

/**
 * Get the number of symbols in this scope
 */
export function getScopeSymbolCount(scope: Scope): number {
  return scope.symbols.size;
}

/**
 * Check if a symbol name exists in this scope
 */
export function scopeHasSymbol(scope: Scope, name: string): boolean {
  return scope.symbols.has(name.toLowerCase());
}

/**
 * Add a child scope
 */
export function addChildScope(scope: Scope, childId: ScopeId): void {
  scope.children.push(childId);
}
