/**
 * Symbol Table
 *
 * The main symbol table that stores all symbols and scopes for a document.
 */

import {
  type SourcePosition,
  type SourceRange,
  moduleRange,
  rangeContains,
  rangeSize,
} from "./position.js";
import {
  type Scope,
  ScopeKind,
  createScope,
  addSymbolToScope,
  lookupLocalSymbol,
  getScopeSymbols,
  addChildScope,
} from "./scope.js";
import {
  type Symbol,
  type SymbolId,
  type ScopeId,
  type SymbolKind,
  type Visibility,
  type TypeInfo,
  type ParameterInfo,
  type SymbolReference,
} from "./types.js";

/**
 * The complete symbol table for a document
 */
export class SymbolTable {
  /** Document file path */
  public filePath: string;

  /** All symbols, indexed by ID */
  private symbols: Symbol[] = [];

  /** All scopes, indexed by ID */
  private scopes: Scope[] = [];

  /** The module-level (root) scope */
  public moduleScope: ScopeId = 0;

  /** All references to symbols */
  private references: SymbolReference[] = [];

  /** Spatial index: map from line number to symbols defined on that line */
  private symbolsByLine: Map<number, SymbolId[]> = new Map();

  /** Spatial index: map from line number to scopes that contain that line */
  private scopesByLine: Map<number, ScopeId[]> = new Map();

  /** Next symbol ID to allocate */
  private nextSymbolId: number = 0;

  /** Next scope ID to allocate */
  private nextScopeId: number = 0;

  constructor(filePath: string) {
    this.filePath = filePath;

    // Create the module scope (covers entire file)
    this.moduleScope = this.createScopeInternal(
      ScopeKind.Module,
      undefined,
      moduleRange()
    );
  }

  // ==========================================
  // Symbol Management
  // ==========================================

  /**
   * Create a new symbol and add it to the table
   */
  createSymbol(
    name: string,
    kind: SymbolKind,
    visibility: Visibility,
    definitionRange: SourceRange,
    nameRange: SourceRange,
    scopeId: ScopeId
  ): SymbolId {
    const id = this.nextSymbolId++;

    const symbol: Symbol = {
      id,
      name,
      kind,
      visibility,
      definitionRange,
      nameRange,
      scopeId,
      parameters: [],
      members: [],
    };

    // Add to spatial index (index by nameRange lines for precise lookup)
    for (let line = nameRange.start.line; line <= nameRange.end.line; line++) {
      const existing = this.symbolsByLine.get(line) || [];
      existing.push(id);
      this.symbolsByLine.set(line, existing);
    }

    // Add to scope
    const scope = this.scopes[scopeId];
    if (scope) {
      addSymbolToScope(scope, name, id);
    }

    this.symbols.push(symbol);
    return id;
  }

  /**
   * Get a symbol by ID
   */
  getSymbol(id: SymbolId): Symbol | undefined {
    return this.symbols[id];
  }

  /**
   * Set type info for a symbol
   */
  setTypeInfo(id: SymbolId, typeInfo: TypeInfo): void {
    const symbol = this.symbols[id];
    if (symbol) {
      symbol.typeInfo = typeInfo;
    }
  }

  /**
   * Set value for a symbol (constants, enum members)
   */
  setValue(id: SymbolId, value: string): void {
    const symbol = this.symbols[id];
    if (symbol) {
      symbol.value = value;
    }
  }

  /**
   * Add parameters to a procedure symbol
   */
  setParameters(id: SymbolId, parameters: ParameterInfo[]): void {
    const symbol = this.symbols[id];
    if (symbol) {
      symbol.parameters = parameters;
    }
  }

  /**
   * Add a member to a type/enum symbol
   */
  addMember(parentId: SymbolId, memberId: SymbolId): void {
    const symbol = this.symbols[parentId];
    if (symbol) {
      symbol.members.push(memberId);
    }
  }

  /**
   * Set documentation for a symbol
   */
  setDocumentation(id: SymbolId, doc: string): void {
    const symbol = this.symbols[id];
    if (symbol) {
      symbol.documentation = doc;
    }
  }

  // ==========================================
  // Scope Management
  // ==========================================

  /**
   * Create a new scope
   */
  createScope(
    kind: ScopeKind,
    parent: ScopeId | undefined,
    range: SourceRange
  ): ScopeId {
    return this.createScopeInternal(kind, parent, range);
  }

  private createScopeInternal(
    kind: ScopeKind,
    parent: ScopeId | undefined,
    range: SourceRange
  ): ScopeId {
    const id = this.nextScopeId++;

    const scope = createScope(id, kind, parent, range);

    // Add to spatial index (limit to reasonable range to avoid memory issues)
    const endLine = Math.min(range.end.line, range.start.line + 10000);
    for (let line = range.start.line; line <= endLine; line++) {
      const existing = this.scopesByLine.get(line) || [];
      existing.push(id);
      this.scopesByLine.set(line, existing);
    }

    // Add as child to parent
    if (parent !== undefined) {
      const parentScope = this.scopes[parent];
      if (parentScope) {
        addChildScope(parentScope, id);
      }
    }

    this.scopes.push(scope);
    return id;
  }

  /**
   * Get a scope by ID
   */
  getScope(id: ScopeId): Scope | undefined {
    return this.scopes[id];
  }

  /**
   * Link a procedure symbol to its scope
   */
  linkProcedureScope(symbolId: SymbolId, scopeId: ScopeId): void {
    const scope = this.scopes[scopeId];
    if (scope) {
      scope.definingSymbol = symbolId;
    }
  }

  /**
   * Set with object for a with block scope
   */
  setWithObject(scopeId: ScopeId, object: string): void {
    const scope = this.scopes[scopeId];
    if (scope) {
      scope.withObject = object;
    }
  }

  // ==========================================
  // Reference Tracking
  // ==========================================

  /**
   * Add a reference to a symbol
   */
  addReference(
    symbolId: SymbolId,
    range: SourceRange,
    scopeId: ScopeId,
    isAssignment: boolean
  ): void {
    this.references.push({
      symbolId,
      range,
      scopeId,
      isAssignment,
    });
  }

  /**
   * Get all references to a symbol
   */
  getReferences(symbolId: SymbolId): SymbolReference[] {
    return this.references.filter((r) => r.symbolId === symbolId);
  }

  // ==========================================
  // Query Methods
  // ==========================================

  /**
   * Find the innermost scope containing a position
   */
  scopeAtPosition(pos: SourcePosition): ScopeId {
    const scopeIds = this.scopesByLine.get(pos.line);
    if (!scopeIds) {
      return this.moduleScope;
    }

    // Find the innermost (smallest) scope that contains the position
    let best = this.moduleScope;
    let bestSize = Number.MAX_SAFE_INTEGER;

    for (const scopeId of scopeIds) {
      const scope = this.scopes[scopeId];
      if (scope && rangeContains(scope.range, pos)) {
        const size = rangeSize(scope.range);
        if (size < bestSize) {
          best = scopeId;
          bestSize = size;
        }
      }
    }

    return best;
  }

  /**
   * Look up a symbol by name, searching from the given scope up to the root
   */
  lookupSymbol(name: string, fromScope: ScopeId): Symbol | undefined {
    let current: ScopeId | undefined = fromScope;

    while (current !== undefined) {
      const scope: Scope | undefined = this.scopes[current];
      if (!scope) break;

      const symbolId = lookupLocalSymbol(scope, name);
      if (symbolId !== undefined) {
        return this.symbols[symbolId];
      }

      current = scope.parent;
    }

    return undefined;
  }

  /**
   * Look up a symbol at a specific position
   */
  lookupAtPosition(name: string, pos: SourcePosition): Symbol | undefined {
    const scopeId = this.scopeAtPosition(pos);
    return this.lookupSymbol(name, scopeId);
  }

  /**
   * Find the symbol whose nameRange contains the given position
   */
  symbolAtPosition(pos: SourcePosition): Symbol | undefined {
    // First check if we're on a symbol definition
    const symbolIds = this.symbolsByLine.get(pos.line);
    if (symbolIds) {
      for (const symbolId of symbolIds) {
        const symbol = this.symbols[symbolId];
        if (symbol && rangeContains(symbol.nameRange, pos)) {
          return symbol;
        }
      }
    }

    // Then check references
    for (const reference of this.references) {
      if (rangeContains(reference.range, pos)) {
        return this.symbols[reference.symbolId];
      }
    }

    return undefined;
  }

  /**
   * Find the reference at a specific position (if any)
   */
  referenceAtPosition(pos: SourcePosition): SymbolReference | undefined {
    return this.references.find((r) => rangeContains(r.range, pos));
  }

  /**
   * Get all symbols in a scope (not recursive)
   */
  symbolsInScope(scopeId: ScopeId): Symbol[] {
    const scope = this.scopes[scopeId];
    if (!scope) return [];

    return getScopeSymbols(scope)
      .map((id) => this.symbols[id])
      .filter((s): s is Symbol => s !== undefined);
  }

  /**
   * Get all module-level symbols
   */
  moduleSymbols(): Symbol[] {
    return this.symbolsInScope(this.moduleScope);
  }

  /**
   * Get all symbols (for document outline)
   */
  allSymbols(): Symbol[] {
    return [...this.symbols];
  }

  /**
   * Get all symbols of a specific kind
   */
  symbolsOfKind(kind: SymbolKind): Symbol[] {
    return this.symbols.filter((s) => s.kind === kind);
  }

  /**
   * Find definition of a symbol by name and position
   */
  findDefinition(name: string, pos: SourcePosition): Symbol | undefined {
    return this.lookupAtPosition(name, pos);
  }

  /**
   * Find all references to a symbol at position (including the definition)
   */
  findAllReferences(pos: SourcePosition): SourceRange[] {
    const symbol = this.symbolAtPosition(pos);
    if (!symbol) return [];

    const ranges = [symbol.nameRange];

    for (const reference of this.getReferences(symbol.id)) {
      ranges.push(reference.range);
    }

    return ranges;
  }

  /**
   * Get visible symbols at a position (for completion)
   */
  visibleSymbols(pos: SourcePosition): Symbol[] {
    const scopeId = this.scopeAtPosition(pos);
    const visible: Symbol[] = [];
    const seenNames = new Set<string>();

    let current: ScopeId | undefined = scopeId;

    // Walk up the scope chain
    while (current !== undefined) {
      const scope: Scope | undefined = this.scopes[current];
      if (!scope) break;

      for (const symbolId of getScopeSymbols(scope)) {
        const symbol = this.symbols[symbolId];
        if (symbol) {
          const lowerName = symbol.name.toLowerCase();
          if (!seenNames.has(lowerName)) {
            visible.push(symbol);
            seenNames.add(lowerName);
          }
        }
      }

      current = scope.parent;
    }

    return visible;
  }

  /**
   * Get all scopes
   */
  allScopes(): Scope[] {
    return [...this.scopes];
  }

  /**
   * Get the count of symbols
   */
  symbolCount(): number {
    return this.symbols.length;
  }

  /**
   * Get the count of scopes
   */
  scopeCount(): number {
    return this.scopes.length;
  }

  /**
   * Get the count of references
   */
  referenceCount(): number {
    return this.references.length;
  }
}
