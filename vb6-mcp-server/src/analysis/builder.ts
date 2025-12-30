/**
 * Symbol Table Builder
 *
 * Builds a symbol table by walking the tree-sitter parse tree.
 */

import type { Tree, SyntaxNode } from "tree-sitter";
import { fromTsNode, type SourceRange } from "./position.js";
import { ScopeKind } from "./scope.js";
import {
  SymbolKind,
  Visibility,
  type TypeInfo,
  type ParameterInfo,
  type ScopeId,
} from "./types.js";
import { SymbolTable } from "./symbolTable.js";

/**
 * Builds a symbol table from a tree-sitter parse tree
 */
export class SymbolTableBuilder {
  private source: string;
  private table: SymbolTable;
  /** Stack of current scopes (innermost last) */
  private scopeStack: ScopeId[];

  constructor(filePath: string, source: string) {
    this.source = source;
    this.table = new SymbolTable(filePath);
    this.scopeStack = [this.table.moduleScope];
  }

  /**
   * Build the symbol table from a parse tree
   */
  build(tree: Tree): SymbolTable {
    // First pass: collect all symbol definitions
    this.visitNode(tree.rootNode);

    // Second pass: collect all references to symbols
    this.scopeStack = [this.table.moduleScope];
    this.collectReferences(tree.rootNode);

    return this.table;
  }

  /** Get the current scope */
  private currentScope(): ScopeId {
    return this.scopeStack[this.scopeStack.length - 1];
  }

  /** Push a new scope onto the stack */
  private pushScope(kind: ScopeKind, range: SourceRange): ScopeId {
    const parent = this.currentScope();
    const scopeId = this.table.createScope(kind, parent, range);
    this.scopeStack.push(scopeId);
    return scopeId;
  }

  /** Pop the current scope */
  private popScope(): void {
    if (this.scopeStack.length > 1) {
      this.scopeStack.pop();
    }
  }

  /** Get text content of a node */
  private nodeText(node: SyntaxNode): string {
    return node.text;
  }

  /** Get range from a node */
  private nodeRange(node: SyntaxNode): SourceRange {
    return fromTsNode(node);
  }

  /** Extract visibility from a declaration node */
  private extractVisibility(node: SyntaxNode): Visibility {
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (!child) continue;
      const text = this.nodeText(child).toUpperCase();
      switch (text) {
        case "PUBLIC":
          return Visibility.Public;
        case "GLOBAL":
          return Visibility.Global;
        case "PRIVATE":
          return Visibility.Private;
        case "FRIEND":
          return Visibility.Friend;
      }
    }
    return Visibility.Private; // Default
  }

  /** Find a child node by field name */
  private findField(node: SyntaxNode, fieldName: string): SyntaxNode | null {
    return node.childForFieldName(fieldName);
  }

  /** Find all children of a specific kind */
  private findChildrenByKind(node: SyntaxNode, kind: string): SyntaxNode[] {
    const result: SyntaxNode[] = [];
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (child && child.type === kind) {
        result.push(child);
      }
    }
    return result;
  }

  /** Check if node has a specific keyword child */
  private hasChildKeyword(node: SyntaxNode, keyword: string): boolean {
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (child && this.nodeText(child).toUpperCase() === keyword.toUpperCase()) {
        return true;
      }
    }
    return false;
  }

  /** Extract type from as_clause node */
  private extractTypeFromAsClause(node: SyntaxNode): TypeInfo | undefined {
    const typeNode = this.findField(node, "type");
    if (typeNode) {
      const name = this.nodeText(typeNode);
      const isArray =
        name.endsWith("()") ||
        this.findChildrenByKind(node, "array_bounds").length > 0;
      const isNew = this.hasChildKeyword(node, "new");

      return {
        name: name.replace(/\(\)$/, ""),
        isArray,
        isNew,
      };
    }
    return undefined;
  }

  /** Extract type from a declaration node (looks for as_clause child) */
  private extractType(node: SyntaxNode): TypeInfo | undefined {
    for (const child of this.findChildrenByKind(node, "as_clause")) {
      const typeInfo = this.extractTypeFromAsClause(child);
      if (typeInfo) return typeInfo;
    }
    return undefined;
  }

  /** Check if currently in module scope */
  private isModuleScope(): boolean {
    const scope = this.table.getScope(this.currentScope());
    return scope?.kind === ScopeKind.Module;
  }

  /** Visit a node and its children */
  private visitNode(node: SyntaxNode): void {
    switch (node.type) {
      // Skip form designer property lines
      case "form_property_line":
      case "form_property_block":
      case "module_config":
      case "module_config_element":
        return;

      // Form blocks - create symbol for the control name
      case "form_block":
        this.visitFormBlock(node);
        return;

      // Declarations that create symbols
      case "variable_declaration":
        this.visitVariableDeclaration(node);
        break;
      case "constant_declaration":
        this.visitConstantDeclaration(node);
        break;
      case "type_declaration":
        this.visitTypeDeclaration(node);
        break;
      case "enum_declaration":
        this.visitEnumDeclaration(node);
        break;
      case "sub_declaration":
        this.visitSubDeclaration(node);
        return;
      case "function_declaration":
        this.visitFunctionDeclaration(node);
        return;
      case "property_declaration":
        this.visitPropertyDeclaration(node);
        return;
      case "declare_statement":
        this.visitDeclareStatement(node);
        break;
      case "event_statement":
        this.visitEventStatement(node);
        break;

      // Scope-creating constructs
      case "with_statement":
        this.visitWithStatement(node);
        return;
      case "for_statement":
        this.visitForStatement(node);
        return;
      case "for_each_statement":
        this.visitForEachStatement(node);
        return;

      // Labels
      case "label":
        this.visitLabel(node);
        break;

      // Preprocessor blocks - process their children
      case "preproc_if":
      case "preproc_elseif":
      case "preproc_else":
        this.visitChildren(node);
        return;

      // Default: visit children
      default:
        this.visitChildren(node);
        return;
    }

    // Visit children for declarations that don't return early
    this.visitChildren(node);
  }

  /** Visit all children of a node */
  private visitChildren(node: SyntaxNode): void {
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (child) this.visitNode(child);
    }
  }

  /** Visit variable declaration */
  private visitVariableDeclaration(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const isLocal = !this.isModuleScope();

    // Find variable_list -> variable_declarator nodes
    for (const vl of this.findChildrenByKind(node, "variable_list")) {
      for (const vd of this.findChildrenByKind(vl, "variable_declarator")) {
        const nameNode = this.findField(vd, "name");
        if (nameNode) {
          const name = this.nodeText(nameNode);
          const definitionRange = this.nodeRange(vd);
          const nameRange = this.nodeRange(nameNode);

          const kind = isLocal ? SymbolKind.LocalVariable : SymbolKind.Variable;

          const symbolId = this.table.createSymbol(
            name,
            kind,
            visibility,
            definitionRange,
            nameRange,
            this.currentScope()
          );

          const typeInfo = this.extractType(vd);
          if (typeInfo) {
            this.table.setTypeInfo(symbolId, typeInfo);
          }
        }
      }
    }
  }

  /** Visit constant declaration */
  private visitConstantDeclaration(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const isLocal = !this.isModuleScope();

    for (const cd of this.findChildrenByKind(node, "constant_declarator")) {
      const nameNode = this.findField(cd, "name");
      if (nameNode) {
        const name = this.nodeText(nameNode);
        const definitionRange = this.nodeRange(cd);
        const nameRange = this.nodeRange(nameNode);

        const valueNode = this.findField(cd, "value");
        const value = valueNode ? this.nodeText(valueNode) : undefined;

        const kind = isLocal ? SymbolKind.LocalConstant : SymbolKind.Constant;

        const symbolId = this.table.createSymbol(
          name,
          kind,
          visibility,
          definitionRange,
          nameRange,
          this.currentScope()
        );

        if (value) {
          this.table.setValue(symbolId, value);
        }

        const typeInfo = this.extractType(cd);
        if (typeInfo) {
          this.table.setTypeInfo(symbolId, typeInfo);
        }
      }
    }
  }

  /** Visit type declaration (User-Defined Type) */
  private visitTypeDeclaration(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      const typeSymbolId = this.table.createSymbol(
        name,
        SymbolKind.UserDefinedType,
        visibility,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      // Process type members
      for (const tm of this.findChildrenByKind(node, "type_member")) {
        const memberNameNode = this.findField(tm, "name");
        if (memberNameNode) {
          const memberName = this.nodeText(memberNameNode);
          const memberDefRange = this.nodeRange(tm);
          const memberNameRange = this.nodeRange(memberNameNode);

          const memberId = this.table.createSymbol(
            memberName,
            SymbolKind.TypeMember,
            Visibility.Public,
            memberDefRange,
            memberNameRange,
            this.currentScope()
          );

          const typeInfo = this.extractType(tm);
          if (typeInfo) {
            this.table.setTypeInfo(memberId, typeInfo);
          }

          this.table.addMember(typeSymbolId, memberId);
        }
      }
    }
  }

  /** Visit enum declaration */
  private visitEnumDeclaration(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      const enumSymbolId = this.table.createSymbol(
        name,
        SymbolKind.Enum,
        visibility,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      // Process enum members
      for (const em of this.findChildrenByKind(node, "enum_member")) {
        const memberNameNode = this.findField(em, "name");
        if (memberNameNode) {
          const memberName = this.nodeText(memberNameNode);
          const memberDefRange = this.nodeRange(em);
          const memberNameRange = this.nodeRange(memberNameNode);

          const valueNode = this.findField(em, "value");
          const value = valueNode ? this.nodeText(valueNode) : undefined;

          const memberId = this.table.createSymbol(
            memberName,
            SymbolKind.EnumMember,
            visibility,
            memberDefRange,
            memberNameRange,
            this.currentScope()
          );

          if (value) {
            this.table.setValue(memberId, value);
          }

          this.table.addMember(enumSymbolId, memberId);
        }
      }
    }
  }

  /** Visit Sub declaration */
  private visitSubDeclaration(node: SyntaxNode): void {
    this.visitProcedure(node, SymbolKind.Sub);
  }

  /** Visit Function declaration */
  private visitFunctionDeclaration(node: SyntaxNode): void {
    this.visitProcedure(node, SymbolKind.Function);
  }

  /** Visit Property declaration */
  private visitPropertyDeclaration(node: SyntaxNode): void {
    const accessor = this.findField(node, "accessor");
    let kind = SymbolKind.PropertyGet;

    if (accessor) {
      const accessorText = this.nodeText(accessor).toUpperCase();
      switch (accessorText) {
        case "GET":
          kind = SymbolKind.PropertyGet;
          break;
        case "LET":
          kind = SymbolKind.PropertyLet;
          break;
        case "SET":
          kind = SymbolKind.PropertySet;
          break;
      }
    }

    this.visitProcedure(node, kind);
  }

  /** Common procedure handling */
  private visitProcedure(node: SyntaxNode, kind: SymbolKind): void {
    const visibility = this.extractVisibility(node);
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      // Create the procedure symbol
      const symbolId = this.table.createSymbol(
        name,
        kind,
        visibility,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      // Extract return type for functions/property get
      if (
        kind === SymbolKind.Function ||
        kind === SymbolKind.PropertyGet
      ) {
        const typeInfo = this.extractType(node);
        if (typeInfo) {
          this.table.setTypeInfo(symbolId, typeInfo);
        }
      }

      // Create a scope for the procedure body
      const procScope = this.pushScope(ScopeKind.Procedure, definitionRange);
      this.table.linkProcedureScope(symbolId, procScope);

      // Extract and register parameters
      const parameters = this.extractParameters(node, procScope);
      this.table.setParameters(symbolId, parameters);

      // Visit the procedure body
      for (const block of this.findChildrenByKind(node, "block")) {
        this.visitChildren(block);
      }

      // Pop the procedure scope
      this.popScope();
    }
  }

  /** Extract parameters from a procedure node */
  private extractParameters(
    node: SyntaxNode,
    procScope: ScopeId
  ): ParameterInfo[] {
    const params: ParameterInfo[] = [];

    for (const pl of this.findChildrenByKind(node, "parameter_list")) {
      for (const param of this.findChildrenByKind(pl, "parameter")) {
        const nameNode = this.findField(param, "name");
        if (nameNode) {
          const name = this.nodeText(nameNode);
          const paramText = this.nodeText(param).toUpperCase();

          const byRef = !paramText.includes("BYVAL");
          const optional = paramText.includes("OPTIONAL");

          const defaultNode = this.findField(param, "default");
          const defaultValue = defaultNode
            ? this.nodeText(defaultNode)
            : undefined;

          const typeInfo = this.extractType(param);
          const paramRange = this.nodeRange(param);
          const nameRange = this.nodeRange(nameNode);

          // Create parameter as a symbol in procedure scope
          const paramId = this.table.createSymbol(
            name,
            SymbolKind.Parameter,
            Visibility.Private,
            paramRange,
            nameRange,
            procScope
          );

          if (typeInfo) {
            this.table.setTypeInfo(paramId, typeInfo);
          }

          params.push({
            name,
            typeInfo,
            byRef,
            optional,
            defaultValue,
            range: paramRange,
            nameRange,
          });
        }
      }
    }

    return params;
  }

  /** Visit Declare statement (API declaration) */
  private visitDeclareStatement(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      const kind = this.hasChildKeyword(node, "function")
        ? SymbolKind.DeclareFunction
        : SymbolKind.DeclareSub;

      const symbolId = this.table.createSymbol(
        name,
        kind,
        visibility,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      // Extract parameters (declares don't create a scope)
      const parameters = this.extractParametersNoScope(node);
      this.table.setParameters(symbolId, parameters);

      // Extract return type for functions
      if (kind === SymbolKind.DeclareFunction) {
        const typeInfo = this.extractType(node);
        if (typeInfo) {
          this.table.setTypeInfo(symbolId, typeInfo);
        }
      }
    }
  }

  /** Extract parameters without creating symbols (for Declare statements) */
  private extractParametersNoScope(node: SyntaxNode): ParameterInfo[] {
    const params: ParameterInfo[] = [];

    for (const pl of this.findChildrenByKind(node, "parameter_list")) {
      for (const param of this.findChildrenByKind(pl, "parameter")) {
        const nameNode = this.findField(param, "name");
        if (nameNode) {
          const name = this.nodeText(nameNode);
          const paramText = this.nodeText(param).toUpperCase();

          const byRef = !paramText.includes("BYVAL");
          const optional = paramText.includes("OPTIONAL");

          const defaultNode = this.findField(param, "default");
          const defaultValue = defaultNode
            ? this.nodeText(defaultNode)
            : undefined;

          const typeInfo = this.extractType(param);

          params.push({
            name,
            typeInfo,
            byRef,
            optional,
            defaultValue,
            range: this.nodeRange(param),
            nameRange: this.nodeRange(nameNode),
          });
        }
      }
    }

    return params;
  }

  /** Visit Event statement */
  private visitEventStatement(node: SyntaxNode): void {
    const visibility = this.extractVisibility(node);
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      const symbolId = this.table.createSymbol(
        name,
        SymbolKind.Event,
        visibility,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      const parameters = this.extractParametersNoScope(node);
      this.table.setParameters(symbolId, parameters);
    }
  }

  /** Visit With statement (creates implicit object scope) */
  private visitWithStatement(node: SyntaxNode): void {
    const range = this.nodeRange(node);

    // Extract the object expression
    const objectNode = this.findField(node, "object");
    const withObject = objectNode ? this.nodeText(objectNode) : undefined;

    const scopeId = this.pushScope(ScopeKind.WithBlock, range);

    if (withObject) {
      this.table.setWithObject(scopeId, withObject);
    }

    // Visit the block
    this.visitChildren(node);

    this.popScope();
  }

  /** Visit For statement */
  private visitForStatement(node: SyntaxNode): void {
    const range = this.nodeRange(node);

    // Create scope for loop variable
    const scopeId = this.pushScope(ScopeKind.ForLoop, range);

    // Register loop variable
    const counterNode = this.findField(node, "counter");
    if (counterNode) {
      const name = this.nodeText(counterNode);
      const nameRange = this.nodeRange(counterNode);

      this.table.createSymbol(
        name,
        SymbolKind.ForLoopVariable,
        Visibility.Private,
        nameRange,
        nameRange,
        scopeId
      );
    }

    // Visit the block
    this.visitChildren(node);

    this.popScope();
  }

  /** Visit For Each statement */
  private visitForEachStatement(node: SyntaxNode): void {
    const range = this.nodeRange(node);

    const scopeId = this.pushScope(ScopeKind.ForEachLoop, range);

    // Register element variable
    const elementNode = this.findField(node, "element");
    if (elementNode) {
      const name = this.nodeText(elementNode);
      const nameRange = this.nodeRange(elementNode);

      this.table.createSymbol(
        name,
        SymbolKind.ForEachVariable,
        Visibility.Private,
        nameRange,
        nameRange,
        scopeId
      );
    }

    this.visitChildren(node);

    this.popScope();
  }

  /** Visit Label */
  private visitLabel(node: SyntaxNode): void {
    const labelNode = node.child(0);
    if (labelNode) {
      const name = this.nodeText(labelNode);
      const range = this.nodeRange(node);
      const nameRange = this.nodeRange(labelNode);

      this.table.createSymbol(
        name,
        SymbolKind.Label,
        Visibility.Private,
        range,
        nameRange,
        this.currentScope()
      );
    }
  }

  /** Visit form block (creates FormControl symbol) */
  private visitFormBlock(node: SyntaxNode): void {
    const nameNode = this.findField(node, "name");

    if (nameNode) {
      const name = this.nodeText(nameNode);
      const definitionRange = this.nodeRange(node);
      const nameRange = this.nodeRange(nameNode);

      // Get the control type (e.g., "VB.TextBox" -> "TextBox")
      const typeNode = this.findField(node, "type");
      let typeInfo: TypeInfo | undefined;
      if (typeNode) {
        const fullType = this.nodeText(typeNode);
        const typeName = fullType.split(".").pop() || fullType;
        typeInfo = {
          name: typeName,
          isArray: false,
          isNew: false,
        };
      }

      const symbolId = this.table.createSymbol(
        name,
        SymbolKind.FormControl,
        Visibility.Private,
        definitionRange,
        nameRange,
        this.currentScope()
      );

      if (typeInfo) {
        this.table.setTypeInfo(symbolId, typeInfo);
      }
    }

    // Recurse into children to find nested form_block elements
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (!child) continue;

      if (child.type === "form_block") {
        this.visitFormBlock(child);
      } else if (child.type === "form_element") {
        for (let j = 0; j < child.childCount; j++) {
          const inner = child.child(j);
          if (inner && inner.type === "form_block") {
            this.visitFormBlock(inner);
          }
        }
      }
    }
  }

  // ==========================================
  // Second Pass: Reference Collection
  // ==========================================

  /** Collect references by walking all identifier nodes */
  private collectReferences(node: SyntaxNode): void {
    switch (node.type) {
      // Skip nodes that contain declarations
      case "form_property_line":
      case "form_property_block":
      case "module_config":
      case "module_config_element":
      case "form_block":
        return;

      // Scope-entering constructs
      case "sub_declaration":
      case "function_declaration":
      case "property_declaration":
        this.collectReferencesInProcedure(node);
        return;

      case "with_statement":
        this.collectReferencesInWith(node);
        return;

      case "for_statement":
      case "for_each_statement":
        this.collectReferencesInFor(node);
        return;

      case "identifier":
        this.tryAddReference(node);
        break;
    }

    // Recurse into all children
    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (child) this.collectReferences(child);
    }
  }

  /** Collect references within a procedure */
  private collectReferencesInProcedure(node: SyntaxNode): void {
    const nameNode = this.findField(node, "name");
    if (nameNode) {
      const name = this.nodeText(nameNode);
      const scopeId = this.findProcedureScope(name);
      if (scopeId !== undefined) {
        this.scopeStack.push(scopeId);
      }
    }

    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (!child) continue;

      // Skip parameter list - parameters are declarations
      if (child.type === "parameter_list") continue;

      // Skip the procedure name itself
      const fieldNameNode = this.findField(node, "name");
      if (fieldNameNode && child.id === fieldNameNode.id) continue;

      this.collectReferences(child);
    }

    if (this.scopeStack.length > 1) {
      this.scopeStack.pop();
    }
  }

  /** Collect references within a With statement */
  private collectReferencesInWith(node: SyntaxNode): void {
    const range = this.nodeRange(node);
    const scopeId = this.findScopeAtRange(range, ScopeKind.WithBlock);
    if (scopeId !== undefined) {
      this.scopeStack.push(scopeId);
    }

    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (child) this.collectReferences(child);
    }

    if (this.scopeStack.length > 1) {
      this.scopeStack.pop();
    }
  }

  /** Collect references within a For loop */
  private collectReferencesInFor(node: SyntaxNode): void {
    const range = this.nodeRange(node);
    const scopeKind =
      node.type === "for_each_statement"
        ? ScopeKind.ForEachLoop
        : ScopeKind.ForLoop;
    const scopeId = this.findScopeAtRange(range, scopeKind);
    if (scopeId !== undefined) {
      this.scopeStack.push(scopeId);
    }

    for (let i = 0; i < node.childCount; i++) {
      const child = node.child(i);
      if (!child) continue;

      // Skip the loop variable declaration
      const varNode = this.findField(node, "variable");
      if (varNode && child.id === varNode.id) continue;

      this.collectReferences(child);
    }

    if (this.scopeStack.length > 1) {
      this.scopeStack.pop();
    }
  }

  /** Find the scope for a procedure by looking up the symbol */
  private findProcedureScope(name: string): ScopeId | undefined {
    const nameLower = name.toLowerCase();
    for (const symbol of this.table.allSymbols()) {
      if (
        symbol.name.toLowerCase() === nameLower &&
        [
          SymbolKind.Sub,
          SymbolKind.Function,
          SymbolKind.PropertyGet,
          SymbolKind.PropertyLet,
          SymbolKind.PropertySet,
        ].includes(symbol.kind)
      ) {
        for (const scope of this.table.allScopes()) {
          if (scope.definingSymbol === symbol.id) {
            return scope.id;
          }
        }
      }
    }
    return undefined;
  }

  /** Find a scope at a specific range with a specific kind */
  private findScopeAtRange(
    range: SourceRange,
    kind: ScopeKind
  ): ScopeId | undefined {
    for (const scope of this.table.allScopes()) {
      if (
        scope.kind === kind &&
        scope.range.start.line === range.start.line &&
        scope.range.start.column === range.start.column
      ) {
        return scope.id;
      }
    }
    return undefined;
  }

  /** Try to add a reference for an identifier node */
  private tryAddReference(node: SyntaxNode): void {
    // Check if this identifier is part of a declaration (skip those)
    if (this.isDeclarationName(node)) {
      return;
    }

    const name = this.nodeText(node);
    const range = this.nodeRange(node);
    const scopeId = this.currentScope();

    // Check if this is an assignment target
    const isAssignment = this.isAssignmentTarget(node);

    // Try to resolve this identifier to a symbol
    const symbol = this.table.lookupSymbol(name, scopeId);
    if (symbol) {
      this.table.addReference(symbol.id, range, scopeId, isAssignment);
    }
  }

  /** Check if an identifier node is the name part of a declaration */
  private isDeclarationName(node: SyntaxNode): boolean {
    const parent = node.parent;
    if (!parent) return false;

    const declarationTypes = [
      "variable_declarator",
      "constant_declarator",
      "enum_member",
      "type_member",
      "parameter",
      "sub_declaration",
      "function_declaration",
      "property_declaration",
      "type_declaration",
      "enum_declaration",
      "declare_statement",
      "event_statement",
      "for_statement",
      "for_each_statement",
    ];

    if (declarationTypes.includes(parent.type)) {
      const nameNode = parent.childForFieldName("name");
      if (nameNode && nameNode.id === node.id) {
        return true;
      }
      const varNode = parent.childForFieldName("variable");
      if (varNode && varNode.id === node.id) {
        return true;
      }
    }

    if (parent.type === "label") {
      return true;
    }

    return false;
  }

  /** Check if an identifier is an assignment target */
  private isAssignmentTarget(node: SyntaxNode): boolean {
    const parent = node.parent;
    if (!parent) return false;

    if (
      parent.type === "assignment_statement" ||
      parent.type === "set_statement"
    ) {
      const target = parent.childForFieldName("target");
      if (target) {
        if (target.id === node.id) return true;
        return this.isDescendantOf(node, target);
      }
    }

    return false;
  }

  /** Check if node is a descendant of ancestor */
  private isDescendantOf(node: SyntaxNode, ancestor: SyntaxNode): boolean {
    let current = node.parent;
    while (current) {
      if (current.id === ancestor.id) return true;
      current = current.parent;
    }
    return false;
  }
}

/**
 * Build a symbol table from source code and tree-sitter tree
 */
export function buildSymbolTable(
  filePath: string,
  source: string,
  tree: Tree
): SymbolTable {
  const builder = new SymbolTableBuilder(filePath, source);
  return builder.build(tree);
}
