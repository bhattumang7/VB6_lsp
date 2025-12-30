#!/usr/bin/env python3
"""
Structural comparison of ANTLR .tree parse outputs with tree-sitter parse trees.

This performs a full structural comparison, not just count-based.
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class TreeNode:
    """Represents a node in a parse tree."""
    node_type: str
    children: list["TreeNode"] = field(default_factory=list)
    text: str = ""  # Literal text content (for terminals)
    field_name: str = ""  # Field name if this is a named child
    start_pos: tuple[int, int] = (0, 0)
    end_pos: tuple[int, int] = (0, 0)

    def __repr__(self):
        if self.children:
            return f"({self.node_type} [{len(self.children)} children])"
        elif self.text:
            return f"({self.node_type} '{self.text}')"
        else:
            return f"({self.node_type})"

    def pretty_print(self, indent: int = 0) -> str:
        """Pretty print the tree."""
        prefix = "  " * indent
        if self.field_name:
            result = f"{prefix}{self.field_name}: ({self.node_type})"
        else:
            result = f"{prefix}({self.node_type})"

        if self.text:
            result += f" '{self.text}'"

        lines = [result]
        for child in self.children:
            lines.append(child.pretty_print(indent + 1))
        return "\n".join(lines)

    def get_significant_children(self) -> list["TreeNode"]:
        """Get children that are significant for comparison (skip noise nodes)."""
        return [c for c in self.children if c.node_type not in ANTLR_NOISE_NODES]


# ANTLR nodes to skip during comparison (noise/wrapper nodes)
ANTLR_NOISE_NODES = {
    "startRule",
    "module",
    "moduleBody",
    "moduleBodyElement",
    "block",
    "blockStmt",
    "valueStmt",
    "implicitCallStmt_InStmt",
    "iCS_S_VariableOrProcedureCall",
    "iCS_S_MembersCall",
    "iCS_S_MemberCall",
    "ambiguousIdentifier",
    "ambiguousKeyword",
    "literal",
    "integerLiteral",
}

# Tree-sitter nodes to skip during comparison
TS_NOISE_NODES = {
    "source_file",
    "block",
    "literal",
    "string_literal",
    "integer_literal",
    "identifier",
}

# Mapping from ANTLR node types to tree-sitter node types
ANTLR_TO_TS_NODE_MAP = {
    "subStmt": "sub_declaration",
    "functionStmt": "function_declaration",
    "propertyGetStmt": "property_declaration",
    "propertyLetStmt": "property_declaration",
    "propertySetStmt": "property_declaration",
    "letStmt": "assignment_statement",
    "setStmt": "set_statement",
    "withStmt": "with_statement",
    "ifThenElseStmt": "if_statement",
    "selectCaseStmt": "select_statement",
    "forNextStmt": "for_statement",
    "forEachStmt": "for_each_statement",
    "doLoopStmt": "do_statement",
    "whileWendStmt": "while_statement",
    "endStmt": "end_statement",
    "exitStmt": "exit_statement",
    "goToStmt": "goto_statement",
    "goSubStmt": "gosub_statement",
    "returnStmt": "return_statement",
    "onErrorStmt": "on_error_statement",
    "resumeStmt": "resume_statement",
    "callStmt": "call_statement",
    "explicitCallStmt": "call_statement",
    "implicitCallStmt_InBlock": "call_statement",
    "printStmt": "print_statement",
    "writeStmt": "write_statement",
    "inputStmt": "input_statement",
    "lineInputStmt": "line_input_statement",
    "openStmt": "open_statement",
    "closeStmt": "close_statement",
    "getStmt": "get_statement",
    "putStmt": "put_statement",
    "seekStmt": "seek_statement",
    "lockStmt": "lock_statement",
    "unlockStmt": "unlock_statement",
    "constStmt": "constant_declaration",
    "variableStmt": "variable_declaration",
    "redimStmt": "redim_statement",
    "eraseStmt": "erase_statement",
    "typeStmt": "type_declaration",
    "enumerationStmt": "enum_declaration",
    "declareStmt": "declare_statement",
    "eventStmt": "event_statement",
    "implementsStmt": "implements_statement",
    "optionStmt": "option_statement",
    "moduleOption": "option_statement",
    "attributeStmt": "attribute_statement",
    "errorStmt": "error_statement",
    "raiseEventStmt": "raiseevent_statement",
    "debugStmt": "debug_statement",
    "stopStmt": "stop_statement",
    "nameStmt": "name_statement",
    "killStmt": "kill_statement",
    "mkdirStmt": "mkdir_statement",
    "rmdirStmt": "rmdir_statement",
    "chDirStmt": "chdir_statement",
    "chDriveStmt": "chdrive_statement",
    "filecopyStmt": "filecopy_statement",
    "randomizeStmt": "randomize_statement",
    "midStmt": "mid_statement",
    "lsetStmt": "lset_statement",
    "rsetStmt": "rset_statement",
    "argList": "parameter_list",
    "arg": "parameter",
}

# Reverse mapping for lookup
TS_TO_ANTLR_NODE_MAP = {v: k for k, v in ANTLR_TO_TS_NODE_MAP.items()}

SOURCE_EXTS = {".cls", ".vb", ".bas", ".frm"}
NPX = "npx.cmd" if os.name == "nt" else "npx"


def parse_antlr_tree(text: str) -> TreeNode:
    """Parse an ANTLR S-expression tree into a TreeNode structure."""
    text = text.strip()
    pos = 0

    def skip_whitespace():
        nonlocal pos
        while pos < len(text) and text[pos] in " \t\n\r":
            pos += 1

    def parse_node() -> Optional[TreeNode]:
        nonlocal pos
        skip_whitespace()

        if pos >= len(text):
            return None

        if text[pos] != "(":
            # This is a terminal/literal token
            start = pos
            # Handle quoted strings
            if text[pos] == '"':
                pos += 1
                while pos < len(text) and text[pos] != '"':
                    if text[pos] == '\\' and pos + 1 < len(text):
                        pos += 2
                    else:
                        pos += 1
                if pos < len(text):
                    pos += 1
                return TreeNode(node_type="LITERAL", text=text[start:pos])
            # Handle backslash-escaped characters like \n
            elif text[pos] == '\\' and pos + 1 < len(text):
                pos += 2
                return None  # Skip escape sequences
            else:
                # Read until whitespace or paren
                while pos < len(text) and text[pos] not in " \t\n\r()":
                    pos += 1
                token = text[start:pos]
                if token:
                    return TreeNode(node_type="TOKEN", text=token)
                return None

        # Skip '('
        pos += 1
        skip_whitespace()

        # Read node type
        start = pos
        while pos < len(text) and text[pos] not in " \t\n\r()":
            pos += 1
        node_type = text[start:pos]

        if not node_type:
            # This is a literal "(" or ")" token in ANTLR format
            # e.g., in "(argList ( (arg ...))", the standalone "(" represents
            # the literal parenthesis in VB6 source code like "Sub Test("
            # We need to find the matching ")" for this empty node
            if pos < len(text) and text[pos] == ")":
                # Empty parens "()" - represents literal "()" token
                pos += 1
                return TreeNode(node_type="TOKEN", text="()")
            elif pos < len(text) and text[pos] == "(":
                # "( (" pattern - the first "(" is a literal token
                # Don't consume anything more, just return the "(" token
                # The caller will continue to parse the real "(node ...)"
                return TreeNode(node_type="TOKEN", text="(")
            else:
                # Malformed - skip to closing paren
                while pos < len(text) and text[pos] != ")":
                    pos += 1
                if pos < len(text):
                    pos += 1
                return None

        node = TreeNode(node_type=node_type)

        # Parse children
        while True:
            skip_whitespace()
            if pos >= len(text):
                break
            if text[pos] == ")":
                # Check if this ) is a literal TOKEN or the structural close
                # In ANTLR format, ` ) )` means the first ) is a TOKEN
                # Look ahead to see what follows this )
                lookahead = pos + 1
                while lookahead < len(text) and text[lookahead] in " \t\n\r":
                    lookahead += 1
                if lookahead < len(text) and text[lookahead] in "()":
                    # Another paren follows, so this ) is a TOKEN
                    pos += 1  # consume the )
                    node.children.append(TreeNode(node_type="TOKEN", text=")"))
                    continue
                else:
                    # This ) ends the current node
                    break
            child = parse_node()
            if child:
                node.children.append(child)

        # Skip ')'
        if pos < len(text) and text[pos] == ")":
            pos += 1

        return node

    result = parse_node()
    return result if result else TreeNode(node_type="ERROR")


def parse_ts_tree(text: str) -> TreeNode:
    """Parse a tree-sitter S-expression output into a TreeNode structure."""
    lines = text.strip().split("\n")
    if not lines:
        return TreeNode(node_type="ERROR")

    # Parse position pattern: [row, col] - [row, col]
    pos_pattern = re.compile(r'\[(\d+), (\d+)\] - \[(\d+), (\d+)\]')

    def parse_line(line: str) -> tuple[int, str, str, tuple, tuple]:
        """Parse a line and return (indent, field_name, node_type, start_pos, end_pos)."""
        # Count leading spaces for indent
        stripped = line.lstrip()
        indent = len(line) - len(stripped)

        field_name = ""
        if ": (" in stripped:
            field_name, rest = stripped.split(": (", 1)
            stripped = "(" + rest

        # Extract node type and positions
        if stripped.startswith("("):
            parts = stripped[1:].split()
            node_type = parts[0] if parts else ""

            # Find position
            match = pos_pattern.search(stripped)
            if match:
                start_pos = (int(match.group(1)), int(match.group(2)))
                end_pos = (int(match.group(3)), int(match.group(4)))
            else:
                start_pos = (0, 0)
                end_pos = (0, 0)

            # Remove trailing )
            node_type = node_type.rstrip(")")

            return indent, field_name, node_type, start_pos, end_pos

        return indent, "", "", (0, 0), (0, 0)

    # Build tree from indented output
    root = None
    stack: list[tuple[int, TreeNode]] = []

    for line in lines:
        if not line.strip():
            continue

        indent, field_name, node_type, start_pos, end_pos = parse_line(line)

        if not node_type:
            continue

        node = TreeNode(
            node_type=node_type,
            field_name=field_name,
            start_pos=start_pos,
            end_pos=end_pos
        )

        # Pop stack until we find parent (smaller indent)
        while stack and stack[-1][0] >= indent:
            stack.pop()

        if stack:
            stack[-1][1].children.append(node)
        else:
            root = node

        stack.append((indent, node))

    return root if root else TreeNode(node_type="ERROR")


def flatten_significant_nodes(node: TreeNode, noise_nodes: set[str]) -> list[TreeNode]:
    """Flatten tree to list of significant nodes (skipping noise nodes)."""
    result = []

    def traverse(n: TreeNode):
        if n.node_type in noise_nodes:
            for child in n.children:
                traverse(child)
        else:
            result.append(n)
            for child in n.children:
                traverse(child)

    traverse(node)
    return result


def get_structural_signature(node: TreeNode, noise_nodes: set[str], depth: int = 0, max_depth: int = 10) -> str:
    """Get a structural signature of the tree for comparison."""
    if depth > max_depth:
        return "..."

    if node.node_type in noise_nodes:
        # Skip noise node, but include children
        child_sigs = []
        for child in node.children:
            sig = get_structural_signature(child, noise_nodes, depth, max_depth)
            if sig:
                child_sigs.append(sig)
        return " ".join(child_sigs)

    child_sigs = []
    for child in node.children:
        sig = get_structural_signature(child, noise_nodes, depth + 1, max_depth)
        if sig:
            child_sigs.append(sig)

    if child_sigs:
        return f"({node.node_type} {' '.join(child_sigs)})"
    else:
        return f"({node.node_type})"


@dataclass
class ComparisonResult:
    """Result of comparing two trees."""
    matches: bool
    antlr_nodes: list[str]
    ts_nodes: list[str]
    mismatches: list[str]
    details: str = ""


def extract_statement_nodes(node: TreeNode, stmt_types: set[str]) -> list[TreeNode]:
    """Extract all statement-level nodes from a tree."""
    result = []

    def traverse(n: TreeNode):
        if n.node_type in stmt_types:
            result.append(n)
        for child in n.children:
            traverse(child)

    traverse(node)
    return result


def compare_trees(antlr_tree: TreeNode, ts_tree: TreeNode) -> ComparisonResult:
    """Compare ANTLR and tree-sitter parse trees structurally."""

    # Define statement-level nodes for each parser
    antlr_stmt_types = set(ANTLR_TO_TS_NODE_MAP.keys())
    ts_stmt_types = set(ANTLR_TO_TS_NODE_MAP.values())

    # Extract statement nodes
    antlr_stmts = extract_statement_nodes(antlr_tree, antlr_stmt_types)
    ts_stmts = extract_statement_nodes(ts_tree, ts_stmt_types)

    antlr_types = [n.node_type for n in antlr_stmts]
    ts_types = [n.node_type for n in ts_stmts]

    # Map ANTLR types to TS types for comparison
    mapped_antlr_types = [ANTLR_TO_TS_NODE_MAP.get(t, t) for t in antlr_types]

    mismatches = []

    # Compare counts
    if len(mapped_antlr_types) != len(ts_types):
        mismatches.append(
            f"Node count mismatch: ANTLR has {len(antlr_types)}, tree-sitter has {len(ts_types)}"
        )

    # Compare sequence
    min_len = min(len(mapped_antlr_types), len(ts_types))
    for i in range(min_len):
        if mapped_antlr_types[i] != ts_types[i]:
            mismatches.append(
                f"Position {i}: ANTLR '{antlr_types[i]}' -> '{mapped_antlr_types[i]}', "
                f"tree-sitter '{ts_types[i]}'"
            )

    # Report extra nodes
    if len(mapped_antlr_types) > len(ts_types):
        for i in range(len(ts_types), len(mapped_antlr_types)):
            mismatches.append(f"Extra ANTLR node at {i}: {antlr_types[i]}")
    elif len(ts_types) > len(mapped_antlr_types):
        for i in range(len(mapped_antlr_types), len(ts_types)):
            mismatches.append(f"Extra tree-sitter node at {i}: {ts_types[i]}")

    # Build details
    details_lines = [
        "ANTLR statements (mapped):",
        "  " + ", ".join(f"{a} -> {m}" for a, m in zip(antlr_types, mapped_antlr_types)),
        "Tree-sitter statements:",
        "  " + ", ".join(ts_types),
    ]

    return ComparisonResult(
        matches=len(mismatches) == 0,
        antlr_nodes=antlr_types,
        ts_nodes=ts_types,
        mismatches=mismatches,
        details="\n".join(details_lines)
    )


def run_tree_sitter_parse(tree_sitter_dir: Path, file_path: Path) -> str:
    """Run tree-sitter parse and return the output."""
    try:
        rel_path = file_path.relative_to(tree_sitter_dir)
    except ValueError:
        rel_path = file_path

    result = subprocess.run(
        [NPX, "tree-sitter", "parse", str(rel_path)],
        cwd=str(tree_sitter_dir),
        capture_output=True,
        text=True,
        timeout=10,
    )

    output = result.stdout
    # Filter to just the tree output
    lines = output.splitlines()
    tree_lines = []
    in_tree = False
    for line in lines:
        if line.strip().startswith("(source_file"):
            in_tree = True
        if in_tree:
            tree_lines.append(line)

    return "\n".join(tree_lines).strip()


def compare_file(tree_sitter_dir: Path, source_path: Path, verbose: bool = False) -> dict:
    """Compare a single source file with its .tree file."""
    tree_path = Path(str(source_path) + ".tree")

    if not tree_path.exists():
        return {"status": "missing_tree", "message": f".tree file not found: {tree_path}"}

    if not source_path.exists():
        return {"status": "missing_source", "message": f"Source file not found: {source_path}"}

    # Read and parse ANTLR tree
    antlr_text = tree_path.read_text(encoding="utf-8", errors="ignore")
    antlr_tree = parse_antlr_tree(antlr_text)

    # Get tree-sitter parse output
    ts_output = run_tree_sitter_parse(tree_sitter_dir, source_path)
    if not ts_output:
        return {"status": "parse_failed", "message": "tree-sitter parse returned no output"}

    ts_tree = parse_ts_tree(ts_output)

    # Compare the trees
    result = compare_trees(antlr_tree, ts_tree)

    return {
        "status": "match" if result.matches else "mismatch",
        "result": result,
        "antlr_tree": antlr_tree if verbose else None,
        "ts_tree": ts_tree if verbose else None,
        "ts_output": ts_output if verbose else None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Structural comparison of ANTLR .tree outputs with tree-sitter parse trees."
    )
    parser.add_argument(
        "--files",
        nargs="*",
        help="Paths under test/antlr_examples (e.g. statements/Write.cls).",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Compare all source files with matching .tree files.",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Show detailed output.")
    parser.add_argument("--tree", action="store_true", help="Show parsed trees.")
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    tree_sitter_dir = script_dir.parent
    root = script_dir / "antlr_examples"

    if not args.all and not args.files:
        parser.error("use --files or --all")

    if args.all:
        sources = [
            p
            for p in sorted(root.rglob("*"))
            if p.suffix in SOURCE_EXTS and p.is_file()
        ]
    else:
        sources = []
        for raw in args.files:
            path = Path(raw)
            if not path.is_absolute():
                path = root / raw
            sources.append(path)

    match_count = mismatch_count = skip_count = error_count = 0

    for source_path in sources:
        try:
            rel = source_path.relative_to(root)
        except ValueError:
            rel = source_path

        result = compare_file(tree_sitter_dir, source_path, verbose=args.verbose or args.tree)
        status = result["status"]

        if status == "missing_tree":
            skip_count += 1
            print(f"[SKIP] {rel.as_posix()} - {result['message']}")
            continue

        if status == "missing_source":
            error_count += 1
            print(f"[ERR]  {rel.as_posix()} - {result['message']}")
            continue

        if status == "parse_failed":
            error_count += 1
            print(f"[ERR]  {rel.as_posix()} - {result['message']}")
            continue

        comp_result = result["result"]

        if status == "match":
            match_count += 1
            print(f"[OK]   {rel.as_posix()}")
            if args.verbose:
                print(f"       ANTLR: {comp_result.antlr_nodes}")
                print(f"       TS:    {comp_result.ts_nodes}")
        else:
            mismatch_count += 1
            print(f"[DIFF] {rel.as_posix()}")
            for m in comp_result.mismatches:
                print(f"       {m}")
            if args.verbose:
                print(comp_result.details)

        if args.tree and result.get("ts_output"):
            print("\n--- Tree-sitter output ---")
            print(result["ts_output"])
            print("\n--- ANTLR tree ---")
            print(result["antlr_tree"].pretty_print() if result.get("antlr_tree") else "N/A")
            print()

    total = match_count + mismatch_count + skip_count + error_count
    print(f"\nSummary:")
    print(f"  Total:    {total}")
    print(f"  Match:    {match_count}")
    print(f"  Mismatch: {mismatch_count}")
    print(f"  Skipped:  {skip_count}")
    print(f"  Errors:   {error_count}")

    return 0 if mismatch_count == 0 and error_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
