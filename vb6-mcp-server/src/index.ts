#!/usr/bin/env node
/**
 * VB6 MCP Server
 *
 * Provides VB6 code analysis capabilities to Claude via MCP.
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { spawn } from "child_process";
import Parser from "tree-sitter";
import Vb6 from "tree-sitter-vb6";

import {
  buildSymbolTable,
  formatSignature,
  kindDisplayName,
  type SymbolTable,
  type SourcePosition,
} from "./analysis/index.js";

// Initialize tree-sitter parser
const parser = new Parser();
parser.setLanguage(Vb6);

// Cache of parsed documents
const documentCache = new Map<
  string,
  { source: string; symbolTable: SymbolTable }
>();

/**
 * Call the Rust vb6-lsp binary for resource file operations
 */
async function callRustCli(
  command: string,
  args: string[]
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    // Try to find the vb6-lsp binary
    // First check in the parent directory's target/release or target/debug
    const possiblePaths = [
      "../target/release/vb6-lsp",
      "../target/debug/vb6-lsp",
      "../target/release/vb6-lsp.exe",
      "../target/debug/vb6-lsp.exe",
    ];

    let binaryPath = "vb6-lsp"; // Fallback to PATH
    for (const path of possiblePaths) {
      if (existsSync(path)) {
        binaryPath = path;
        break;
      }
    }

    const proc = spawn(binaryPath, [command, ...args]);
    let stdout = "";
    let stderr = "";

    proc.stdout.on("data", (data) => {
      stdout += data.toString();
    });

    proc.stderr.on("data", (data) => {
      stderr += data.toString();
    });

    proc.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(new Error(`CLI exited with code ${code}: ${stderr}`));
      }
    });

    proc.on("error", (error) => {
      reject(
        new Error(
          `Failed to spawn vb6-lsp binary (${binaryPath}): ${error.message}`
        )
      );
    });
  });
}

/**
 * Get or create a symbol table for a file
 */
function getSymbolTable(filePath: string): SymbolTable | null {
  // Check cache first
  const cached = documentCache.get(filePath);

  // Read file
  if (!existsSync(filePath)) {
    return null;
  }

  const source = readFileSync(filePath, "utf-8");

  // Return cached if source unchanged
  if (cached && cached.source === source) {
    return cached.symbolTable;
  }

  // Parse and build symbol table
  const tree = parser.parse(source);
  const symbolTable = buildSymbolTable(filePath, source, tree);

  // Cache it
  documentCache.set(filePath, { source, symbolTable });

  return symbolTable;
}

/**
 * Tool definitions
 */
const tools: Tool[] = [
  {
    name: "vb6_get_symbols",
    description:
      "Get all symbols (variables, functions, classes, etc.) defined in a VB6 file. Returns a list of symbol definitions with their names, kinds, types, and locations.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file (.bas, .cls, .frm)",
        },
      },
      required: ["file_path"],
    },
  },
  {
    name: "vb6_find_definition",
    description:
      "Find the definition of a symbol at a specific position in a VB6 file. Returns the location and details of where the symbol is defined.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file",
        },
        line: {
          type: "number",
          description: "Line number (0-indexed)",
        },
        column: {
          type: "number",
          description: "Column number (0-indexed)",
        },
      },
      required: ["file_path", "line", "column"],
    },
  },
  {
    name: "vb6_find_references",
    description:
      "Find all references to a symbol at a specific position. Returns all locations where the symbol is used throughout the file.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file",
        },
        line: {
          type: "number",
          description: "Line number (0-indexed)",
        },
        column: {
          type: "number",
          description: "Column number (0-indexed)",
        },
      },
      required: ["file_path", "line", "column"],
    },
  },
  {
    name: "vb6_get_hover",
    description:
      "Get hover information (type, signature, documentation) for a symbol at a specific position.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file",
        },
        line: {
          type: "number",
          description: "Line number (0-indexed)",
        },
        column: {
          type: "number",
          description: "Column number (0-indexed)",
        },
      },
      required: ["file_path", "line", "column"],
    },
  },
  {
    name: "vb6_get_completions",
    description:
      "Get code completion suggestions at a specific position. Returns symbols that are visible/accessible at that location.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file",
        },
        line: {
          type: "number",
          description: "Line number (0-indexed)",
        },
        column: {
          type: "number",
          description: "Column number (0-indexed)",
        },
      },
      required: ["file_path", "line", "column"],
    },
  },
  {
    name: "vb6_get_diagnostics",
    description:
      "Get parse errors and warnings for a VB6 file. Returns syntax errors found during parsing.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the VB6 file",
        },
      },
      required: ["file_path"],
    },
  },
  {
    name: "vb6_read_res_file",
    description:
      "Read and parse a VB6 resource file (.res). Returns all resource entries including bitmaps, icons, strings, and custom resources with their types, names, and metadata.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the .res file",
        },
      },
      required: ["file_path"],
    },
  },
  {
    name: "vb6_write_res_file",
    description:
      "Write resource entries to a VB6 resource file (.res). Creates or overwrites the file with the provided resources.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the output .res file",
        },
        resources: {
          type: "array",
          description: "Array of resource entries to write",
          items: {
            type: "object",
            properties: {
              resource_type: {
                type: "string",
                description: "Resource type (e.g., 'Bitmap', 'Icon', 'String', or 'Named(\"CUSTOM\")')",
              },
              name: {
                type: "object",
                description: "Resource identifier (either numeric ID or string name)",
                properties: {
                  type: {
                    type: "string",
                    enum: ["Id", "Name"],
                  },
                  value: {
                    description: "Numeric ID or string name",
                  },
                },
                required: ["type", "value"],
              },
              language_id: {
                type: "number",
                description: "Language identifier (e.g., 0x0409 for US English)",
              },
              data_base64: {
                type: "string",
                description: "Resource data encoded as base64",
              },
            },
            required: ["resource_type", "name", "language_id", "data_base64"],
          },
        },
      },
      required: ["file_path", "resources"],
    },
  },
  {
    name: "vb6_get_string_table",
    description:
      "Parse a string table resource from a .res file. Returns individual string entries with their IDs and values.",
    inputSchema: {
      type: "object" as const,
      properties: {
        file_path: {
          type: "string",
          description: "Absolute path to the .res file",
        },
        block_id: {
          type: "number",
          description: "String table block ID (resource name must be numeric)",
        },
      },
      required: ["file_path", "block_id"],
    },
  },
];

/**
 * Handle tool calls
 */
async function handleToolCall(
  name: string,
  args: Record<string, unknown>
): Promise<unknown> {
  switch (name) {
    case "vb6_get_symbols": {
      const filePath = args.file_path as string;
      const symbolTable = getSymbolTable(filePath);

      if (!symbolTable) {
        return { error: `File not found: ${filePath}` };
      }

      const symbols = symbolTable.allSymbols().map((s) => ({
        name: s.name,
        kind: s.kind,
        kindDisplayName: kindDisplayName(s.kind),
        visibility: s.visibility,
        signature: formatSignature(s),
        location: {
          line: s.nameRange.start.line,
          column: s.nameRange.start.column,
          endLine: s.nameRange.end.line,
          endColumn: s.nameRange.end.column,
        },
        type: s.typeInfo
          ? {
              name: s.typeInfo.name,
              isArray: s.typeInfo.isArray,
            }
          : undefined,
        parameters:
          s.parameters.length > 0
            ? s.parameters.map((p) => ({
                name: p.name,
                type: p.typeInfo?.name,
                byRef: p.byRef,
                optional: p.optional,
              }))
            : undefined,
        memberCount: s.members.length > 0 ? s.members.length : undefined,
      }));

      return {
        file: filePath,
        symbolCount: symbols.length,
        symbols,
      };
    }

    case "vb6_find_definition": {
      const filePath = args.file_path as string;
      const line = args.line as number;
      const column = args.column as number;
      const symbolTable = getSymbolTable(filePath);

      if (!symbolTable) {
        return { error: `File not found: ${filePath}` };
      }

      const pos: SourcePosition = { line, column };
      const symbol = symbolTable.symbolAtPosition(pos);

      if (!symbol) {
        return { found: false, message: "No symbol found at position" };
      }

      return {
        found: true,
        symbol: {
          name: symbol.name,
          kind: symbol.kind,
          signature: formatSignature(symbol),
          definitionLocation: {
            file: filePath,
            line: symbol.nameRange.start.line,
            column: symbol.nameRange.start.column,
            endLine: symbol.nameRange.end.line,
            endColumn: symbol.nameRange.end.column,
          },
        },
      };
    }

    case "vb6_find_references": {
      const filePath = args.file_path as string;
      const line = args.line as number;
      const column = args.column as number;
      const symbolTable = getSymbolTable(filePath);

      if (!symbolTable) {
        return { error: `File not found: ${filePath}` };
      }

      const pos: SourcePosition = { line, column };
      const ranges = symbolTable.findAllReferences(pos);

      if (ranges.length === 0) {
        return { found: false, message: "No symbol found at position" };
      }

      return {
        found: true,
        referenceCount: ranges.length,
        references: ranges.map((r) => ({
          file: filePath,
          line: r.start.line,
          column: r.start.column,
          endLine: r.end.line,
          endColumn: r.end.column,
        })),
      };
    }

    case "vb6_get_hover": {
      const filePath = args.file_path as string;
      const line = args.line as number;
      const column = args.column as number;
      const symbolTable = getSymbolTable(filePath);

      if (!symbolTable) {
        return { error: `File not found: ${filePath}` };
      }

      const pos: SourcePosition = { line, column };
      const symbol = symbolTable.symbolAtPosition(pos);

      if (!symbol) {
        return { found: false };
      }

      return {
        found: true,
        contents: {
          signature: formatSignature(symbol),
          kind: kindDisplayName(symbol.kind),
          documentation: symbol.documentation,
          type: symbol.typeInfo?.name,
        },
      };
    }

    case "vb6_get_completions": {
      const filePath = args.file_path as string;
      const line = args.line as number;
      const column = args.column as number;
      const symbolTable = getSymbolTable(filePath);

      if (!symbolTable) {
        return { error: `File not found: ${filePath}` };
      }

      const pos: SourcePosition = { line, column };
      const symbols = symbolTable.visibleSymbols(pos);

      return {
        completions: symbols.map((s) => ({
          label: s.name,
          kind: s.kind,
          detail: formatSignature(s),
          documentation: s.documentation,
        })),
      };
    }

    case "vb6_get_diagnostics": {
      const filePath = args.file_path as string;

      if (!existsSync(filePath)) {
        return { error: `File not found: ${filePath}` };
      }

      const source = readFileSync(filePath, "utf-8");
      const tree = parser.parse(source);

      // Collect ERROR nodes from the parse tree
      const errors: Array<{
        message: string;
        line: number;
        column: number;
        endLine: number;
        endColumn: number;
      }> = [];

      function collectErrors(node: Parser.SyntaxNode): void {
        if (node.type === "ERROR" || node.isMissing) {
          errors.push({
            message: node.isMissing
              ? `Missing ${node.type}`
              : `Syntax error: unexpected ${node.text.slice(0, 20)}${node.text.length > 20 ? "..." : ""}`,
            line: node.startPosition.row,
            column: node.startPosition.column,
            endLine: node.endPosition.row,
            endColumn: node.endPosition.column,
          });
        }

        for (let i = 0; i < node.childCount; i++) {
          const child = node.child(i);
          if (child) collectErrors(child);
        }
      }

      collectErrors(tree.rootNode);

      return {
        file: filePath,
        errorCount: errors.length,
        diagnostics: errors,
      };
    }

    case "vb6_read_res_file": {
      const filePath = args.file_path as string;

      if (!existsSync(filePath)) {
        return { error: `File not found: ${filePath}` };
      }

      try {
        // Call Rust CLI to read the .res file
        const { stdout } = await callRustCli("read-res", [filePath]);
        const result = JSON.parse(stdout);

        return {
          file: filePath,
          resourceCount: result.resources?.length || 0,
          resources: result.resources || [],
        };
      } catch (error) {
        return {
          error: `Failed to read .res file: ${error instanceof Error ? error.message : String(error)}`,
        };
      }
    }

    case "vb6_write_res_file": {
      const filePath = args.file_path as string;
      const resources = args.resources as Array<{
        resource_type: string;
        name: { type: string; value: number | string };
        language_id: number;
        data_base64: string;
      }>;

      try {
        // Create temporary JSON file with resources
        const tempFile = `${filePath}.tmp.json`;
        writeFileSync(
          tempFile,
          JSON.stringify({ resources }, null, 2),
          "utf-8"
        );

        // Call Rust CLI to write the .res file
        await callRustCli("write-res", [tempFile, filePath]);

        // Clean up temp file
        if (existsSync(tempFile)) {
          require("fs").unlinkSync(tempFile);
        }

        return {
          success: true,
          file: filePath,
          resourceCount: resources.length,
        };
      } catch (error) {
        return {
          error: `Failed to write .res file: ${error instanceof Error ? error.message : String(error)}`,
        };
      }
    }

    case "vb6_get_string_table": {
      const filePath = args.file_path as string;
      const blockId = args.block_id as number;

      if (!existsSync(filePath)) {
        return { error: `File not found: ${filePath}` };
      }

      try {
        // Call Rust CLI to parse string table
        const { stdout } = await callRustCli("parse-string-table", [
          filePath,
          blockId.toString(),
        ]);
        const result = JSON.parse(stdout);

        return {
          file: filePath,
          blockId,
          stringCount: result.strings?.length || 0,
          strings: result.strings || [],
        };
      } catch (error) {
        return {
          error: `Failed to parse string table: ${error instanceof Error ? error.message : String(error)}`,
        };
      }
    }

    default:
      return { error: `Unknown tool: ${name}` };
  }
}

/**
 * Main entry point
 */
async function main(): Promise<void> {
  const server = new Server(
    {
      name: "vb6-mcp-server",
      version: "0.1.0",
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Handle list tools request
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return { tools };
  });

  // Handle tool calls
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      const result = await handleToolCall(name, args as Record<string, unknown>);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ error: errorMessage }, null, 2),
          },
        ],
        isError: true,
      };
    }
  });

  // Start the server
  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("VB6 MCP Server started");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
