#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "tree_sitter/api.h"

const TSLanguage *tree_sitter_vb6(void);

static char *read_all(FILE *f) {
    size_t cap = 1 << 16, len = 0;
    char *buf = malloc(cap);
    int c;
    while ((c = fgetc(f)) != EOF) {
        if (len + 1 >= cap) { cap *= 2; buf = realloc(buf, cap); }
        buf[len++] = (char)c;
    }
    buf[len] = '\0';
    return buf;
}

int main(int argc, char **argv) {
    FILE *f = argc > 1 ? fopen(argv[1], "rb") : stdin;
    if (!f) { fprintf(stderr, "cannot open %s\n", argv[1]); return 2; }
    char *src = read_all(f);
    if (argc > 1) fclose(f);

    TSParser *parser = ts_parser_new();
    ts_parser_set_language(parser, tree_sitter_vb6());
    TSTree *tree = ts_parser_parse_string(parser, NULL, src, (uint32_t)strlen(src));
    TSNode root = ts_tree_root_node(tree);
    char *s = ts_node_string(root);
    printf("%s\n", s);
    free(s);
    ts_tree_delete(tree);
    ts_parser_delete(parser);
    free(src);
    return 0;
}
