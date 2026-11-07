    use super::*;

    fn fold(src: &str) -> Vec<(u32, u32)> {
        folding_ranges_for_source(src.as_bytes())
            .into_iter()
            .map(|r| (r.start_line, r.end_line))
            .collect()
    }

    #[test]
    fn folds_sub() {
        let src = "Private Sub Foo()\n    Dim x As Long\nEnd Sub\n";
        assert_eq!(fold(src), vec![(0, 2)]);
    }

    #[test]
    fn folds_function() {
        let src = "Public Function Add(a As Long) As Long\n    Add = a\nEnd Function\n";
        assert_eq!(fold(src), vec![(0, 2)]);
    }

    #[test]
    fn folds_block_if() {
        let src = "Sub Foo()\n    If x > 0 Then\n        y = 1\n    End If\nEnd Sub\n";
        let ranges = fold(src);
        assert!(ranges.contains(&(0, 4)), "proc range missing: {:?}", ranges);
        assert!(ranges.contains(&(1, 3)), "if range missing: {:?}", ranges);
    }

    #[test]
    fn single_line_if_not_folded() {
        let src = "Sub Foo()\n    If x > 0 Then y = 1\nEnd Sub\n";
        let ranges = fold(src);
        // Only the Sub should fold, not the single-line If
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0], (0, 2));
    }

    #[test]
    fn folds_for_loop() {
        let src = "Sub Foo()\n    For i = 1 To 10\n        x = i\n    Next i\nEnd Sub\n";
        let ranges = fold(src);
        assert!(ranges.contains(&(1, 3)), "for range missing: {:?}", ranges);
    }

    #[test]
    fn folds_nested_procs() {
        let src = "Sub A()\nEnd Sub\nSub B()\nEnd Sub\n";
        assert_eq!(fold(src), vec![(0, 1), (2, 3)]);
    }

    #[test]
    fn closer_keywords_must_end_at_a_word_boundary() {
        // `NEXTVAR`, `LOOPS`, `WENDY` and `END SUBX` close nothing.
        let src = "For i = 1 To 3\nNEXTVAR = 1\nNext\nDo\nLOOPS = 2\nLoop Until x\nWhile y\nWENDY = 3\nWend\nSub S()\nEND SUBX\nEnd Sub\n";
        assert_eq!(fold(src), vec![(0, 2), (3, 5), (6, 8), (9, 11)]);
    }

    #[test]
    fn every_block_kind_folds() {
        let src = "Type T\nx As Long\nEnd Type\nEnum E\nA\nEnd Enum\nProperty Get P()\nEnd Property\nWith o\nx\nEnd With\nSelect Case 1\nCase 1\nEnd Select\n";
        assert_eq!(fold(src), vec![(0, 2), (3, 5), (6, 7), (8, 10), (11, 13)]);
    }

    #[test]
    fn openers_need_a_word_boundary_and_modifiers_are_stripped() {
        // A bare `Sub`/`Function`/`Do` opens a block; the others need a
        // following space.
        let src = "Public Static Sub\nEnd Sub\nDo\nLoop\nDoIt\nFunction\nEnd Function\nSubway = 1\n";
        assert_eq!(fold(src), vec![(0, 1), (2, 3), (5, 6)]);
    }

    #[test]
    fn a_lone_carriage_return_ends_a_line() {
        let src = "Sub A()\r    If x Then\r        y = 1\r    End If\rEnd Sub\r";
        assert_eq!(fold(src), vec![(1, 3), (0, 4)]);
        // The same lines, ended three different ways.
        let mixed = "Sub A()\r\n    If x Then\n        y = 1\r    End If\r\nEnd Sub\n";
        assert_eq!(fold(mixed), fold("Sub A()\n    If x Then\n        y = 1\n    End If\nEnd Sub\n"));
    }
