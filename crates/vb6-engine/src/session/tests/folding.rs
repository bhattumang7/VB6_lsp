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
    fn a_lone_carriage_return_ends_a_line() {
        let src = "Sub A()\r    If x Then\r        y = 1\r    End If\rEnd Sub\r";
        assert_eq!(fold(src), vec![(1, 3), (0, 4)]);
        // The same lines, ended three different ways.
        let mixed = "Sub A()\r\n    If x Then\n        y = 1\r    End If\r\nEnd Sub\n";
        assert_eq!(fold(mixed), fold("Sub A()\n    If x Then\n        y = 1\n    End If\nEnd Sub\n"));
    }
