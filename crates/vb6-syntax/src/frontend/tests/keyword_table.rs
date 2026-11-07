use super::*;

#[test]
fn table_length_matches_declared_size() {
    assert_eq!(KEYWORD_TABLE.len(), 271);
}

#[test]
fn token_id_equals_array_index_for_every_entry() {
    // The core invariant relied on by `intern_keywords`: the low 16 bits of
    // w0 (== `token()`) are exactly the entry's index. A regeneration bug
    // that desynchronised the two tables would trip this.
    for (i, e) in KEYWORD_TABLE.iter().enumerate() {
        assert_eq!(e.token() as usize, i, "index {i} (name {:?}) has token {:#x}", e.name, e.token());
    }
}

#[test]
fn known_keywords_sit_at_their_vb6_token_ids() {
    assert_eq!(KEYWORD_TABLE[0x13].name, "Call");
    assert_eq!(KEYWORD_TABLE[0x5f].name, "If");
    assert_eq!(KEYWORD_TABLE[0xaf].name, "Sub");
}

#[test]
fn reserved_word_marker_is_w1_bit_4() {
    // `intern_keywords` reads bit 4 of w1 as the reserved-word marker.
    let sub = KEYWORD_TABLE.iter().find(|e| e.name == "Sub").unwrap();
    let access = KEYWORD_TABLE.iter().find(|e| e.name == "Access").unwrap();
    assert_eq!(sub.w1 >> 4 & 1, 1, "Sub is a reserved word");
    assert_eq!(access.w1 >> 4 & 1, 0, "Access is contextual, not reserved");
}
