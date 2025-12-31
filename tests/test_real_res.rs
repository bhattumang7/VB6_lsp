use vb6_lsp::workspace::*;
use std::path::PathBuf;

/// Get the path to a test fixture file
fn fixture_path(filename: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push(filename);
    path
}

#[test]
fn test_parse_game2048_res() {
    // Read the real Game2048.RES file
    let res_path = fixture_path("Game2048.RES");

    println!("\n=== Reading Game2048.RES ===");

    let resources = match read_res_file(res_path.to_str().unwrap()) {
        Ok(r) => r,
        Err(e) => {
            panic!("Failed to read {}: {}", res_path.display(), e);
        }
    };

    println!("Successfully read {} resource entries\n", resources.len());

    // Analyze what's in the file
    let mut type_counts = std::collections::HashMap::new();

    for entry in &resources {
        let type_name = format!("{:?}", entry.resource_type);
        *type_counts.entry(type_name).or_insert(0) += 1;

        // Print detailed info
        let name = match &entry.name {
            ResourceId::Id(id) => format!("ID {}", id),
            ResourceId::Name(name) => format!("\"{}\"", name),
        };

        println!("Resource: {:?}", entry.resource_type);
        println!("  Name: {}", name);
        println!("  Language: 0x{:04X}", entry.language_id);
        println!("  Data size: {} bytes", entry.data.len());

        // If it's a string table, parse and display it
        if entry.resource_type == ResourceType::String {
            if let Some(block_id) = entry.name.as_id() {
                match parse_string_table(&entry.data, block_id) {
                    Ok(strings) => {
                        println!("  String table entries:");
                        for str_entry in strings {
                            println!("    String {}: \"{}\"", str_entry.id, str_entry.value);
                        }
                    }
                    Err(e) => {
                        println!("  Failed to parse string table: {}", e);
                    }
                }
            }
        }

        println!();
    }

    println!("\n=== Resource Type Summary ===");
    for (type_name, count) in type_counts.iter() {
        println!("{}: {} entries", type_name, count);
    }

    // Verify we got some resources
    assert!(!resources.is_empty(), "Should have found resources in Game2048.RES");
}

#[test]
fn test_roundtrip_game2048_res() {
    // Test that we can read and write back the file
    let res_path = fixture_path("Game2048.RES");
    let output_path = fixture_path("Game2048_copy.RES");

    println!("\n=== Testing round-trip for Game2048.RES ===");

    // Read original
    let resources = read_res_file(res_path.to_str().unwrap())
        .expect("Failed to read original file");

    println!("Read {} resources", resources.len());

    // Write copy
    write_res_file(output_path.to_str().unwrap(), &resources)
        .expect("Failed to write copy");

    println!("Wrote copy to {}", output_path.display());

    // Read copy back
    let resources_copy = read_res_file(output_path.to_str().unwrap())
        .expect("Failed to read copy");

    println!("Read back {} resources from copy", resources_copy.len());

    // Compare
    assert_eq!(resources.len(), resources_copy.len(),
        "Resource count should match");

    for (i, (orig, copy)) in resources.iter().zip(resources_copy.iter()).enumerate() {
        assert_eq!(orig.resource_type, copy.resource_type,
            "Resource {} type mismatch", i);
        assert_eq!(orig.name, copy.name,
            "Resource {} name mismatch", i);
        assert_eq!(orig.language_id, copy.language_id,
            "Resource {} language mismatch", i);
        assert_eq!(orig.data.len(), copy.data.len(),
            "Resource {} data length mismatch", i);
        assert_eq!(orig.data, copy.data,
            "Resource {} data content mismatch", i);
    }

    println!("✓ Round-trip successful - files are identical!");

    // Clean up
    let _ = std::fs::remove_file(&output_path);
}

#[test]
fn test_parse_sheep_res() {
    // Read the real sheep.res file
    let res_path = fixture_path("sheep.res");

    println!("\n=== Reading sheep.res ===");

    let resources = match read_res_file(res_path.to_str().unwrap()) {
        Ok(r) => r,
        Err(e) => {
            panic!("Failed to read {}: {}", res_path.display(), e);
        }
    };

    println!("Successfully read {} resource entries\n", resources.len());

    // Analyze what's in the file
    let mut type_counts = std::collections::HashMap::new();

    for entry in &resources {
        let type_name = format!("{:?}", entry.resource_type);
        *type_counts.entry(type_name).or_insert(0) += 1;
    }

    println!("=== Resource Type Summary ===");
    for (type_name, count) in type_counts.iter() {
        println!("{}: {} entries", type_name, count);
    }

    // Verify we got some resources
    assert!(!resources.is_empty(), "Should have found resources in sheep.res");
}

#[test]
fn test_roundtrip_sheep_res() {
    // Test that we can read and write back the file
    let res_path = fixture_path("sheep.res");
    let output_path = fixture_path("sheep_copy.res");

    println!("\n=== Testing round-trip for sheep.res ===");

    // Read original
    let resources = read_res_file(res_path.to_str().unwrap())
        .expect("Failed to read original file");

    println!("Read {} resources", resources.len());

    // Write copy
    write_res_file(output_path.to_str().unwrap(), &resources)
        .expect("Failed to write copy");

    println!("Wrote copy to {}", output_path.display());

    // Read copy back
    let resources_copy = read_res_file(output_path.to_str().unwrap())
        .expect("Failed to read copy");

    println!("Read back {} resources from copy", resources_copy.len());

    // Compare
    assert_eq!(resources.len(), resources_copy.len(),
        "Resource count should match");

    for (i, (orig, copy)) in resources.iter().zip(resources_copy.iter()).enumerate() {
        assert_eq!(orig.resource_type, copy.resource_type,
            "Resource {} type mismatch", i);
        assert_eq!(orig.name, copy.name,
            "Resource {} name mismatch", i);
        assert_eq!(orig.language_id, copy.language_id,
            "Resource {} language mismatch", i);
        assert_eq!(orig.data.len(), copy.data.len(),
            "Resource {} data length mismatch", i);
        assert_eq!(orig.data, copy.data,
            "Resource {} data content mismatch", i);
    }

    println!("✓ Round-trip successful - files are identical!");

    // Clean up
    let _ = std::fs::remove_file(&output_path);
}

#[test]
fn test_compare_both_res_files() {
    // Compare structure of both .res files
    let game2048_path = fixture_path("Game2048.RES");
    let sheep_path = fixture_path("sheep.res");

    println!("\n=== Comparing Game2048.RES and sheep.res ===");

    let game2048_resources = read_res_file(game2048_path.to_str().unwrap())
        .expect("Failed to read Game2048.RES");

    let sheep_resources = read_res_file(sheep_path.to_str().unwrap())
        .expect("Failed to read sheep.res");

    println!("Game2048.RES: {} resources", game2048_resources.len());
    println!("sheep.res: {} resources", sheep_resources.len());

    // Count types in each file
    let mut game2048_types = std::collections::HashMap::new();
    for entry in &game2048_resources {
        let type_name = format!("{:?}", entry.resource_type);
        *game2048_types.entry(type_name).or_insert(0) += 1;
    }

    let mut sheep_types = std::collections::HashMap::new();
    for entry in &sheep_resources {
        let type_name = format!("{:?}", entry.resource_type);
        *sheep_types.entry(type_name).or_insert(0) += 1;
    }

    println!("\nGame2048.RES types:");
    for (type_name, count) in game2048_types.iter() {
        println!("  {}: {}", type_name, count);
    }

    println!("\nsheep.res types:");
    for (type_name, count) in sheep_types.iter() {
        println!("  {}: {}", type_name, count);
    }

    // Both files should have resources
    assert!(!game2048_resources.is_empty());
    assert!(!sheep_resources.is_empty());
}
