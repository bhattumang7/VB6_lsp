//! Integration test for VBP parsing with real project files

use std::path::Path;
use vb6_lsp::workspace::{VbpFile, ProjectType};

#[test]
fn test_parse_real_vbp_file() {
    let vbp_path = Path::new("c:\\projects\\VB6_lsp\\vb6parse-master\\tests\\data\\ppdm\\ppdm.vbp");

    if !vbp_path.exists() {
        println!("VBP file not found, skipping test");
        return;
    }

    let result = VbpFile::parse(vbp_path);
    assert!(result.is_ok(), "Failed to parse VBP file: {:?}", result.err());

    let vbp = result.unwrap();

    // Verify project type
    assert_eq!(vbp.project_type, ProjectType::Exe);

    // Verify references were parsed
    println!("References found: {}", vbp.references.len());
    assert!(vbp.references.len() > 0, "Should have references");

    // Check for sub-project references
    let subprojects = vbp.get_subproject_references();
    println!("Sub-project references: {}", subprojects.len());

    // Check for compiled references
    let compiled = vbp.get_compiled_references();
    println!("Compiled references: {}", compiled.len());
    assert!(compiled.len() > 0, "Should have compiled type library references");

    // Verify UUID validation worked
    for reference in &compiled {
        if let Some(uuid) = reference.uuid() {
            println!("Found reference: {} (UUID: {})", reference.description(), uuid);
        }
    }

    // Verify objects were parsed
    println!("Objects found: {}", vbp.objects.len());
    assert!(vbp.objects.len() > 0, "Should have ActiveX objects");

    // Verify modules were parsed
    println!("Modules found: {}", vbp.modules.len());
    for module in &vbp.modules {
        println!("  Module: {} -> {}", module.name, module.relative_path.display());
    }

    // Verify forms were parsed
    println!("Forms found: {}", vbp.forms.len());

    // Verify classes were parsed
    println!("Classes found: {}", vbp.classes.len());
    for class in &vbp.classes {
        println!("  Class: {} -> {}", class.name, class.relative_path.display());
    }

    // Verify user controls
    println!("User controls found: {}", vbp.user_controls.len());

    // Verify designers
    println!("Designers found: {}", vbp.designers.len());
    for designer in &vbp.designers {
        println!("  Designer: {} -> {}", designer.name, designer.relative_path.display());
    }

    // Test member lookup
    if let Some(first_module) = vbp.modules.first() {
        let found = vbp.find_member_by_name(&first_module.name);
        assert!(found.is_some(), "Should find module by name");

        let found_member = found.unwrap();
        assert_eq!(found_member.name, first_module.name);
    }

    // Test all_source_files iterator
    let total_files: usize = vbp.all_source_files().count();
    println!("Total source files: {}", total_files);
    assert!(total_files > 0, "Should have at least one source file");

    // Test compilation settings
    println!("Compilation type: {:?}", vbp.compilation.compilation_type);
    println!("Optimization: {:?}", vbp.compilation.optimization_type);
    println!("Bounds check: {}", vbp.compilation.bounds_check);
    println!("Overflow check: {}", vbp.compilation.overflow_check);

    // Test threading settings
    println!("Threading model: {:?}", vbp.threading.threading_model);
    println!("Max threads: {}", vbp.threading.max_threads);

    // Test version info
    println!("Version: {}.{}.{}",
        vbp.version_info.major,
        vbp.version_info.minor,
        vbp.version_info.revision
    );

    // Test compatibility settings
    println!("Compatibility mode: {:?}", vbp.compatibility.mode);

    // Test custom sections
    if !vbp.custom_sections.is_empty() {
        println!("Custom sections found:");
        for (name, properties) in &vbp.custom_sections {
            println!("  [{}] with {} properties", name, properties.len());
        }
    }

    println!("\n✓ VBP parsing integration test passed!");
}

#[test]
fn test_vbp_workspace_integration() {
    use vb6_lsp::workspace::Vb6Project;

    let vbp_path = Path::new("c:\\projects\\VB6_lsp\\vb6parse-master\\tests\\data\\ppdm\\ppdm.vbp");

    if !vbp_path.exists() {
        println!("VBP file not found, skipping test");
        return;
    }

    let result = Vb6Project::from_vbp(vbp_path);
    assert!(result.is_ok(), "Failed to create project: {:?}", result.err());

    let project = result.unwrap();

    // Test project name
    println!("Project name: {}", project.name());

    // Test VBP path
    assert_eq!(project.vbp_path(), vbp_path);

    // Test root directory
    println!("Root directory: {}", project.root_dir().display());

    // Test source files iterator
    let file_count = project.source_files().count();
    println!("Source files in project: {}", file_count);
    assert!(file_count > 0, "Project should have source files");

    // Test member lookup by name
    for member in project.source_files().take(3) {
        let found = project.get_member_by_name(&member.name);
        assert!(found.is_some(), "Should find member: {}", member.name);

        // Test case-insensitive lookup
        let found_lower = project.get_member_by_name(&member.name.to_lowercase());
        assert!(found_lower.is_some(), "Should find member case-insensitively");
    }

    // Test statistics
    let stats = project.stats();
    println!("Project statistics:");
    println!("  Modules: {}", stats.module_count);
    println!("  Classes: {}", stats.class_count);
    println!("  Forms: {}", stats.form_count);
    println!("  User controls: {}", stats.user_control_count);

    println!("\n✓ VBP workspace integration test passed!");
}

#[test]
fn test_workspace_manager() {
    use vb6_lsp::workspace::WorkspaceManager;
    use std::path::PathBuf;

    let mut manager = WorkspaceManager::new();

    // Test with the vb6parse-master directory
    let root = PathBuf::from("c:\\projects\\VB6_lsp\\vb6parse-master\\tests\\data");

    if !root.exists() {
        println!("Test data directory not found, skipping test");
        return;
    }

    // Add root and discover VBP files
    let discovered = manager.add_root(root.clone());
    println!("Discovered {} VBP files", discovered.len());

    if discovered.is_empty() {
        println!("No VBP files found in test data");
        return;
    }

    // Test workspace statistics
    let stats = manager.stats();
    println!("Workspace statistics:");
    println!("  Roots: {}", stats.root_count);
    println!("  Projects: {}", stats.project_count);
    println!("  Total source files: {}", stats.total_source_files);

    assert_eq!(stats.root_count, 1);
    assert!(stats.project_count > 0, "Should have loaded at least one project");

    // Test project iteration
    for project in manager.projects() {
        println!("Project: {} ({} files)",
            project.name(),
            project.source_files().count()
        );
    }

    println!("\n✓ Workspace manager integration test passed!");
}
