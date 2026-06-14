# SonarQube analysis (vb6_lsp)

This project is analysed as a **second project** on the SonarQube container that
is owned and started by the sibling **`vba6_rs`** project. There is no separate
server here — both projects publish to the same local instance.

| | |
|---|---|
| Server | `http://localhost:9000` (container `vba6_sonarqube`, SonarQube Community) |
| Project key | `vb6_lsp` |
| Dashboard | http://localhost:9000/dashboard?id=vb6_lsp |
| Container / token owner | `..\vba6_rs` (`docker-compose.sonar.yml`, `.sonar-token`) |

## One-time: start the shared server

Run the setup script in the **vba6_rs** project once (it starts the container and
mints a global analysis token valid for every project on the server):

```powershell
..\vba6_rs\scripts\start-sonar.ps1
```

## Run a scan

From this project root:

```powershell
.\scripts\sonar-scan.ps1            # tests + coverage + clippy, then submit
.\scripts\sonar-scan.ps1 -SkipCoverage   # reuse existing reports, just resubmit
```

The script:

1. runs `cargo llvm-cov` over the default workspace members → `target/coverage/lcov.info`
   (the stale `vb6-core` crate is excluded; it is not in `default-members`);
2. runs `cargo clippy --message-format=json` → `target/clippy-report.json`
   (read by the Sonar Rust/Clippy sensor — cargo is not present in the scanner container);
3. runs `sonarsource/sonar-scanner-cli` via Docker against `host.docker.internal:9000`;
4. polls the analysis task and prints the Quality Gate result.

The analysis token is resolved in order: `-Token` arg → local `.sonar-token` →
`$env:SONAR_TOKEN` → the shared `..\vba6_rs\.sonar-token`. Config lives in
`sonar-project.properties`; the host URL and token are passed at invocation time,
never stored in that file.
