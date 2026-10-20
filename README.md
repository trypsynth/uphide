# Uphide

Hide individual Windows updates, or restore updates you hid earlier. Useful when a driver update causes problems and keeps coming back.

Uphide uses the Windows Update API. It does not download, install, or uninstall updates. Restoring an update makes it available again. Windows may refuse to hide mandatory updates, and a newer revision can appear as a new update.

## Download

Download the standalone x64 or ARM64 executable from the [latest release](https://github.com/trypsynth/uphide/releases/latest). Windows 10 or later is required.

## Usage

1. Run Uphide and accept the administrator prompt.
2. Choose Hide updates or Restore hidden updates.
3. Wait for the search to finish.
4. Enter the update numbers, separated by spaces.
5. Review the selection and confirm.

Leave the selection blank to exit. Invalid numbers reject the whole selection so you can correct it. The result shows which updates succeeded or failed.

You can also use `uphide --hide` or `uphide --restore` to skip the menu. `uphide --help` shows usage, and `uphide --version` shows the version.

Windows Update services must be available. The search uses the update source configured on your computer, including an organization's update server if one is configured.

## Building

Install Rust 1.85.0 or later and the Visual Studio C++ Build Tools. Use a Windows MSVC Rust toolchain.

```text
cargo build --release --locked
```

The executable is in `target/release/uphide.exe`. The C runtime is linked statically. `embed-manifest` embeds the administrator manifest without a separate XML file or resource compiler.

## Checks

```text
cargo fmt --check
cargo clippy --all-targets --locked -- -D warnings
cargo test --locked
```

Clippy's all, pedantic, and nursery groups are denied in Cargo.toml. Tests check command and selection handling without changing Windows updates. CI runs the checks and builds on native x64 and ARM64 runners. A separate job tests and builds with Rust 1.85.0.

Pushing a version tag such as `v0.2.0` publishes both executables after every check succeeds.

## License

[MIT](LICENSE).
