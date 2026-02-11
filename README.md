# Uphide

When Windows 10 was first released, Microsoft provided a tool called wushowhide to allow users to show or hide specific Windows updates—especially useful when certain updates caused problems. However, this tool has become hard to find and is based on deprecated technology. Hence, uphide was born! Uphide is a small, standalone Windows utility that brings back this functionality. It allows you to easily hide unwanted updates directly through the Windows Update API.

## Features

* Lists all pending (uninstalled and unhidden) updates.
* Lets you select updates to hide by number.
* Fully standalone and lightweight, with its only runtime dependency being the Windows Update Service.
* Written in memory-safe Rust.

## Usage

1. Run the program and answer yes to the UAC prompt.
2. Wait for the update list to populate.
3. Type the numbers of the updates to hide, separated by spaces.
4. Press Enter.

If you leave the input blank, the program exits without hiding anything.

## Download

[Download Uphide](https://quinbox.xyz/files/uphide.exe)

## Building

You'll need Rustup and Cargo installed. Once you do, simply running

```batch
cargo build --release
```

will produce a fully functional uphide binary.

## License

This project is licensed under the [MIT License](LICENSE).
