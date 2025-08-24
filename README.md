# Uphide
When Windows 10 was first released, Microsoft provided a tool called wushowhide to allow users to show or hide specific Windows updates—especially useful when certain updates caused problems. However, this tool has become hard to find and is based on deprecated technology.

Uphide is a small, standalone Windows utility that brings back this functionality. It allows you to easily hide unwanted updates directly through the Windows Update API.

## Features
* Lists all pending (uninstalled and unhidden) updates.
* Lets you select updates to hide by number.
* Fully standalone and lightweight executable, coming in at just under 50 KB when compiled with MSVC. It has no runtime dependencies other than the Windows Update Service.
* Based on native Windows Update COM interfaces.

## Usage
1. Run the program as administrator.
2. Wait for the update list to populate.
3. Type the numbers of the updates to hide, separated by spaces.
4. Press Enter.

If you leave the input blank, the program exits without hiding anything.

## Requirements
* Windows 10 or newer.
* Administrative privileges.
    * Note: by default when built with CMake using MSVC, a manifest will be linked that automatically shows a UAC prompt on launch.

## Download
[Download Uphide](https://quinbox.xyz/files/uphide.exe)

## License
This project is licensed under the [MIT License](LICENSE.md).
