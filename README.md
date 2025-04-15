# Simple Real-time Right-Click Menu Executor

Windows allows you to customize what happens when you right-click through the registry. So, why create this?

This program is designed to let you perform specific actions in the right-click menu only when you select a certain number of files. For example, you can have a "Compare Files" option appear only when you select exactly two files, allowing you to use a third-party diff tool. This option won't clutter your regular right-click menu when only one or no files are selected.

## Screenshots

**1 File Selected:**
![Screenshot of right-click menu with one file selected](pic/1.pic)

**2 Files Selected:**
![Screenshot of right-click menu with two files selected](pic/2.pic)

## Usage

Here's how to use this program:

1.  **Installation:** Open Command Prompt as administrator and run:
    ```
    regsvr32 MultiSelectMenu.dll
    ```
2.  **Uninstallation:** Open Command Prompt as administrator and run:
    ```
    regsvr32 /u MultiSelectMenu.dll
    ```
3.  **Filename:** You can change the name of the `.dll` file, but make sure to also change the name of the `.conf` file to match. The program automatically looks for a `.conf` file with the same name as the `.dll`.
4.  **Real-time Configuration:** The program reads the `.conf` file every time you right-click. You don't need to reinstall the `.dll` after making changes to the `.conf` file.
5.  **Configuration File:** Please modify the `.conf` file according to its format to define the actions you want in the right-click menu for different numbers of selected files.

