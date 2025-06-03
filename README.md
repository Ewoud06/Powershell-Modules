# WordReplacer

A lightweight PowerShell script that allows you to replace specific text inside `.docx` Word documents by manipulating the underlying XML.

## 🚀 Features

- Replace text in Word documents without needing Microsoft Word
- Works on any `.docx` file (since `.docx` is essentially a ZIP archive)
- Keeps document structure intact
- No external dependencies

## 📦 Prerequisites

- Windows PowerShell (tested on Windows 10/11)
- A `.docx` document as input

## 📂 Usage

1. Modify the `$docxPath` variable in the script to point to your `.docx` file.
2. Define your replacements in the `$replacements` hashtable.
3. Run the script. A modified `.docx` will be saved to the output folder.

## 🔧 Example

```powershell
$replacements = @{
    "Hello" = "Hi"
    "World" = "Universe"
}
