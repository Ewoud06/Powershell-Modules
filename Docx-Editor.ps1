# Creating documentation

# Path to the original .docx document
$docxPath = Join-Path -Path $scriptDir -ChildPath "C:\Temp\Test.docx"

# Path where the modified document will be saved
$outputFolder = Join-Path -Path $scriptDir -ChildPath "C:\Temp"

# Key-value pairs to replace in the document
$replacements = @{
    "Example" = "Test"
}

# Create a temporary folder for the modified document
$tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName())
mkdir $tempPath

try {
    # Extract the contents of the .docx (ZIP) to the temp folder
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($docxPath, $tempPath)

    # Load the document.xml file which contains the actual text
    $documentXmlPath = [System.IO.Path]::Combine($tempPath, "word", "document.xml")

    # Use XmlDocument for better encoding support
    [System.Xml.XmlDocument]$documentXml = New-Object System.Xml.XmlDocument
    $documentXml.Load($documentXmlPath)

    # Define namespace for WordprocessingML
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
    $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

    # Perform search and replace
    foreach ($node in $documentXml.SelectNodes("//w:t", $namespaceManager)) {
        foreach ($searchText in $replacements.Keys) {
            $replaceText = $replacements[$searchText]
            if ($node.InnerText -match [regex]::Escape($searchText)) {
                $node.InnerText = $node.InnerText -replace [regex]::Escape($searchText), $replaceText
            }
        }
    }

    # Save the modified XML without BOM
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.Xml.XmlTextWriter($documentXmlPath, $utf8WithoutBom)
    $writer.Formatting = "Indented"
    $documentXml.Save($writer)
    $writer.Close()

    # Define path for the new .docx file
    $newDocxPath = [System.IO.Path]::Combine($outputFolder, [System.IO.Path]::GetFileName($docxPath))

    # Ensure the output folder exists
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }

    # Remove old file if it exists
    if (Test-Path $newDocxPath) {
        Remove-Item $newDocxPath
    }

    # Create a new ZIP (.docx) from the modified files
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($newDocxPath, [System.IO.Compression.ZipArchiveMode]::Create)

    Get-ChildItem -Path $tempPath -Recurse | ForEach-Object {
        if (-not $_.PSIsContainer) {
            $relativePath = $_.FullName.Substring($tempPath.Length + 1) -replace '\\','/'
            $entry = $zip.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $fileStream = [System.IO.File]::OpenRead($_.FullName)
            $fileStream.CopyTo($stream)
            $fileStream.Close()
            $stream.Close()
        }
    }
    $zip.Dispose()

} catch {
    Write-Error "An error occurred: $_"
} finally {
    # Clean up temporary folder
    Remove-Item $tempPath -Recurse -Force
}
