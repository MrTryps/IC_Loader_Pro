# --- Configuration ---
# The path where the final prompt file will be saved.
$outputFile = ".\GeminiPrompt.txt"

# --- MODIFIED: Add the relative paths to ALL project folders you want to include ---
$projectPaths = @(
    ".",  # The current directory (for IC_Loader_Pro)
    "..\..\BIS_Tools_2025\IC_Rules_2025\IC_Rules_2025",
    "..\..\BIS_Tools_2025\BIS_Tools_DataModels_2025\BIS_Tools_DataModels_2025\BIS_Tools_DataModels_2025",
    "..\..\BIS_Tools_2025\BIS_Tools_2025_C_Core\BIS_Tools_2025_C_Core\BIS_Tools_2025_C_Core"
)

# File extensions to include in the prompt.
$includeExtensions = @("*.cs", "*.xaml", "*.csproj", "*.daml", "*.sln")

# Directories to exclude.
$excludeDirs = @("*bin", "*obj", "*packages*", "*.vs", "*Properties*")

# --- Script Start ---

$preamble = @"
You are an expert C# developer specializing in the ArcGIS Pro SDK. I am providing you with the complete source code for a multi-project Visual Studio solution. Please analyze all the provided files to understand the project's architecture, goals, and current state. Your task is to act as a collaborative partner to help me debug and enhance this application.

Here is the project structure and the full content of each relevant file:
"@

$promptContent = $preamble

# Generate a tree view from the parent of the first project path for a broader view
$treeOutput = (tree (Split-Path -Path (Resolve-Path -Path $projectPaths[0]).Path -Parent) /F /A) -join "`n"
$promptContent += "--- OVERALL SOLUTION STRUCTURE ---`n"
$promptContent += "````n"
$promptContent += $treeOutput
$promptContent += "`n````n`n"

# Loop through each specified project path
foreach ($projectPath in $projectPaths) {
    # Get all relevant files in the current project path
    $allFiles = Get-ChildItem -Path $projectPath -Recurse -Include $includeExtensions -Exclude $excludeDirs

    foreach ($file in $allFiles) {
        # Add a clear delimiter and the file path
        $promptContent += "--- FILE: $($file.FullName.Replace($PWD.Path + '\', '')) ---`n"
        $promptContent += "```csharp`n"
        
        $fileContent = Get-Content -Path $file.FullName -Raw
        $promptContent += $fileContent
        
        $promptContent += "`n````n`n"
    }
}

$promptContent += "--- MY QUESTION ---`n"
$promptContent += "Now that you have the complete context, here is my question: `n`n`n`n"

Set-Content -Path $outputFile -Value $promptContent

Write-Host "Prompt successfully generated at: $outputFile"
Write-Host "You can now open this file, add your question, and paste the entire content into Gemini."