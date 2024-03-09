<#
.SYNOPSIS
This script calculates the sizes of folders starting from a specified entry point and generates an HTML report including a summary of the total size and detailed information for each folder.

.DESCRIPTION
The script consists of two main functions: `DetermineFolderSizes`, which recursively calculates folder sizes up to a specified depth, and `CreateReport`, which generates an HTML report based on the gathered data. The report includes a company logo and a summary of the total size at the root folder level.

.AUTHOR
Christian Casutt

.COMPANY
Solvia GmbH

.DATE
2024-03-09

.NOTES
This script is designed for internal use at Solvia GmbH and tailored to specific reporting requirements. It demonstrates advanced PowerShell scripting techniques, including recursion, custom object creation, and HTML report generation.

.EXAMPLE
# Example usage:
$folderData = DetermineFolderSizes -FolderEntryPoint "D:\CompanyData\Refolio\Mitarbeitende" -Depth 2
CreateReport -FolderData $folderData
This example calculates the sizes of folders starting from "D:\CompanyData\Refolio\Mitarbeitende" up to two levels deep and then generates an HTML report.

#>

function DetermineFolderSizes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $FolderEntryPoint,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $Depth
    )

    # This internal function traverses the directories up to the specified depth and calculates their sizes.
    function Get-SizeRecursively {
        param (
            [string]
            $CurrentPath,

            [int]
            $CurrentDepth
        )

        # Guard clause to ensure we don't exceed the specified depth.
        if ($CurrentDepth -gt $Depth) {
            return
        }

        try {
            # Get directories at the current path.
            $directories = Get-ChildItem -Path $CurrentPath -Directory -ErrorAction Stop

            foreach ($directory in $directories) {
                $dirSize = Get-ChildItem -Path $directory.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue

                $folderDetails = [PSCustomObject]@{
                    "Name"   = $directory.Name
                    "Path"   = $directory.FullName
                    "SizeGB" = [Math]::Round(($dirSize.Sum / 1GB), 2)
                    "Depth"  = $CurrentDepth
                }

                # Output the current folder's details
                $folderDetails

                # Recursive call to process subdirectories, increasing the depth.
                Get-SizeRecursively -CurrentPath $directory.FullName -CurrentDepth ($CurrentDepth + 1)
            }
        }
        catch {
            Write-Warning "An error occurred processing ${CurrentPath}: $_"
        }
    }

    # Start the recursive directory size calculation.
    Get-SizeRecursively -CurrentPath $FolderEntryPoint -CurrentDepth 1
}

function CreateReport {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[Object]]$FolderData
    )

    $htmlReportPath = "FolderSizeReport.html"

    $htmlContent = @"
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Folder Size Report</title>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
        <style>
        :root {
            --primary-color: #358997; /* A vibrant, futuristic blue */
            --text-color: #333; /* Darker text for better readability */
            --bg-color: #f4f7fa; /* A lighter, modern background */
            --accent-color: #ffc107; /* A bright accent color for important highlights */
            --font-family: 'Roboto', sans-serif; /* A modern, clean font */
        }
        body {
            font-family: var(--font-family);
            color: var(--text-color);
            background-color: var(--bg-color);
            padding: 20px;
            margin: 0;
        }
        .container {
            max-width: 96vw;
            margin: 20px auto;
            background: #fff;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            padding: 20px;
            border-left: 4px solid var(--primary-color); /* Add a touch of color */
        }

        .logo { 
            width: 180px; /* Adjusted for modern design */
            height: auto; 
            margin-bottom: 20px; 
        }

        h1 {
            color: var(--primary-color);
            font-size: 24px; /* Slightly larger for emphasis */
            margin-bottom: 20px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            text-align: left;
            padding: 12px;
            border-bottom: 2px solid #eee; /* Thicker for modern look */
        }
        th {
            background-color: var(--primary-color);
            color: white;
            font-weight: 500;
            font-size: 14px; /* Slightly larger for emphasis */
        }
        td {
            font-weight: 400;
            font-size: 13px; /* Reduced font size for table content */
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            font-size: 14px;
            color: var(--text-color);
        }
    </style>
</head>
<body>
    <div class="container">
        <img src="assets/solvia.svg" alt="Logo" class="logo">
        <h1>Folder Size Report</h1>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Path</th>
                    <th>Size (GB)</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($folder in $FolderData) {
        $htmlContent += @"
                <tr>
                    <td>$($folder.Name)</td>
                    <td>$($folder.Path)</td>
                    <td>$($folder.SizeGB)</td>
                </tr>
"@
    }

    $htmlContent += @"
            </tbody>
        </table>
        <div class="footer">
            <p>Report generated on $(Get-Date)<p>
            <p><a href="https://www.solvia.ch">Solvia GmbH</a><p>
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File $htmlReportPath -Encoding UTF8
    Write-Host "Report generated at: $htmlReportPath"
}

$folderdata = DetermineFolderSizes -FolderEntryPoint "C:\Solvia" -Depth 3
CreateReport -FolderData $folderdata
