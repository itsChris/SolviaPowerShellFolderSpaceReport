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
CreateReport -FolderData $folderData -HtmlReportPath "C:\Reports"
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

    # Create an ArrayList for better performance and correct scoping behavior
    $folderDetailsCollection = New-Object System.Collections.ArrayList

    # Calculate and add root folder summary
    function AddRootFolderSummary {
        $rootDirSize = Get-ChildItem -Path $FolderEntryPoint -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue

        $rootFolderDetails = [PSCustomObject]@{
            "Name"   = (Get-Item -Path $FolderEntryPoint).Name
            "Path"   = $FolderEntryPoint
            "SizeMB" = [Math]::Round(($rootDirSize.Sum / 1MB), 2)
            "Depth"  = 0
        }

        [void]$folderDetailsCollection.Add($rootFolderDetails)
    }

    function Get-SizeRecursively {
        param (
            [string]
            $CurrentPath,
            [int]
            $CurrentDepth
        )

        if ($CurrentDepth -gt $Depth) {
            return
        }

        try {
            $directories = Get-ChildItem -Path $CurrentPath -Directory -ErrorAction Stop

            foreach ($directory in $directories) {
                $dirSize = Get-ChildItem -Path $directory.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
            
                $folderDetails = [PSCustomObject]@{
                    Name   = $directory.Name
                    Path   = $directory.FullName
                    SizeMB = [Math]::Round(($dirSize.Sum / 1MB), 2)
                    Depth  = $CurrentDepth + 1 # Adjusted to reflect correct depth
                }
            
                [void]$folderDetailsCollection.Add($folderDetails)
            
                Get-SizeRecursively -CurrentPath $directory.FullName -CurrentDepth ($CurrentDepth + 1)
            }
            
        }
        catch {
            Write-Warning "An error occurred processing ${CurrentPath}: $_"
        }
    }

    # Add summary for the root folder
    AddRootFolderSummary

    # Start the recursive directory size calculation
    Get-SizeRecursively -CurrentPath $FolderEntryPoint -CurrentDepth 1

    return $folderDetailsCollection
}

function CreateReport {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[Object]]$FolderData, 
        [Parameter(Mandatory = $true)]
        [string]$ReportPath
    )

    # create folder if not exists
    if (-not (Test-Path -Path $ReportPath -PathType Container)) {
        New-Item -Path $ReportPath -ItemType Directory -Force
    }

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
        th:hover {
            cursor: pointer;
        }
        .sort-asc::after {
            content: " ▲";
        }
        .sort-desc::after {
            content: " ▼";
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            font-size: 14px;
            color: var(--text-color);
        }
    </style>

    <script>
    function sortTable(tableId, col) {
        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
        table = document.getElementById(tableId);
        switching = true;
        // Set the sorting direction to ascending:
        dir = "asc";
        // Make a loop that will continue until no switching has been done:
        while (switching) {
            // Start by saying: no switching is done:
            switching = false;
            rows = table.getElementsByTagName("TR");
            // Loop through all table rows (except the first, which contains table headers):
            for (i = 1; i < (rows.length - 1); i++) {
                // Start by saying there should be no switching:
                shouldSwitch = false;
                // Get the two elements you want to compare, one from current row and one from the next:
                x = rows[i].getElementsByTagName("TD")[col];
                y = rows[i + 1].getElementsByTagName("TD")[col];
                // Check if the two rows should switch place:
                if (dir == "asc") {
                    if (!isNaN(parseFloat(x.innerHTML)) && !isNaN(parseFloat(y.innerHTML))) {
                        // Numeric comparison
                        if (parseFloat(x.innerHTML) > parseFloat(y.innerHTML)) {
                            shouldSwitch = true;
                            break;
                        }
                    } else {
                        // String comparison
                        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                } else if (dir == "desc") {
                    if (!isNaN(parseFloat(x.innerHTML)) && !isNaN(parseFloat(y.innerHTML))) {
                        // Numeric comparison
                        if (parseFloat(x.innerHTML) < parseFloat(y.innerHTML)) {
                            shouldSwitch = true;
                            break;
                        }
                    } else {
                        // String comparison
                        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
            }
            if (shouldSwitch) {
                // If a switch has been marked, make the switch and mark that a switch has been done:
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
                // Each time a switch is completed, increase this count by 1:
                switchcount++;      
            } else {
                // If no switching has been done AND the direction is "asc", set the direction to "desc" and run the while loop again.
                if (switchcount == 0 && dir == "asc") {
                    dir = "desc";
                    switching = true;
                }
            }
        }
    }
    </script>


</head>
<body>
    <div class="container">
        <img src="assets/solvia.svg" alt="Logo" class="logo">
        <h1>Folder Size Report</h1>
        <table id="myTable">
        <thead>
            <tr>
                <th onclick="sortTable('myTable', 0)">Name</th>
                <th onclick="sortTable('myTable', 1)">Path</th>
                <th onclick="sortTable('myTable', 2)">Size (MB)</th>
            </tr>
        </thead>
            <tbody>
"@

    foreach ($folder in $FolderData) {
        $htmlContent += @"
                <tr>
                    <td>$($folder.Name)</td>
                    <td>$($folder.Path)</td>
                    <td>$($folder.SizeMB)</td>
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

    # create filename with datetime stamp
    $ReportPathFile = $ReportPath + "\FolderSizeReport_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html"
    $htmlContent | Out-File $ReportPathFile -Encoding UTF8

    # copy logo to report folder
    $logoPath = $ReportPath + "\assets\solvia.svg"
    if (-not (Test-Path -Path $logoPath)) {
        New-Item -Path $ReportPath\assets -ItemType Directory -Force
        Copy-Item -Path ".\assets\solvia.svg" -Destination $ReportPath\assets -Force
    }

    $csvReportPathFile = $ReportPath + "\FolderSizeReport_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
    $FolderData | Export-Csv -NoTypeInformation -Path $csvReportPathFile -Encoding UTF8 -Delimiter ";"

    Write-Host "Report generated at: $ReportPathFile"
    Write-Host "CSV report generated at: $csvReportPathFile"
}
$folderdata = DetermineFolderSizes -FolderEntryPoint C:\windows -Depth 1
CreateReport -FolderData $folderdata -ReportPath C:\Reports
