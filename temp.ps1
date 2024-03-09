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

    $folderDetailsCollection = New-Object System.Collections.ArrayList

    # Calculate the size of files directly in the root folder
    function AddRootFolderFilesSize {
        $rootFilesSize = (Get-ChildItem -Path $FolderEntryPoint -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        $rootSizeMB = [Math]::Round(($rootFilesSize / 1MB), 2)
        
        # Optionally create a summary object for the root folder's files (not including subdirectories)
        $rootFolderFilesSummary = [PSCustomObject]@{
            Name   = "RootFolderFiles"
            Path   = $FolderEntryPoint
            SizeMB = $rootSizeMB
            Depth  = 0
        }

        [void]$folderDetailsCollection.Add($rootFolderFilesSummary)
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

        $items = Get-ChildItem -Path $CurrentPath -Recurse -Depth 0 -Directory -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $itemSize = (Get-ChildItem -Path $item.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            $itemSizeMB = [Math]::Round(($itemSize / 1MB), 2)

            $folderDetails = [PSCustomObject]@{
                Name   = $item.Name
                Path   = $item.FullName
                SizeMB = $itemSizeMB
                Depth  = $CurrentDepth + 1
            }

            [void]$folderDetailsCollection.Add($folderDetails)

            # Recursively call for subdirectories, if the current depth allows
            if ($CurrentDepth -lt $Depth) {
                Get-SizeRecursively -CurrentPath $item.FullName -CurrentDepth ($CurrentDepth + 1)
            }
        }
    }

    AddRootFolderFilesSize
    Get-SizeRecursively -CurrentPath $FolderEntryPoint -CurrentDepth 0

    return $folderDetailsCollection
}

$ret = DetermineFolderSizes -FolderEntryPoint "C:\solvia" -Depth 1
$ret | Format-Table -AutoSize
$ret | Export-Csv -Path "C:\solvia\folderSizes.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"

