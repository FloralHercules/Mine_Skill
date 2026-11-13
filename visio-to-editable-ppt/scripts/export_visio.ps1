<#
Export a Visio page for native PowerPoint reconstruction, modifying a private
copy only. All parameter names and script text are ASCII for Windows PS5.
Reference images are emitted before and after ungrouping for visual QA.
The caller next runs inspect_vsdx.py on flattened.vsdx to build the manifest.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$PageName,
    [Parameter(Mandatory=$true)][string]$WorkDir,
    [ValidateRange(72,600)][int]$RasterDpi=216,
    [switch]$ExportNative3DDiagnostics
)
$ErrorActionPreference = 'Stop'

function Release-Com($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        # Collection access can return the same RCW held in another variable.
        # Release only this acquisition, never all references of a shared RCW.
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value) } catch {}
    }
}

function Find-Page($Document, [string]$WantedName) {
    for ($index=1; $index -le $Document.Pages.Count; $index++) {
        $candidate = $Document.Pages.Item($index)
        if ($candidate.Name -eq $WantedName -or $candidate.NameU -eq $WantedName) { return $candidate }
        Release-Com $candidate
    }
    throw "Page '$WantedName' was not found in the source drawing."
}

function Has-3DEffects($Shape) {
    foreach ($cellName in @('RotationXAngle','RotationYAngle','RotationZAngle','RotationType','Perspective','DistanceFromGround','Depth','BevelTopHeight','BevelBottomHeight')) {
        try {
            if ($Shape.CellExistsU($cellName,0) -ne 0 -and [Math]::Abs([double]$Shape.CellsU($cellName).ResultIU) -gt 0.000001) { return $true }
        } catch {}
    }
    return $false
}

$sourcePath = (Resolve-Path -LiteralPath $Source).ProviderPath
if ([IO.Path]::GetExtension($sourcePath) -ine '.vsdx') { throw 'Source must be a .vsdx file.' }
$workPath = [IO.Path]::GetFullPath($WorkDir)
$null = New-Item -ItemType Directory -Path $workPath -Force
$workingPath = Join-Path $workPath 'working.vsdx'
$flattenedPath = Join-Path $workPath 'flattened.vsdx'
$referencePath = Join-Path $workPath 'reference.png'
$flattenedReferencePath = Join-Path $workPath 'flattened-reference.png'
$svgPath = Join-Path $workPath 'source.svg'
$reportPath = Join-Path $workPath 'export-report.json'
foreach ($reserved in @($workingPath,$flattenedPath,$referencePath,$flattenedReferencePath,$svgPath,$reportPath)) {
    if (Test-Path -LiteralPath $reserved) { throw "Export output already exists: $reserved. Use a fresh work directory." }
}
if ($sourcePath -eq $workingPath -or $sourcePath -eq $flattenedPath) { throw 'The private copy must differ from the source.' }
$assetPath = Join-Path $workPath 'assets'
$null = New-Item -ItemType Directory -Path $assetPath -Force
Copy-Item -LiteralPath $sourcePath -Destination $workingPath
(Get-Item -LiteralPath $workingPath).IsReadOnly = $false
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash

$app = $null; $doc = $null; $page = $null; $frame = $null; $ownsApp = $false
$warnings = [Collections.Generic.List[object]]::new()
$exported = [Collections.Generic.List[object]]::new()
$blockedGroups = [Collections.Generic.HashSet[int]]::new()
$ungroupCount = 0; $pageWidth = 0.0; $pageHeight = 0.0; $frameId = $null
$oldResolution = 0; $oldResolutionWidth = 0.0; $oldResolutionHeight = 0.0; $oldResolutionUnit = 0
$oldSize = 0; $oldSizeWidth = 0.0; $oldSizeHeight = 0.0; $oldSizeUnit = 0
$settingsCaptured = $false
$report = $null
try {
    # Create a dedicated invisible instance. Never attach to a user's Visio app.
    $app = New-Object -ComObject Visio.InvisibleApp
    $ownsApp = $true
    $app.AlertResponse = 7
    $settings = $app.Settings
    $settings.GetRasterExportResolution([ref]$oldResolution,[ref]$oldResolutionWidth,[ref]$oldResolutionHeight,[ref]$oldResolutionUnit)
    $settings.GetRasterExportSize([ref]$oldSize,[ref]$oldSizeWidth,[ref]$oldSizeHeight,[ref]$oldSizeUnit)
    $oldBackground = $settings.RasterExportBackgroundColor
    $oldTransparency = $settings.RasterExportTransparencyColor
    $oldUseTransparency = $settings.RasterExportUseTransparencyColor
    $settingsCaptured = $true
    $settings.SetRasterExportResolution(3,$RasterDpi,$RasterDpi,0)
    $settings.SetRasterExportSize(2,0,0,0)
    $settings.RasterExportBackgroundColor = 16777215
    $settings.RasterExportTransparencyColor = 16777215
    # Opaque reference snapshots allow a direct before/after visual comparison.
    $settings.RasterExportUseTransparencyColor = $false
    $doc = $app.Documents.OpenEx($workingPath,64)
    $page = Find-Page $doc $PageName
    $pageWidth = [double]$page.PageSheet.CellsU('PageWidth').ResultIU
    $pageHeight = [double]$page.PageSheet.CellsU('PageHeight').ResultIU
    $initialTopLevelCount = $page.Shapes.Count
    $page.Export($referencePath)
    Write-Output 'Saved the original-page reference; flattening groups in the private copy.'

    for ($pass=0; $pass -lt 1000; $pass++) {
        $pending = [Collections.Generic.List[int]]::new()
        for ($index=1; $index -le $page.Shapes.Count; $index++) {
            $shape = $page.Shapes.Item($index)
            if ($shape.Type -eq 2 -and -not $blockedGroups.Contains([int]$shape.ID)) { $pending.Add([int]$shape.ID) }
            Release-Com $shape
        }
        if ($pending.Count -eq 0) { break }
        foreach ($shapeId in $pending) {
            $shape = $null
            try {
                $shape = $page.Shapes.ItemFromID($shapeId)
                $null = $shape.Ungroup()
                $ungroupCount++
            } catch {
                [void]$blockedGroups.Add($shapeId)
                $warnings.Add([pscustomobject]@{code='ungroup-failed';source_id=$shapeId;message=$_.Exception.Message})
            } finally { Release-Com $shape }
        }
        if ($pass -eq 999) { throw 'Group flattening exceeded 1000 passes.' }
    }
    $remainingGroups = [Collections.Generic.List[object]]::new()
    for ($index=1; $index -le $page.Shapes.Count; $index++) {
        $shape = $page.Shapes.Item($index)
        if ($shape.Type -eq 2) { $remainingGroups.Add([pscustomobject]@{source_id=[int]$shape.ID;name=[string]$shape.Name}) }
        Release-Com $shape
    }
    $page.Export($flattenedReferencePath)
    $page.Export($svgPath)
    $doc.SaveAs($flattenedPath)
    $flattenedTopLevelCount = $page.Shapes.Count

    # The thin visible frame forces full-page export extents without grouping
    # or moving any original object. Later processing clears only the 3px
    # perimeter (no coordinate shift) and alpha-crops each resulting asset.
    $frame = $page.DrawRectangle(0,0,$pageWidth,$pageHeight)
    $frameId = [int]$frame.ID
    $frame.NameU = 'EditablePptExportExtent'
    $frame.CellsU('LineColor').FormulaU = 'RGB(255,0,255)'
    $frame.CellsU('LineWeight').FormulaU = '0.0001 in'
    $frame.CellsU('LinePattern').FormulaU = '1'
    $frame.CellsU('FillPattern').FormulaU = '0'
    $frame.CellsU('ShdwPattern').FormulaU = '0'
    $frame.SendToBack()
    # White transparency avoids magenta color-key fringes around source assets.
    # The only magenta pixels belong to the removable full-page extent frame.
    $settings.RasterExportUseTransparencyColor = $true
    for ($index=1; $index -le $page.Shapes.Count; $index++) {
        $shape = $null; $selection = $null; $shapeId = $null
        try {
            $shape = $page.Shapes.Item($index)
            $shapeId = [int]$shape.ID
            if ($shapeId -eq $frameId) { continue }
            $foreign = $shape.Type -eq 4
            $diagnostic = (-not $foreign) -and $ExportNative3DDiagnostics -and (Has-3DEffects $shape)
            if (-not $foreign -and -not $diagnostic) { continue }
            $filename = if ($foreign) { "shape-$shapeId.png" } else { "diagnose-$shapeId.png" }
            $asset = Join-Path $assetPath $filename
            $selection = $page.CreateSelection(0,0)
            $selection.Select($shape,2)
            $selection.Select($frame,2)
            $selection.Export($asset)
            $exported.Add([pscustomobject]@{source_id=$shapeId;raster=$foreign;diagnostic=[bool]$diagnostic;asset=$asset})
        } finally {
            Release-Com $selection
            # The extent frame is retained through the entire export loop.
            # A continue still enters finally, so do not release its alias.
            if ($null -ne $shape -and $shapeId -ne $frameId) { Release-Com $shape }
        }
    }
    $frame.Delete()
    Release-Com $frame
    $frame = $null
    # Do not save the temporary extent frame or its transient IDs into the copy.
    $sourceHashAfter = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    if ($sourceHashAfter -ne $sourceHash) { throw 'The source changed during export; compare it before proceeding.' }
    $report = [ordered]@{
        schema_version='visio-export.v1';status=$(if($remainingGroups.Count){'needs-review'}else{'exported'});
        source=$sourcePath;source_sha256=$sourceHash;source_unchanged=$true;work_dir=$workPath;
        page_name=[string]$page.Name;page_id=[int]$page.ID;page_width_in=$pageWidth;page_height_in=$pageHeight;
        working_vsdx=$workingPath;flattened_vsdx=$flattenedPath;reference_png=$referencePath;
        flattened_reference_png=$flattenedReferencePath;source_svg=$svgPath;asset_dir=$assetPath;
        initial_top_level_count=$initialTopLevelCount;flattened_top_level_count=$flattenedTopLevelCount;
        ungrouped_group_count=$ungroupCount;unsupported_groups=@($remainingGroups.ToArray());
        raster_dpi=$RasterDpi;frame_shape_id=$frameId;frame_color='#FF00FF';clear_border_px=3;
        assets=@($exported.ToArray());warnings=@($warnings.ToArray());
        requires_before_after_visual_review=$true
    }
    $json = $report | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText($reportPath,$json,[Text.UTF8Encoding]::new($false))
    Write-Output ("Exported {0} raster/diagnostic assets. Report: {1}" -f $exported.Count,$reportPath)
    if ($remainingGroups.Count) { Write-Warning 'Unresolved groups remain; review them before native assembly.' }
} finally {
    if ($null -ne $frame) { try { $frame.Delete() } catch {} }
    Release-Com $frame
    Release-Com $page
    if ($null -ne $doc) { try { $doc.Saved=$true; $doc.Close() } catch {} }
    Release-Com $doc
    if ($null -ne $app -and $settingsCaptured) {
        try {
            $settings.SetRasterExportResolution($oldResolution,$oldResolutionWidth,$oldResolutionHeight,$oldResolutionUnit)
            $settings.SetRasterExportSize($oldSize,$oldSizeWidth,$oldSizeHeight,$oldSizeUnit)
            $settings.RasterExportBackgroundColor=$oldBackground
            $settings.RasterExportTransparencyColor=$oldTransparency
            $settings.RasterExportUseTransparencyColor=$oldUseTransparency
        } catch {}
    }
    Release-Com $settings
    if ($ownsApp -and $null -ne $app) { try { $app.Quit() } catch {} }
    Release-Com $app
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
