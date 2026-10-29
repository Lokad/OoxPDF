# Office/EMF capture belongs to validation tooling, never the shipped library.
function Test-OfficeImageReference([string] $Directory, [string] $InputSha256) {
    $path=Join-Path $Directory 'reference-image-metadata.json'
    $metadata=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    if ($metadata.Schema -ne 1 -or $metadata.ReferenceKind -ne 'OfficeImage' -or
        $metadata.InputSha256 -ne $InputSha256 -or !$metadata.OfficeBuild -or !$metadata.Capture.Method -or
        (Get-FileHash (Join-Path $Directory 'reference.pdf')).Hash -ne $metadata.SecondaryPdfSha256) {
        throw 'Office image reference provenance differs'
    }
    $pages=@($metadata.Capture.Pages)
    if (!$pages.Count -or @(Get-ChildItem -LiteralPath $Directory -File -Filter 'page-*.png').Count -ne $pages.Count) {
        throw 'Office image reference page count differs'
    }
    for ($index=0; $index -lt $pages.Count; $index++) {
        $page=$pages[$index]
        if ($page.Page -ne $index+1 -or $page.File -cne ('page-{0:000}.png' -f ($index+1)) -or
            $page.Width -le 0 -or $page.Height -le 0 -or $page.WidthPoints -le 0 -or $page.HeightPoints -le 0) {
            throw 'Office image page identity or frame differs'
        }
        $png=Join-Path $Directory $page.File
        if ((Get-FileHash -LiteralPath $png).Hash -ne $page.Sha256) { throw 'Office image page hash differs' }
        $stream=[IO.File]::OpenRead($png)
        try { $header=[byte[]]::new(24);$stream.ReadExactly($header,0,24) } finally { $stream.Dispose() }
        if ([Convert]::ToHexString($header[0..7]) -cne '89504E470D0A1A0A' -or
            [Text.Encoding]::ASCII.GetString($header,12,4) -cne 'IHDR') { throw 'Office image PNG header differs' }
        $width=$header[16]*16777216L+$header[17]*65536L+$header[18]*256L+$header[19]
        $height=$header[20]*16777216L+$header[21]*65536L+$header[22]*256L+$header[23]
        if ($width -ne $page.Width -or $height -ne $page.Height) { throw 'Office image PNG dimensions differ' }
        if ($page.Emf -and ($page.Emf -cne ('page-{0:000}.emf' -f ($index+1)) -or
            (Get-FileHash (Join-Path $Directory $page.Emf)).Hash -ne $page.EmfSha256)) { throw 'Office EMF identity differs' }
    }
    return $metadata
}

function Set-OfficeWordReferenceView($Document, [hashtable] $Settings) {
    $view = $Document.ActiveWindow.View
    $inherited = [ordered]@{MarkupMode=[int]$view.MarkupMode;RevisionsMode=[int]$view.RevisionsMode}
    try {
        if ($Settings.RejectAll) { $Document.Revisions.RejectAll() }
        $view.Type = 3 # wdPrintView
        $Document.ShowRevisions = [bool]$Settings.ShowRevisions
        $view.RevisionsView = [int]$Settings.RevisionsView
        $view.ShowComments = [bool]$Settings.ShowComments
        $view.ShowFormatChanges = [bool]$Settings.ShowFormatChanges
        $view.ShowInsertionsAndDeletions = [bool]$Settings.ShowInsertionsAndDeletions
        # Word can reset MarkupMode while changing category visibility.
        $view.MarkupMode = [int]$Settings.MarkupMode
        $view.RevisionsMode = [int]$Settings.RevisionsMode
        $actual = [ordered]@{ShowRevisions=[bool]$Document.ShowRevisions;RevisionsView=[int]$view.RevisionsView;
            MarkupMode=[int]$view.MarkupMode;RevisionsMode=[int]$view.RevisionsMode;ShowComments=[bool]$view.ShowComments;
            ShowFormatChanges=[bool]$view.ShowFormatChanges;ShowInsertionsAndDeletions=[bool]$view.ShowInsertionsAndDeletions}
        foreach ($key in $actual.Keys) {
            if (!$Settings.ContainsKey($key) -or $actual[$key] -ne $Settings[$key]) { throw ('Word did not apply requested view setting: '+$key) }
        }
        if ($view.Type -ne 3) { throw 'Word did not enter Print Layout' }
        [ordered]@{Actual=$actual;Inherited=$inherited;PrintLayout=$true;RejectAll=[bool]$Settings.RejectAll}
    }
    finally { Release-ComObject $view }
}

function Export-OfficeSlideImages($Presentation, [string] $Directory, [int] $Dpi) {
    $widthPoints = [double]$Presentation.PageSetup.SlideWidth
    $heightPoints = [double]$Presentation.PageSetup.SlideHeight
    $width = [int][Math]::Round($widthPoints * $Dpi / 72, [MidpointRounding]::AwayFromZero)
    $height = [int][Math]::Round($heightPoints * $Dpi / 72, [MidpointRounding]::AwayFromZero)
    $rows = @()
    for ($index=1; $index -le $Presentation.Slides.Count; $index++) {
        $slide = $Presentation.Slides.Item($index)
        try {
            $name = 'page-{0:000}.png' -f $index
            $path = Join-Path $Directory $name
            [void]$slide.Export($path, 'PNG', $width, $height)
            $rows += [ordered]@{Page=$index;File=$name;Sha256=(Get-FileHash $path).Hash;
                Width=$width;Height=$height;WidthPoints=$widthPoints;HeightPoints=$heightPoints}
        }
        finally { Release-ComObject $slide }
    }
    [ordered]@{Method='PowerPoint.Slide.Export(PNG)';View='All source slides, authored static state; no animation playback';Pages=$rows}
}

function Export-OfficeWordImages($Document, [string] $Directory, [int] $Dpi) {
    Add-Type -AssemblyName System.Drawing.Common
    [void]$Document.Repaginate()
    $expected = [int]$Document.ComputeStatistics(2) # wdStatisticPages
    $window = $Document.ActiveWindow
    $pane = $window.Panes.Item(1)
    $pages = $pane.Pages
    try {
        if ($pages.Count -ne $expected -or !$expected) { throw 'Word page-image collection is incomplete after repagination' }
        $rows = @()
        for ($index=1; $index -le $expected; $index++) {
            $page = $pages.Item($index)
            $stream=$null; $metafile=$null; $bitmap=$null; $graphics=$null
            try {
                $widthPoints=[double]$page.Width; $heightPoints=[double]$page.Height
                $width=[int][Math]::Round($widthPoints*$Dpi/72,[MidpointRounding]::AwayFromZero)
                $height=[int][Math]::Round($heightPoints*$Dpi/72,[MidpointRounding]::AwayFromZero)
                [byte[]]$bits=$page.EnhMetaFileBits
                $emfName='page-{0:000}.emf' -f $index
                $emfPath=Join-Path $Directory $emfName
                [IO.File]::WriteAllBytes($emfPath,$bits)
                $stream=[IO.MemoryStream]::new($bits,$false)
                $metafile=[Drawing.Imaging.Metafile]::new($stream)
                $header=$metafile.GetMetafileHeader()
                $bitmap=[Drawing.Bitmap]::new($width,$height)
                $bitmap.SetResolution($Dpi,$Dpi)
                $graphics=[Drawing.Graphics]::FromImage($bitmap)
                $graphics.Clear([Drawing.Color]::White)
                # Replay the complete page frame; no crop, registration or blur.
                $graphics.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,$width,$height))
                $name='page-{0:000}.png' -f $index
                $path=Join-Path $Directory $name
                $bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)
                $rows += [ordered]@{Page=$index;File=$name;Sha256=(Get-FileHash $path).Hash;
                    Width=$width;Height=$height;WidthPoints=$widthPoints;HeightPoints=$heightPoints;
                    Emf=$emfName;EmfSha256=(Get-FileHash $emfPath).Hash;
                    EmfBounds=@($header.Bounds.X,$header.Bounds.Y,$header.Bounds.Width,$header.Bounds.Height);
                    EmfDpi=@($header.DpiX,$header.DpiY)}
            }
            finally {
                foreach ($resource in @($graphics,$bitmap,$metafile,$stream)) { if ($null -ne $resource) { $resource.Dispose() } }
                Release-ComObject $page
            }
        }
        [ordered]@{Method='Word.Pane.Pages.Page.EnhMetaFileBits/GDI+';View='Print Layout';Pages=$rows;
            RasterizerAssembly=[Drawing.Bitmap].Assembly.FullName;ExpectedPages=$expected}
    }
    finally {
        Release-ComObject $pages; Release-ComObject $pane; Release-ComObject $window
    }
}
