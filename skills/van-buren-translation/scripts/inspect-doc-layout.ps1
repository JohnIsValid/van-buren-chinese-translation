param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$resolved = (Resolve-Path -LiteralPath $Path).Path
$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($resolved, $false, $true)

    $items = @()
    $index = 0

    foreach ($shape in $document.InlineShapes) {
        $index++
        $items += [PSCustomObject]@{
            index = $index
            kind = "inline"
            anchor_start = $shape.Range.Start
            page = $shape.Range.Information(3)
            width = [math]::Round($shape.Width, 2)
            height = [math]::Round($shape.Height, 2)
            left = $null
            top = $null
            wrap_type = $null
        }
    }

    foreach ($shape in $document.Shapes) {
        $index++
        $items += [PSCustomObject]@{
            index = $index
            kind = "floating"
            anchor_start = $shape.Anchor.Start
            page = $shape.Anchor.Information(3)
            width = [math]::Round($shape.Width, 2)
            height = [math]::Round($shape.Height, 2)
            left = [math]::Round($shape.Left, 2)
            top = [math]::Round($shape.Top, 2)
            wrap_type = $shape.WrapFormat.Type
        }
    }

    [PSCustomObject]@{
        path = $resolved
        pages = $document.ComputeStatistics(2)
        inline_shape_count = $document.InlineShapes.Count
        floating_shape_count = $document.Shapes.Count
        objects = $items
    } | ConvertTo-Json -Depth 5
}
catch {
    [Console]::Error.WriteLine("DOC layout inspection failed: " + $_.Exception.Message)
    exit 1
}
finally {
    if ($document -ne $null) {
        $document.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
