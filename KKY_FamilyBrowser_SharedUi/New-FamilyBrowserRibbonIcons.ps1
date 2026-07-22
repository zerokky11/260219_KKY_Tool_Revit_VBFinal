param(
    [string]$OutputDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundedRectanglePath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$Radius
    )

    $diameter = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function Save-ScaledPng {
    param(
        [System.Drawing.Bitmap]$Source,
        [int]$Size,
        [string]$Path
    )

    $target = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $target.SetResolution(96, 96)
        $graphics = [System.Drawing.Graphics]::FromImage($target)
        try {
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.DrawImage($Source, 0, 0, $Size, $Size)
        }
        finally {
            $graphics.Dispose()
        }
        $target.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $target.Dispose()
    }
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$master = New-Object System.Drawing.Bitmap(256, 256, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
try {
    $master.SetResolution(96, 96)
    $graphics = [System.Drawing.Graphics]::FromImage($master)
    try {
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $tilePath = New-RoundedRectanglePath 10 10 236 236 42
        try {
            $tileBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                (New-Object System.Drawing.RectangleF(10, 10, 236, 236)),
                ([System.Drawing.Color]::FromArgb(255, 18, 39, 78)),
                ([System.Drawing.Color]::FromArgb(255, 8, 19, 42)),
                90.0
            )
            $tilePen = New-Object System.Drawing.Pen(([System.Drawing.Color]::FromArgb(255, 47, 107, 255)), 8)
            try {
                $graphics.FillPath($tileBrush, $tilePath)
                $graphics.DrawPath($tilePen, $tilePath)
            }
            finally {
                $tileBrush.Dispose()
                $tilePen.Dispose()
            }
        }
        finally {
            $tilePath.Dispose()
        }

        $shadowPath = New-RoundedRectanglePath 35 91 188 122 18
        $folderPath = New-RoundedRectanglePath 33 86 188 122 18
        try {
            $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(95, 0, 0, 0))
            $folderBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 247, 250, 255))
            $folderPen = New-Object System.Drawing.Pen(([System.Drawing.Color]::FromArgb(255, 190, 211, 241)), 5)
            try {
                $graphics.FillPath($shadowBrush, $shadowPath)
                $graphics.FillPath($folderBrush, $folderPath)
                $graphics.DrawPath($folderPen, $folderPath)
            }
            finally {
                $shadowBrush.Dispose()
                $folderBrush.Dispose()
                $folderPen.Dispose()
            }
        }
        finally {
            $shadowPath.Dispose()
            $folderPath.Dispose()
        }

        $tabPoints = [System.Drawing.PointF[]]@(
            (New-Object System.Drawing.PointF(45, 93)),
            (New-Object System.Drawing.PointF(45, 72)),
            (New-Object System.Drawing.PointF(106, 72)),
            (New-Object System.Drawing.PointF(127, 93)),
            (New-Object System.Drawing.PointF(208, 93)),
            (New-Object System.Drawing.PointF(208, 110)),
            (New-Object System.Drawing.PointF(45, 110))
        )
        $tabBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 222, 233, 249))
        try {
            $graphics.FillPolygon($tabBrush, $tabPoints)
        }
        finally {
            $tabBrush.Dispose()
        }

        $cardColors = @(
            [System.Drawing.Color]::FromArgb(255, 47, 107, 255),
            [System.Drawing.Color]::FromArgb(255, 47, 183, 255),
            [System.Drawing.Color]::FromArgb(255, 31, 78, 174)
        )
        $cardXs = @(59, 103, 147)
        for ($index = 0; $index -lt $cardXs.Count; $index++) {
            $cardPath = New-RoundedRectanglePath $cardXs[$index] 118 31 61 6
            $cardBrush = New-Object System.Drawing.SolidBrush($cardColors[$index])
            try {
                $graphics.FillPath($cardBrush, $cardPath)
            }
            finally {
                $cardBrush.Dispose()
                $cardPath.Dispose()
            }
        }

        $linePen = New-Object System.Drawing.Pen(([System.Drawing.Color]::FromArgb(220, 255, 255, 255)), 4)
        try {
            foreach ($x in $cardXs) {
                $graphics.DrawLine($linePen, $x + 7, 133, $x + 24, 133)
                $graphics.DrawLine($linePen, $x + 7, 146, $x + 24, 146)
                $graphics.DrawLine($linePen, $x + 7, 159, $x + 20, 159)
            }
        }
        finally {
            $linePen.Dispose()
        }

        $accentPath = New-RoundedRectanglePath 184 127 16 52 5
        $accentBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 166, 31))
        try {
            $graphics.FillPath($accentBrush, $accentPath)
        }
        finally {
            $accentBrush.Dispose()
            $accentPath.Dispose()
        }
    }
    finally {
        $graphics.Dispose()
    }

    Save-ScaledPng $master 16 (Join-Path $OutputDirectory 'family-browser-ribbon-16.png')
    Save-ScaledPng $master 32 (Join-Path $OutputDirectory 'family-browser-ribbon-32.png')
}
finally {
    $master.Dispose()
}

Write-Output "Family Browser ribbon icons generated in $OutputDirectory"
