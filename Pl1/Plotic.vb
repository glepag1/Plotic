﻿Public Class Plotic
    Private strGunName As String
    Private strFileName As String
    Private strTitle As String
    Private strInfo As String
    Private strSubText As String

    Private strBackgroundColor As String

    Private bmpImage As Bitmap = New Bitmap(2000, 2000)
    Private bmpSilh As Bitmap = New Bitmap(2000, 2000)
    Private bmpMask As Bitmap = New Bitmap(2000, 2000)
    Private bmpHeatMap As Bitmap = New Bitmap(2000, 2000)
    Private bmpTTK As Bitmap = New Bitmap(2000, 2000)
    Private grhImageGraphic As Graphics = Graphics.FromImage(bmpImage)
    Private grhMaskGraphic As Graphics = Graphics.FromImage(bmpMask)
    Private grhHeatGraphic As Graphics = Graphics.FromImage(bmpHeatMap)
    Private grhTTKGraphic As Graphics = Graphics.FromImage(bmpTTK)

    Private dblRecoilUp As Double
    Private dblRecoilRight As Double
    Private dblRecoilLeft As Double

    'Dropping recoil decrease
    'Private dblRecoilDecrease As Double
    Private dblSpreadMin As Double
    Private dblSpreadInc As Double
    Private dblAdjRecoilH As Double
    Private dblAdjRecoilV As Double
    Private dblAdjSpreadMin As Double
    Private dblAdjSpreadInc As Double
    Private dblGridLineSpace As Double

    Private dblHealthPercent As Double

    Private intBurst As Integer
    Private intBulletsPerBurst As Integer
    Private dblFirstShot As Double
    Private intScale As Integer

    Private dblBulletDrop As Double
    Private dblBulletVelocity As Double
    Private dblMaxDistance As Double
    Private dblTargetRange As Double
    Private dblRateOfFire As Double

    Private dblTTKDamageMax As Double
    Private dblTTKDamageMin As Double
    Private dblTTKRangeMax As Double
    Private dblTTKRangeMin As Double

    'Adding All options to the plotic object

    Private dblHeatMapIntensity As Double

    Private dblBackgroundColorAlpha As Integer
    Private dblBackgroundColorRed As Integer
    Private dblBackgroundColorGreen As Integer
    Private dblBackgroundColorBlue As Integer

    Private blnRenderTitle As Boolean
    Private blnRenderBars As Boolean
    Private blnRenderAdjustment As Boolean
    Private blnRenderHitRates As Boolean
    Private blnRenderTTK As Boolean
    Private blnRenderScaleTarget As Boolean

    Private blnRenderDropInfo As Boolean
    Private blnRenderBulletDrop As Boolean
    Private blnRenderAmmoInfo As Boolean
    Private blnRenderTTKGrid As Boolean
    Private blnRenderDropGrid As Boolean
    Private blnRenderGrid As Boolean

    Private blnScaleRadius As Boolean

    Private blnRenderHeatMap As Boolean
    Private blnRenderHeatTitle As Boolean
    Private blnRenderHeatBars As Boolean
    Private blnRenderHeatAdjust As Boolean





    Public Sub New()

    End Sub
#Region "Functions"
    Public Function bulletHit(ByVal x As Integer, ByVal y As Integer) As Boolean
        Dim colo As Object
        Dim rgbb As Integer
        If x < 0 Or y < 0 Or x > 1999 Or y > 1999 Then
            rgbb = 0
        Else
            colo = Me.Mask.GetPixel(x, y)
            rgbb = Val(colo.R) + Val(colo.G) + Val(colo.B)
        End If
        'Debug.WriteLine("X: " & x & " Y: " & y & " Val: " & rgbb)
        If rgbb > 100 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overloads Function dropInMeters(ByVal presision As Integer) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = -(Me.BulletDrop * (Me.TargetRange ^ 2)) / (2 * (Me.BulletVelocity ^ 2))
        Return Math.Round(bulletDrop, presision)
    End Function
    Public Overloads Function dropInMeters(ByVal presision As Integer, ByVal range As Double) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = -(Me.BulletDrop * (range ^ 2)) / (2 * (dblBulletVelocity ^ 2))
        Return Math.Round(bulletDrop, presision)
    End Function

    Public Overloads Function correctionInMeters(ByVal presision As Integer) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = (Math.Asin((Me.TargetRange * (Me.BulletDrop)) / (Me.BulletVelocity ^ 2)))
        bulletDrop = (bulletDrop * 0.5)
        bulletDrop = Me.TargetRange * (Math.Tan(bulletDrop))
        Return Math.Round(bulletDrop, presision)
    End Function
    Public Overloads Function correctionInMeters(ByVal presision As Integer, ByVal range As Double) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = (Math.Asin((range * (Me.BulletDrop)) / (Me.BulletVelocity ^ 2)))
        bulletDrop = (bulletDrop * 0.5)
        bulletDrop = range * (Math.Tan(bulletDrop))
        Return Math.Round(bulletDrop, presision)
    End Function

    Public Overloads Function correctionAngle(ByVal presision As Integer) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = 0.5 * (Math.Asin((Me.TargetRange * Me.BulletDrop) / (Me.BulletVelocity ^ 2)))
        bulletDrop = bulletDrop * (180 / Math.PI)
        Return Math.Round(bulletDrop, presision)
    End Function
    Public Overloads Function correctionAngle(ByVal presision As Integer, ByVal range As Double) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = 0.5 * (Math.Asin((range * Me.BulletDrop) / (Me.BulletVelocity ^ 2)))
        bulletDrop = bulletDrop * (180 / Math.PI)
        Return Math.Round(bulletDrop, presision)
    End Function

    Public Function bulletHeight(ByVal presision As Integer, ByVal range As Double) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = (range * Math.Tan(correctionAngle(5))) - (Me.BulletDrop * (range ^ 2) / (2 * (630 * Math.Cos(correctionAngle(5))) ^ 2))
        bulletDrop = bulletDrop * (180 / Math.PI)
        Return Math.Round(bulletDrop, presision)
    End Function



    Public Overloads Function timeOfFlight(ByVal presision As Integer) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = Me.TargetRange / (Me.BulletVelocity * Math.Cos(((Math.Asin((Me.TargetRange * Me.BulletDrop) / (Me.BulletVelocity ^ 2))))))
        Return Math.Round(bulletDrop, presision)
    End Function
    Public Overloads Function timeOfFlight(ByVal presision As Integer, ByVal range As Double) As Double
        Dim bulletDrop As Double = 0
        bulletDrop = range / (Me.BulletVelocity * Math.Cos(((Math.Asin((range * Me.BulletDrop) / (Me.BulletVelocity ^ 2))))))
        Return Math.Round(bulletDrop, presision)
    End Function

    Public Overloads Function dropInPixels() As Integer
        Return Math.Round((Math.Atan(Me.dropInMeters(5) / Me.TargetRange) * (180 / Math.PI)) * Me.Scale, 0)
    End Function
    Public Overloads Function dropInPixels(ByVal range As Double, ByVal scale As Integer) As Integer
        Return Math.Round((Math.Atan(Me.dropInMeters(5, range) / range) * (180 / Math.PI)) * scale, 0)
    End Function

    Public Function correctionInPixels() As Integer
        Return Math.Round((Math.Atan(Me.correctionInMeters(5) / Me.TargetRange) * (180 / Math.PI)) * Me.Scale, 0)
    End Function
    Public Function correctionInPixels(ByVal range As Double, ByVal scale As Integer) As Integer
        Return Math.Round((Math.Atan(Me.correctionInMeters(5, range) / range) * (180 / Math.PI)) * scale, 0)
    End Function
#End Region
#Region "Properties"

    Property HeatMapIntensity() As Double
        Get
            Return dblHeatMapIntensity
        End Get
        Set(ByVal Value As Double)
            dblHeatMapIntensity = Value
        End Set
    End Property

    Property BackgroundColorAlpha() As Integer
        Get
            Return dblBackgroundColorAlpha
        End Get
        Set(ByVal Value As Integer)
            dblBackgroundColorAlpha = Value
        End Set
    End Property
    Property BackgroundColorRed() As Integer
        Get
            Return dblBackgroundColorRed
        End Get
        Set(ByVal Value As Integer)
            dblBackgroundColorRed = Value
        End Set
    End Property
    Property BackgroundColorGreen() As Integer
        Get
            Return dblBackgroundColorGreen
        End Get
        Set(ByVal Value As Integer)
            dblBackgroundColorGreen = Value
        End Set
    End Property
    Property BackgroundColorBlue() As Integer
        Get
            Return dblBackgroundColorBlue
        End Get
        Set(ByVal Value As Integer)
            dblBackgroundColorBlue = Value
        End Set
    End Property

    Property RenderTitle() As Boolean
        Get
            Return blnRenderTitle
        End Get
        Set(ByVal Value As Boolean)
            blnRenderTitle = Value
        End Set
    End Property
    Property RenderBars() As Boolean
        Get
            Return blnRenderBars
        End Get
        Set(ByVal Value As Boolean)
            blnRenderBars = Value
        End Set
    End Property
    Property RenderAdjustment() As Boolean
        Get
            Return blnRenderAdjustment
        End Get
        Set(ByVal Value As Boolean)
            blnRenderAdjustment = Value
        End Set
    End Property
    Property RenderScaleTarget() As Boolean
        Get
            Return blnRenderScaleTarget
        End Get
        Set(ByVal Value As Boolean)
            blnRenderScaleTarget = Value
        End Set
    End Property
    Property ScaleRadius() As Boolean
        Get
            Return blnScaleRadius
        End Get
        Set(ByVal Value As Boolean)
            blnScaleRadius = Value
        End Set
    End Property
    Property RenderTTK() As Boolean
        Get
            Return blnRenderTTK
        End Get
        Set(ByVal Value As Boolean)
            blnRenderTTK = Value
        End Set
    End Property

    Property RenderHeatMap() As Boolean
        Get
            Return blnRenderHeatMap
        End Get
        Set(ByVal Value As Boolean)
            blnRenderHeatMap = Value
        End Set
    End Property
    Property RenderHeatTitle() As Boolean
        Get
            Return blnRenderHeatTitle
        End Get
        Set(ByVal Value As Boolean)
            blnRenderHeatTitle = Value
        End Set
    End Property
    Property RenderHeatBars() As Boolean
        Get
            Return blnRenderHeatBars
        End Get
        Set(ByVal Value As Boolean)
            blnRenderHeatBars = Value
        End Set
    End Property
    Property RenderHeatAdjust() As Boolean
        Get
            Return blnRenderHeatAdjust
        End Get
        Set(ByVal Value As Boolean)
            blnRenderHeatAdjust = Value
        End Set
    End Property

    Property RenderDropInfo() As Boolean
        Get
            Return blnRenderDropInfo
        End Get
        Set(ByVal Value As Boolean)
            blnRenderDropInfo = Value
        End Set
    End Property
    Property RenderBulletDrop() As Boolean
        Get
            Return blnRenderBulletDrop
        End Get
        Set(ByVal Value As Boolean)
            blnRenderBulletDrop = Value
        End Set
    End Property
    Property RenderAmmoInfo() As Boolean
        Get
            Return blnRenderAmmoInfo
        End Get
        Set(ByVal Value As Boolean)
            blnRenderAmmoInfo = Value
        End Set
    End Property
    Property RenderGrid() As Boolean
        Get
            Return blnRenderGrid
        End Get
        Set(ByVal Value As Boolean)
            blnRenderGrid = Value
        End Set
    End Property
    Property RenderTTKGrid() As Boolean
        Get
            Return blnRenderTTKGrid
        End Get
        Set(ByVal Value As Boolean)
            blnRenderTTKGrid = Value
        End Set
    End Property
    Property RenderDropGrid() As Boolean
        Get
            Return blnRenderDropGrid
        End Get
        Set(ByVal Value As Boolean)
            blnRenderDropGrid = Value
        End Set
    End Property


    Property DamageMax() As Double
        Get
            Return dblTTKDamageMax
        End Get
        Set(ByVal Value As Double)
            dblTTKDamageMax = Value
        End Set
    End Property
    Property DamageMin() As Double
        Get
            Return dblTTKDamageMin
        End Get
        Set(ByVal Value As Double)
            dblTTKDamageMin = Value
        End Set
    End Property
    Property RangeMax() As Double
        Get
            Return dblTTKRangeMax
        End Get
        Set(ByVal Value As Double)
            dblTTKRangeMax = Value
        End Set
    End Property
    Property RangeMin() As Double
        Get
            Return dblTTKRangeMin
        End Get
        Set(ByVal Value As Double)
            dblTTKRangeMin = Value
        End Set
    End Property

    Property BulletDrop() As Double
        Get
            Return dblBulletDrop
        End Get
        Set(ByVal Value As Double)
            dblBulletDrop = Value
        End Set
    End Property
    Property BulletVelocity() As Double
        Get
            Return dblBulletVelocity
        End Get
        Set(ByVal Value As Double)
            dblBulletVelocity = Value
        End Set
    End Property
    Property MaxDistance() As Double
        Get
            Return dblMaxDistance
        End Get
        Set(ByVal Value As Double)
            dblMaxDistance = Value
        End Set
    End Property
    Property TargetRange() As Double
        Get
            Return dblTargetRange
        End Get
        Set(ByVal Value As Double)
            dblTargetRange = Value
        End Set
    End Property
    Property RateOfFire() As Double
        Get
            Return dblRateOfFire
        End Get
        Set(ByVal Value As Double)
            dblRateOfFire = Value
        End Set
    End Property

    Property Title() As String
        Get
            Return strTitle
        End Get
        Set(ByVal Value As String)
            strTitle = Value
        End Set
    End Property
    Property Gun() As String
        Get
            Return strGunName
        End Get
        Set(ByVal Value As String)
            strGunName = Value
        End Set
    End Property
    Property FileName() As String
        Get
            Return strFileName
        End Get
        Set(ByVal Value As String)
            strFileName = Value
        End Set
    End Property
    Property Info() As String
        Get
            Return strInfo
        End Get
        Set(ByVal Value As String)
            strInfo = Value
        End Set
    End Property
    Property SubText() As String
        Get
            Return strSubText
        End Get
        Set(ByVal Value As String)
            strSubText = Value
        End Set
    End Property

    'Drop this in favor of individual values.
    Property BackgroundColor() As String
        Get
            Return strBackgroundColor
        End Get
        Set(ByVal Value As String)
            strBackgroundColor = Value
        End Set
    End Property

    Property RecoilUp() As Double
        Get
            Return dblRecoilUp
        End Get
        Set(ByVal Value As Double)
            dblRecoilUp = Value
        End Set
    End Property
    Property RecoilRight() As Double
        Get
            Return dblRecoilRight
        End Get
        Set(ByVal Value As Double)
            dblRecoilRight = Value
        End Set
    End Property
    Property RecoilLeft() As Double
        Get
            Return dblRecoilLeft
        End Get
        Set(ByVal Value As Double)
            dblRecoilLeft = Value
        End Set
    End Property
    'Recoil decrease drop v2.23
    ' Property RecoilDecrease() As Double
    'Get
    'Return dblRecoilDecrease
    'End Get
    'Set(ByVal Value As Double)
    'dblRecoilDecrease = Value
    'End Set
    'End Property
    Property FirstShot() As Double
        Get
            Return dblFirstShot
        End Get
        Set(ByVal Value As Double)
            dblFirstShot = Value
        End Set
    End Property

    Property GridLineSpace() As Double
        Get
            Return dblGridLineSpace
        End Get
        Set(ByVal Value As Double)
            dblGridLineSpace = Value
        End Set
    End Property

    Property SpreadMin() As Double
        Get
            Return dblSpreadMin
        End Get
        Set(ByVal Value As Double)
            dblSpreadMin = Value
        End Set
    End Property
    Property SpreadInc() As Double
        Get
            Return dblSpreadInc
        End Get
        Set(ByVal Value As Double)
            dblSpreadInc = Value
        End Set
    End Property

    Property AdjRecoilH() As Double
        Get
            Return dblAdjRecoilH
        End Get
        Set(ByVal Value As Double)
            dblAdjRecoilH = Value
        End Set
    End Property
    Property AdjRecoilV() As Double
        Get
            Return dblAdjRecoilV
        End Get
        Set(ByVal Value As Double)
            dblAdjRecoilV = Value
        End Set
    End Property
    Property AdjSpreadMin() As Double
        Get
            Return dblAdjSpreadMin
        End Get
        Set(ByVal Value As Double)
            dblAdjSpreadMin = Value
        End Set
    End Property
    Property AdjSpreadInc() As Double
        Get
            Return dblAdjSpreadInc
        End Get
        Set(ByVal Value As Double)
            dblAdjSpreadInc = Value
        End Set
    End Property
    Property Burst() As Integer
        Get
            Return intBurst
        End Get
        Set(ByVal Value As Integer)
            intBurst = Value
        End Set
    End Property
    Property BulletsPerBurst() As Integer
        Get
            Return intBulletsPerBurst
        End Get
        Set(ByVal Value As Integer)
            intBulletsPerBurst = Value
        End Set
    End Property
    Property Scale() As Integer
        Get
            Return intScale
        End Get
        Set(ByVal Value As Integer)
            intScale = Value
        End Set
    End Property

    Property HealthPercent() As Double
        Get
            Return dblHealthPercent
        End Get
        Set(ByVal Value As Double)
            dblHealthPercent = Value
        End Set
    End Property
    Property Image() As Bitmap
        Get
            Return bmpImage
        End Get
        Set(ByVal Value As Bitmap)
            bmpImage = Value
        End Set
    End Property
    Property Mask() As Bitmap
        Get
            Return bmpMask
        End Get
        Set(ByVal Value As Bitmap)
            bmpMask = Value
        End Set
    End Property
    Property HeatMap() As Bitmap
        Get
            Return bmpHeatMap
        End Get
        Set(ByVal Value As Bitmap)
            bmpHeatMap = Value
        End Set
    End Property
    Property TTK() As Bitmap
        Get
            Return bmpTTK
        End Get
        Set(ByVal Value As Bitmap)
            bmpTTK = Value
        End Set
    End Property
    Property Silh() As Bitmap
        Get
            Return bmpSilh
        End Get
        Set(ByVal Value As Bitmap)
            bmpSilh = Value
        End Set
    End Property

    Property ImageGraphic() As Graphics
        Get
            Return grhImageGraphic
        End Get
        Set(ByVal Value As Graphics)
            grhImageGraphic = Value
        End Set
    End Property
    Property MaskGraphic() As Graphics
        Get
            Return grhMaskGraphic
        End Get
        Set(ByVal Value As Graphics)
            grhMaskGraphic = Value
        End Set
    End Property
    Property HeatGraphic() As Graphics
        Get
            Return grhHeatGraphic
        End Get
        Set(ByVal Value As Graphics)
            grhHeatGraphic = Value
        End Set
    End Property
    Property TTKGraphic() As Graphics
        Get
            Return grhTTKGraphic
        End Get
        Set(ByVal Value As Graphics)
            grhTTKGraphic = Value
        End Set
    End Property
#End Region

End Class

