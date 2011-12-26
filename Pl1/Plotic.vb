﻿Public Class Plotic
    Private strTitle As String
    Private bmpImage As Bitmap = New Bitmap(2000, 2000)
    Private bmpMask As Bitmap = New Bitmap(2000, 2000)
    Private bmpHeatMap As Bitmap = New Bitmap(2000, 2000)
    Private grhImageGraphic As Graphics = Graphics.FromImage(bmpImage)
    Private grhMaskGraphic As Graphics = Graphics.FromImage(bmpMask)
    Private grhHeatGraphic As Graphics = Graphics.FromImage(bmpHeatMap)
    Private dblRecoilUp As Double
    Private dblRecoilRight As Double
    Private dblRecoilLeft As Double
    Private dblSpreadMin As Double
    Private dblSpreadInc As Double
    Private dblAdjRecoilH As Double
    Private dblAdjRecoilV As Double
    Private dblAdjSpreadMin As Double
    Private dblAdjSpreadInc As Double
    Private intBurst As Integer
    Private intBulletsPerBurst As Integer
    Private dblFirstShot As Double

    Private intScale As Integer

    Property Title() As String
        Get
            Return strTitle
        End Get
        Set(ByVal Value As String)
            strTitle = Value
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
    Property FirstShot() As Double
        Get
            Return dblFirstShot
        End Get
        Set(ByVal Value As Double)
            dblFirstShot = Value
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
    Property Scale() As Integer
        Get
            Return intScale
        End Get
        Set(ByVal Value As Integer)
            intScale = Value
        End Set
    End Property


    Public Sub New()
        '        Dim b As Bitmap = New Bitmap(2000, 2000)
        'Dim g As Graphics = Graphics.FromImage(b)
        'Me.Image = b
        'Me.Mask = b
        'Me.HeatMap = b
        '        Me.grhImageGraphic = Graphics.FromImage(Me.Image)
        '       Me.grhMaskGraphic = Graphics.FromImage(Me.Mask)
        '      Me.grhHeatGraphic = Graphics.FromImage(Me.HeatMap)
        'Initialize the Bitmaps and Graphics objects
        'Me.grhImageGraphic = g
    End Sub
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
    Public Sub SaveMask()

    End Sub
End Class
