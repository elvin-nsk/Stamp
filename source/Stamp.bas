Attribute VB_Name = "Stamp"
'===============================================================================
'   Макрос          : Stamp
'   Версия          : 2024.06.19
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Stamp"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.06.19"

'===============================================================================
' # Globals

Private Const DUPLICATE_OFFSET As Double = 20
Private Const COLOR_BLACK As String = "RGB255,USER,0,0,0"
Private Const COLOR_WHITE As String = "RGB255,USER,255,255,255"
Private Const SECOND_CONTOUR As Double = 0.65
Private Const SECOND_CONTOUR_COLOR As String = "RGB255,USER,255,0,0"

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim Cfg As PresetsConfig
    If Not ShowStampView(Cfg) Then Exit Sub
    
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    
    BoostStart APP_DISPLAYNAME
    
    ProcessStamp Shapes, Cfg
    
    Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Sub ProcessStamp( _
                ByVal SourceShapes As ShapeRange, _
                ByVal Cfg As PresetsConfig _
            )
    Dim Shapes As ShapeRange: Set Shapes = _
        SourceShapes.Duplicate( _
            OffsetY:=-(SourceShapes.SizeHeight + DUPLICATE_OFFSET) _
        )
    Set Shapes = Shapes.UngroupAllEx
    
    Dim Boundary As Shape: Set Boundary = CreateBoundary(Shapes)
    Dim FirstContour As Shape: Set FirstContour = _
        MakeSeparatedContour(Boundary, Cfg!FirstContour)
    With FirstContour
        .OrderToBack
        .Outline.SetNoOutline
        .Fill.UniformColor = CreateColor(COLOR_WHITE)
    End With
    Dim SecondContour As Shape: Set SecondContour = _
        MakeSeparatedContour(Boundary, SECOND_CONTOUR)
    With SecondContour
        .OrderToFront
        .Outline.Color = CreateColor(SECOND_CONTOUR_COLOR)
        .Fill.ApplyNoFill
    End With
    Boundary.Delete
        
    Dim TextShapes As ShapeRange: Set TextShapes = _
        Shapes.Shapes.FindShapes(Type:=cdrTextShape)
    ModifyOutlines TextShapes, Cfg!AddToTextOutline
    Shapes.RemoveRange TextShapes
    ModifyOutlines Shapes, Cfg!AddToVectorOutline
    
    Shapes.AddRange TextShapes
    Shapes.Add FirstContour
    If Cfg!Invert Then Invert Shapes
    Shapes.Add SecondContour
    If Cfg!Mirror Then Shapes.Flip cdrFlipHorizontal

End Sub

Private Function MakeSeparatedContour( _
                     ByVal Shape As Shape, _
                     ByVal Offset As Double _
                 ) As Shape
    Set MakeSeparatedContour = _
        Shape.CreateContour( _
            Direction:=cdrContourOutside, _
            Offset:=Offset, _
            Steps:=1 _
        ).Separate(1)
End Function

Private Sub ModifyOutlines(ByVal Shapes As ShapeRange, ByVal AddWidth As Double)
    If AddWidth <= 0 Then Exit Sub
    Dim Shape As Shape
    For Each Shape In Shapes
        ModifyOutlineOfShape Shape, AddWidth
    Next Shape
End Sub

Private Sub ModifyOutlineOfShape(ByVal Shape As Shape, ByVal AddWidth As Double)
    With Shape.Outline
        If .Width = 0 Then
            .Width = AddWidth
            If Shape.Fill.Type = cdrUniformFill Then
                .Color = Shape.Fill.UniformColor
            Else
                .Color = CreateColor(COLOR_BLACK)
            End If
        Else
            .Width = .Width + AddWidth
        End If
        .LineCaps = cdrOutlineButtLineCaps
        .LineJoin = cdrOutlineRoundLineJoin
    End With
End Sub

Private Sub Invert(ByVal Shapes As ShapeRange)
    Dim Shape As Shape
    For Each Shape In Shapes
        InvertShape Shape
    Next Shape
End Sub

Private Sub InvertShape(ByVal Shape As Shape)
    With Shape.Fill
        If .Type = cdrUniformFill Then _
            .UniformColor = InvertedColor(.UniformColor)
    End With
    With Shape.Outline
        If Not .Type = cdrNoOutline Then .Color = InvertedColor(.Color)
    End With
End Sub

Private Property Get InvertedColor(ByVal Color As Color) As Color
    If IsDark(Color) Then
        Set InvertedColor = CreateColor(COLOR_WHITE)
    Else
        Set InvertedColor = CreateColor(COLOR_BLACK)
    End If
End Property

Private Function ShowStampView(ByRef Cfg As PresetsConfig) As Boolean
    Dim Fields As Dictionary: Set Fields = ConfigFields
    Set Cfg = PresetsConfig.New_("elvin_" & APP_NAME, Fields)
    Dim View As New StampView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg.Current, _
            View:=View, _
            ControlNames:=Fields.Keys _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowStampView = View.IsOk
End Function

Private Property Get ConfigFields() As Dictionary
    Set ConfigFields = New Dictionary
    With ConfigFields
        !FirstContour = 2#
        !AddToTextOutline = 0.15
        !AddToVectorOutline = 0.1
        !Invert = True
        !Mirror = True
    End With
End Property

'===============================================================================
' # Tests

Private Sub testSomething()
'
End Sub
