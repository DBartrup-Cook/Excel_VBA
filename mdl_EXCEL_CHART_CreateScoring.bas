Attribute VB_Name = "mdl_EXCEL_CHART_CreateScoring"
Option Explicit


Public Sub Scoring_Test()
    'Places Score chart over first range in procedure arguments.
    On Error Resume Next
        ActiveSheet.Shapes("Scoring_Chart_Container").Delete
        ActiveSheet.ChartObjects("Scoring_Chart").Delete
    On Error GoTo 0

    CreateScoringWidget ActiveSheet.Range("H5:Q11")

End Sub

Private Sub CreateScoringWidget(DisplayRange As Range)

    Dim oChtContainer As Shape
    Dim oMarker As Shape
    Dim oCht As ChartObject
    Dim x As Long
    Dim GradStop1() As Variant
    Dim GradStop2() As Variant

    With DisplayRange
        'Add data table.
        With .Parent
            .Range("A1") = "Indicator:"
            .Range("B1") = WorksheetFunction.RandBetween(0, 100) / 100
            .Range("A2") = "Marker Size:"
            .Range("B2") = 0.05
            .Range("A3") = "Marker Spacer:"
            .Range("B3").FormulaR1C1 = "=R1C2-(R2C2/2)"
            .Range("B1:B3").NumberFormat = "0%"
        End With
        
        'Add rectangle to background.
        Set oChtContainer = .Parent.Shapes.AddShape _
            (msoShapeRectangle, .Left, .Top, .Width, .Height)
    End With
    
    'Format rectangle
    With oChtContainer
        .Name = "Scoring_Chart_Container"
        With .Fill
            .TwoColorGradient msoGradientHorizontal, 1
            With .GradientStops(1)
                .Color = RGB(28, 28, 28)
                .Position = 0
                .Transparency = 0
            End With
            With .GradientStops(2)
                .Color = RGB(127, 127, 127)
                .Position = 1
                .Transparency = 0
            End With
            .GradientStops.Insert RGB:=RGB(242, 242, 242), _
                Position:=0.5, Transparency:=0
            .Visible = msoTrue
        End With
        With .Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Weight = 1
        End With
    End With
    
    'Add chart
    With DisplayRange
        Set oCht = .Parent.ChartObjects.Add _
            (.Left, .Top, .Width, .Height)
    End With
    
    With oCht
        .Name = "Scoring_Chart"
        .Parent.Shapes(.Name).Fill.Visible = msoFalse
        
        With .Chart
            .ChartType = xlBarStacked
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1
            .Axes(xlValue).MajorUnit = 0.1
            .Axes(xlValue).TickLabels.NumberFormat = "0%"
            
            'Add five slicers.  GradStop1 & GradStop2 are the colours for each slicer.
            GradStop1 = Array(1118719, 683236, 65535, 5296274, 5287936)
            GradStop2 = Array(138, 409992, 37525, 1662269, 2181120)
            For x = LBound(GradStop1) To UBound(GradStop1)
                With .SeriesCollection.NewSeries
                    .Name = "Slicer" & x
                    .Values = 0.2
                    
                    With .Format.ThreeD
                        .BevelTopType = msoBevelSoftRound
                        .BevelTopInset = 12
                        .BevelTopDepth = 4
                        .PresetMaterial = msoMaterialTranslucentPowder
                    End With
                    
                    With .Format.Fill
                        .TwoColorGradient msoGradientHorizontal, 1
                        With .GradientStops(1)
                            .Color = GradStop1(x)
                            .Position = 0
                            .Transparency = 0
                        End With
                        With .GradientStops(2)
                            .Color = GradStop2(x)
                            .Position = 1
                            .Transparency = 0
                        End With
                    End With
                End With
            Next x
            
            'Add MarkerSpacer
            With .SeriesCollection.NewSeries
                .Name = "MarkerSpacer"
                .Values = "'" & DisplayRange.Parent.Name & "'!$B$3"
                .AxisGroup = 2
                .Format.Fill.Visible = msoFalse
                .Format.Line.Visible = msoFalse
            End With
            
            'Add SlicerMarker Series
            With .SeriesCollection.NewSeries
                .Name = "='" & DisplayRange.Parent.Name & "'!$B$1" '"SlicerMarker"
                .Values = "'" & DisplayRange.Parent.Name & "'!$B$2"
                .AxisGroup = 2
            End With
                
            .Axes(xlValue, xlSecondary).MinimumScale = 0
            .Axes(xlValue, xlSecondary).MaximumScale = 1
            
            .SetElement (msoElementLegendNone)
            .SetElement (msoElementPrimaryValueGridLinesNone)
            .SetElement (msoElementPlotAreaNone)
            .Axes(xlCategory).Delete
            .Axes(xlValue, xlSecondary).Delete
            
            'Create SlicerMarker shape.
            With DisplayRange
                Set oMarker = .Parent.Shapes.AddShape _
                    (msoShapeFlowchartMerge, .Left, .Top, .Cells(1, 1).Width, .Resize(.Rows.Count, 1).Height * 0.6)
            End With
            With oMarker
                .Name = "DownMarker"
                With .Fill
                    .TwoColorGradient msoGradientHorizontal, 1
                    With .GradientStops(1)
                        .Color = 9326848
                        .Position = 0
                        .Transparency = 0
                    End With
                    With .GradientStops(2)
                        .Color = 4335104
                        .Position = 1
                        .Transparency = 0
                    End With
                End With
                With .ThreeD
                    .BevelTopType = msoBevelSoftRound
                    .BevelTopInset = 12
                    .BevelTopDepth = 4
                End With
                
                .Copy
                
            End With
        
            With .SeriesCollection(7).Points(1)
                'Copy "DownMarker" and place into Point(1) of "SlicerMarker" series.
                .Select
                .Paste
                
                'Add a label above the marker.
                .ApplyDataLabels
                With .Parent.DataLabels
                    With .Format.TextFrame2.TextRange.Font.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(255, 255, 255)
                    End With
                    .ShowSeriesName = True
                    .ShowValue = False
                End With
                .DataLabel.Top = 5
            End With
            
            oMarker.Delete
            
            .ChartGroups(2).GapWidth = 65
            
        End With
    End With
    
    DisplayRange.Parent.Shapes.Range(Array("Scoring_Chart_Container", "Scoring_Chart")).Group

End Sub
