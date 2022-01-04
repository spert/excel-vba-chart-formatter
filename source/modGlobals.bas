Option Explicit

Public ErrorMod As New clsError

Enum MyFrequency
    None = 0
    Daily = 1
    Weekly = 2
    Monthly = 3
    Quarterly = 4
    Yearly = 5
End Enum

Enum MyPosition
    None = 0
    Top = 1
    Bottom = 2
End Enum

Public Type MyShapeSize
    
    Top As Double
    
    Left As Double
    
    Height As Double
    
    With As Double
    
End Type

Public Type MyFont
    
    Name As String
    
    Size As Integer
    
    Bold As Boolean
    
    RGB As Long
    
End Type

Public Type MyPlotArea
       
    Font As MyFont
    
    Size As MyShapeSize
    
End Type

Public Type MyLegend
       
    Font As MyFont
    
    Size As MyShapeSize
    
End Type

Public Type MyAxisTitle
    
    Font As MyFont
        
End Type


Public Type MyTitle
    
    BoxEnabled As Boolean
    
    BackgroundRGB As Long
    
    Text As String
    
    Position As MyPosition
    
    Font As MyFont
    
    Size As MyShapeSize
    
End Type

Public Type MySourceTextBox
   
    Font As MyFont
    
    Size As MyShapeSize
    
    TextAlignment As MsoParagraphAlignment
    
End Type

Public Type MyChart

   FirstCellOfOutput As Variant
    
   SheetName As String
   
   Top As Integer
   
   Left As Integer

   Height As Integer
   
   With As Integer
   
   ColumnWith As Double
   
   Language_Descriptions As Variant
   
   Language_ScaleValidation As Variant
         
   SeriesDataOffset As Variant
       
   AxisTitle As MyAxisTitle
         
   Title As MyTitle
   
   Legend As MyLegend
   
   PlotArea As MyPlotArea
   
   SourceTextBox As MySourceTextBox
   
   SeriesColor As Variant
   
   SeriesWeight As Double

End Type

Public Type MyNewDateScale

    Rescaled As Boolean

    MinDate As Date

    MaxDate As Date

    BaseUnit As XlTimeUnit

    MinorUnit As XlTimeUnit

    MajorUnit As Integer
    
    MajorUnitScale As XlTimeUnit
    
    FormatCode As String

End Type



