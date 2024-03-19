Attribute VB_Name = "DefinedDataType"
Public Type gridRange
    SheetId As Variant
    startRowIndex As Variant
    endRowIndex As Variant
    startColumnIndex As Variant
    endColumnIndex As Variant
End Type

Public Type unionFieldScope
    rng As gridRange
    SheetId As Variant
    allSheets As Variant
End Type


