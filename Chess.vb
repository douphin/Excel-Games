Sub CreateChessBoard()

Dim ULRS As String
Dim BR, tBR As Range
Dim UL_R, UL_C As Integer
Dim PieceArr(4) As String

ULRS = CStr(ActiveSheet.Cells(1, 1))

Set BR = Range(ULRS).Resize(8, 8)


UL_R = Range(ULRS).Row
UL_C = Range(ULRS).Column
MsgBox Cells(1, 1).Height
MsgBox Cells(1, 1).Width




For i = 1 To 8
    BR.Cells(2, i) = "Pawn"
    BR.Cells(7, i) = "Pawn"
    
    For j = 1 To 8
        BR.Cells(i, j).BorderAround _
        Weight:=xlThin
    Next
    
    BR.Cells(7, i).Interior.Color = RGB(0, 0, 0)
    BR.Cells(8, i).Interior.Color = RGB(0, 0, 0)
    BR.Cells(7, i).Font.Color = RGB(255, 255, 255)
    BR.Cells(8, i).Font.Color = RGB(255, 255, 255)
    
Next

BR.BorderAround _
Weight:=xlThick

BR.Cells(1, 1) = "Rook"
BR.Cells(8, 1) = "Rook"
BR.Cells(1, 8) = "Rook"
BR.Cells(8, 8) = "Rook"

BR.Cells(1, 2) = "Knight"
BR.Cells(8, 2) = "Knight"
BR.Cells(1, 7) = "Knight"
BR.Cells(8, 7) = "Knight"

BR.Cells(1, 3) = "Bishop"
BR.Cells(8, 3) = "Bishop"
BR.Cells(1, 6) = "Bishop"
BR.Cells(8, 6) = "Bishop"

BR.Cells(1, 4) = "Queen"
BR.Cells(8, 4) = "Queen"
BR.Cells(1, 5) = "King"
BR.Cells(8, 5) = "King"



End Sub
Sub ClearBoard()

    ULRS = CStr(ActiveSheet.Cells(1, 1))

    Set BR = Range(ULRS).Resize(8, 8)

    Dim shp As Shape

    For Each shp In ActiveSheet.Shapes
        
        shp.Delete
        
    Next shp
    
    BR.ClearContents
    BR.Interior.ColorIndex = xlColorIndexNone
    BR.Borders.ColorIndex = xlColorIndexNone
    
    ActiveSheet.Buttons.Add(20, 20, 55, 50).Select
    Application.CutCopyMode = False
    Selection.OnAction = "ClearBoard"
    Selection.Characters.Text = "ClearBoard"
    Range("A1").Select
    
    ActiveSheet.Buttons.Add(20, 100, 55, 50).Select
    Application.CutCopyMode = False
    Selection.OnAction = "NewBoard"
    Selection.Characters.Text = "NewBoard"
    Range("A1").Select
    
End Sub
Sub NewBoard()


    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 48, 291, 79.5, 34.5).Select
    'Selection.OnAction = "Test1"
    'Range("C25").Select
    
    Range("B2:Z50").ClearContents
    
    ' Set up piece arrays for automatic piece naming
    Dim PieceArr As Variant
    Dim PawnArr As Variant
    Dim CurrPieceArr As Variant
    
    PieceArr = Array("Rook", "Knight", "Bishop", "Queen", "King", "Bishop", "Knight", "Rook")
    PawnArr = Array("Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn")
    'TestArr(3) = Array("", "", "", "", "", "", "", "")
    
    
    ' Set up other variables
    Dim ULRS As String ' Upper Left Row String
    Dim BR, tBR As Range ' Board Range, temporary Board Range
    Dim UL_Row, UL_Col As Integer ' Upper Left Row, Upper Left Column
    Dim LeftCount, TopCount As Double ' Total width from left border, total height from top border
    Dim LeftIter, TopIter As Double ' Left/Top iterator, increment to add to Left/TopCount
    Dim textColor As Integer ' Used to set the text color of the pieces
    Dim ci, cXR(2), cXG(2), cXB(2) As Integer
    
    cXR(0) = 146
    cXR(1) = 25
    
    cXG(0) = 208
    cXG(1) = 107
    
    cXB(0) = 80
    cXB(1) = 36
    

    ' Get a string of the selected cell to create the board
    ULRS = CStr(ActiveSheet.Cells(1, 1))

    Set BR = Range(ULRS).Resize(8, 8)
    
    With BR.Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    
    UL_Row = Range(ULRS).Row
    UL_Col = Range(ULRS).Column
    
    LeftCount = 0
    TopCount = 0
    
    ' Get the total width from left border to starting point
    For i = 1 To UL_Col - 1
        LeftCount = LeftCount + Cells(i, 1).Width
    Next
    
    ' Get the total height from top border to starting point
    For i = 1 To UL_Row - 1
        TopCount = TopCount + Cells(1, i).Height
    Next
    
    ' Set up variables for the array
    TopIter = TopCount
    h = 1
    CurrPieceArr = PieceArr
    
    textColor = 0

    
    ' For each row (moving down)
    For i = 1 To 8
        LeftIter = LeftCount
        
        ' If top 2 rows or bottom 2 rows
       
            ' For each column (moving right)
        For j = 1 To 8
             If i < 3 Or i > 6 Then
                ' Create each piece as a shape
                ActiveSheet.Shapes.AddShape(msoShapeRectangle, LeftIter, TopIter, Cells(UL_Row, _
                UL_Col).Width, Cells(UL_Row, UL_Col).Height).Select
                

                
                ' Add the name of the piece, pulling from one of two piece arrays, adjusts for index at 0
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = CurrPieceArr(j - 1)
                
                ' Set color of text, setting top to black and bottom to white
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(textColor, textColor, textColor)
                
                ' Making the color of the shape transparent so the colors of the board will show through
                Selection.ShapeRange.Fill.Visible = msoFalse
                
                ' Set buttons to show potential moves
                Selection.OnAction = "PotentialLocation" ' CurrPieceArr(j - 1)
                
                ' Centering the text verticaling and horizontaly
                With Selection.ShapeRange.TextFrame2
                    .VerticalAnchor = msoAnchorMiddle
                    .HorizontalAnchor = msoAnchorCenter
                End With
                
                ' Moving one cells to the right by using the width of the cell
                LeftIter = LeftIter + Cells(UL_Row, UL_Col).Width
            
            ' This is used to determine which piece array to use

            End If
            
            ci = ((i Mod 2) + (j Mod 2)) Mod 2
            'BR.Cells(i, j).Interior.Color = RGB(cXR(ci), cXG(ci), cXB(ci))
        Next
        ' Weird logic used to set up piece names and colors
        If i > 4 Then
            textColor = 125
        End If
        
        If i = 1 Then
            CurrPieceArr = PawnArr
        ElseIf i = 7 Then
            CurrPieceArr = PieceArr
        
        End If
        
        ' Move down a row, using the height of the current row
        TopIter = TopIter + Cells(UL_Row, UL_Col).Height
    Next
    
    
End Sub


Sub Macro8()
'
' Macro8 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Rectangle 3620")).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
End Sub



Function GetCellRow(ByVal Top As Double) As Integer

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Row As Integer ' Upper Left Row
    Dim Curr_Row As Integer
    Dim TopCount As Double ' Total height from top border
    
    ULRS = CStr(ActiveSheet.Cells(1, 1))

    Set BR = Range(ULRS).Resize(8, 8)
    
    UL_Row = Range(ULRS).Row
    
    TopCount = 0
    
    For i = 1 To UL_Row - 1
        TopCount = TopCount + Cells(1, i).Height
    Next
    
    Curr_Row = 1
    Do While TopCount <> Top And Curr_Row < 9
        TopCount = TopCount + Cells(1, i).Height
        Curr_Row = Curr_Row + 1
    Loop
    
    GetCellRow = Curr_Row

End Function

Function GetCellCol(ByVal Left As Double) As Integer

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Col As Integer ' Upper Left Column
    Dim Curr_Col As Integer
    Dim LeftCount As Double ' Total width from left border
    
    ULRS = CStr(ActiveSheet.Cells(1, 1))

    Set BR = Range(ULRS).Resize(8, 8)
    
    UL_Col = Range(ULRS).Column
    
    LeftCount = 0
    
    For i = 1 To UL_Col - 1
        LeftCount = LeftCount + Cells(i, 1).Width
    Next
    
    Curr_Col = 1
    Do While LeftCount <> Left And Curr_Col < 9
        LeftCount = LeftCount + Cells(i, 1).Width
        Curr_Col = Curr_Col + 1
    Loop
    
    GetCellCol = Curr_Col

End Function





Sub PotentialLocation()



    Dim CallingShape As Shape
    Dim ShpPiece As String
    Dim ShpTop, ShpLeft As Double
    Dim Curr_Row, Curr_Col As Integer

    
    
    Set CallingShape = ActiveSheet.Shapes(Application.Caller)
    ShpPiece = CallingShape.TextFrame2.TextRange.Characters.Text
    ShpTop = CallingShape.Top
    ShpLeft = CallingShape.Left
    
    Curr_Row = GetCellRow(ShpTop)
    Curr_Col = GetCellCol(ShpLeft)

    
    
    
    
    

    If ShpPiece = "Pawn" Then
        Call Pawn(Curr_Row, Curr_Col, CallingShape)
        
    ElseIf ShpPiece = "Rook" Then
        Call Rook(Curr_Row, Curr_Col, CallingShape)
        
    ElseIf ShpPiece = "Knight" Then
        Call Knight(Curr_Row, Curr_Col, CallingShape)
        
    ElseIf ShpPiece = "Bishop" Then
        Call Bishop(Curr_Row, Curr_Col, CallingShape)
        
    ElseIf ShpPiece = "Queen" Then
        Call Queen(Curr_Row, Curr_Col, CallingShape)
        
    ElseIf ShpPiece = "King" Then
        Call King(Curr_Row, Curr_Col, CallingShape)
        
    End If
    
End Sub
Sub Pawn(ByVal Curr_Row, Curr_Col, CallingShape)

    MsgBox Curr_Row
    MsgBox Curr_Col

End Sub
Sub Rook(ByVal Curr_Row, Curr_Col, CallingShape)

    Dim ULRS As String ' Upper Left Row String
    Dim BR, tBR As Range ' Board Range, temporary Board Range
    Dim UL_Row, UL_Col As Integer ' Upper Left Row, Upper Left Column
    
    Dim deleted As Boolean
    
    ' Get a string of the selected cell to create the board
    ULRS = CStr(ActiveSheet.Cells(1, 1))

    Set BR = Range(ULRS).Resize(8, 8)
    
    UL_Row = Range(ULRS).Row
    UL_Col = Range(ULRS).Column
    
    Dim Lb, Rb, Tb, Ub As Double ' left, right, top, bottom bound
    Dim VerticalBound() As Double
 
    
    For Each shp In ActiveSheet.Shapes
        If shp.Line.ForeColor.RGB = RGB(255, 255, 0) Then
            shp.Delete
            deleted = True
        End If
        
        
    Next shp
    
    If deleted Then
            Exit Sub
        End If
    
    i = 0
    For Each shp In ActiveSheet.Shapes
        If shp.Left = CallingShape.Left And shp.Top <> CallingShape.Top Then
            ReDim Preserve VerticalBound(i)
            VerticalBound(i) = shp.Top
        End If
    Next shp
    
    For Each Vert In VerticalBound
        iVert = GetCellRow(Vert)
        
        If iVert + 1 < Curr_Row Then
            
            For i = iVert + 1 To Curr_Row - 1
                
                ActiveSheet.Shapes.AddShape(msoShapeRectangle, CallingShape.Left, Cells(i + UL_Row - 1, Curr_Col).Top, Cells(i, _
                Curr_Col).Width, Cells(i, Curr_Col).Height).Select
                

                ' Making the color of the shape transparent so the colors of the board will show through
                Selection.ShapeRange.Fill.Visible = msoFalse
                
                ' Set buttons to show potential moves
                Selection.OnAction = "PotentialLocation" ' CurrPieceArr(j - 1)
                
                Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 0)
                
            Next
            
        End If
    Next
    
    
    
    
    

End Sub
Sub Knight(ByVal Curr_Row, Curr_Col, CallingShape)

    MsgBox Curr_Row
    MsgBox Curr_Col

End Sub
Sub Bishop(ByVal Curr_Row, Curr_Col, CallingShape)

    MsgBox Curr_Row
    MsgBox Curr_Col

End Sub
Sub Queen(ByVal Curr_Row, Curr_Col, CallingShape)

    MsgBox Curr_Row
    MsgBox Curr_Col

End Sub
Sub King(ByVal Curr_Row, Curr_Col, CallingShape)

    MsgBox Curr_Row
    MsgBox Curr_Col

End Sub
Sub buffer()






































































End Sub
