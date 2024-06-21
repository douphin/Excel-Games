Sub ClearBoard()

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim shpRow, shpCol As Integer ' used to hold the row and column of the current shape
    Dim shpMacro As String 'Used to hold the macro of the current shape
    Dim NBLoc(4) As Double ' Array for storing the size and location of the NewBoard button
    Dim CBLoc(4) As Double ' Array for storing the size and location of the ClearBoard button
    
    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    ' Set default button size and location for NewBoard and ClearBoard
    NBLoc(0) = 20
    NBLoc(1) = 20
    NBLoc(2) = 55
    NBLoc(3) = 50
    
    CBLoc(0) = 20
    CBLoc(1) = 100
    CBLoc(2) = 55
    CBLoc(3) = 50

    ' Loop through shapes to clear the board
    For Each shp In ActiveSheet.Shapes
         
        ' Get shape parameters
        shpRow = GetCellRow(shp.Top)
        shpCol = GetCellCol(shp.Left)
        shpMacro = shp.OnAction
        
        If 1 <= shpRow And 1 <= shpCol And shpCol <= 8 And shpRow <= 8 Then ' Check to see if the shape is on board
        
            shp.Delete
        ElseIf shpMacro = ActiveWorkbook.Name + "!ClearBoard" Then ' If is ClearBoard Button, Save size and location
            NBLoc(0) = shp.Left
            NBLoc(1) = shp.Top
            NBLoc(2) = shp.Width
            NBLoc(3) = shp.Height
            
            shp.Delete
        ElseIf shpMacro = ActiveWorkbook.Name + "!NewBoard" Then ' If is NewBoard Button, Save size and location
            CBLoc(0) = shp.Left
            CBLoc(1) = shp.Top
            CBLoc(2) = shp.Width
            CBLoc(3) = shp.Height
        
            shp.Delete
        End If
        
        
    Next shp
    
    ' Clear borders and any checker colors
    BR.Interior.ColorIndex = xlColorIndexNone
    BR.Borders.ColorIndex = xlColorIndexNone
    
    ' Recreate NewBoard and ClearBoard buttons at saved size and location
    ActiveSheet.Buttons.Add(NBLoc(0), NBLoc(1), NBLoc(2), NBLoc(3)).Select
    Application.CutCopyMode = False
    Selection.OnAction = "ClearBoard"
    Selection.Characters.Text = "ClearBoard"
    Range("A1").Select
    
    ActiveSheet.Buttons.Add(CBLoc(0), CBLoc(1), CBLoc(2), CBLoc(3)).Select
    Application.CutCopyMode = False
    Selection.OnAction = "NewBoard"
    Selection.Characters.Text = "NewBoard"
    Range("A1").Select
    
End Sub
Sub NewBoard()

    ' Clear the old board
    ClearBoard
    
    ' Set up piece arrays for automatic piece naming
    Dim PieceArr As Variant
    Dim PawnArr As Variant
    Dim CurrPieceArr As Variant
    
    PieceArr = Array("Rook", "Knight", "Bishop", "Queen", "King", "Bishop", "Knight", "Rook")
    PawnArr = Array("Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn", "Pawn")
    
    
    ' Set up other variables
    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Row, UL_Col As Integer ' Upper Left Row, Upper Left Column
    Dim LeftCount, TopCount As Double ' Total width from left border, total height from top border
    Dim LeftIter, TopIter As Double ' Left/Top iterator, increment to add to Left/TopCount
    Dim textColor As Integer ' Used to set the text color of the pieces
    Dim ci, cXR(2), cXG(2), cXB(2) As Integer ' Variables for setting up a checkered board
    Dim checkeredboard As Boolean ' Turns checkered board on and off
    
    ' Set this to true if you want a checkered board
    checkeredboard = False
    
    ' These are rgb values for the checkered board
    cXR(0) = 146
    cXG(0) = 208
    cXB(0) = 80
    
    cXR(1) = 25
    cXG(1) = 107
    cXB(1) = 36
    

    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    UL_Row = Range(ULRS).Row ' Getting the top-most row of the board
    UL_Col = Range(ULRS).Column ' Getting the left-most column of the board
    
    LeftCount = 0
    TopCount = 0
    
    ' Set up grid for board
    With BR.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
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
            
            If checkeredboard Then
                ci = ((i Mod 2) + (j Mod 2)) Mod 2
                BR.Cells(i, j).Interior.Color = RGB(cXR(ci), cXG(ci), cXB(ci))
            End If
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

Function GetCellRow(ByVal Top As Double) As Integer

    ' This function will return the row number relative to the board of a shape
    ' based on the passed through distance from the top of the sheet

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Row As Integer ' Upper Left Row
    Dim Curr_Row As Integer ' Current Row
    Dim TopCount As Double ' Total height from top border
    
    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    UL_Row = Range(ULRS).Row ' Getting the top-most row of the board
    
    TopCount = 0 ' Set up for upcoming loop
    
    Curr_Row = 1 ' Set up for the next loop
    
    ' Adding up height to get to the edge of the board, moving from the top row of the sheet to just above the boards top row
    For i = 1 To UL_Row - 1
        TopCount = TopCount + Cells(1, i).Height
    Next
    
    ' Move through rows until we get one that is the same distance from the top of the sheet as the passed through value
    ' A limit of 9 is set to prevent an infinite loop
    Do While TopCount <> Top And Curr_Row <= 9
        TopCount = TopCount + Cells(1, i).Height
        Curr_Row = Curr_Row + 1
    Loop
    
    GetCellRow = Curr_Row ' Return our value

End Function

Function GetCellCol(ByVal Left As Double) As Integer

    ' This function will return the column number relative to the board of a shape
    ' based on the passed through distance from the left edge of the sheet

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Col As Integer ' Upper Left Column
    Dim Curr_Col As Integer ' Current Column
    Dim LeftCount As Double ' Total width from left border
    
    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    UL_Col = Range(ULRS).Column ' Getting the left-most column of the board
    
    LeftCount = 0 ' Set up for upcoming loop
    
    Curr_Col = 1 ' Set up for the next loop
     
    ' Adding up width to get to the edge of the board, moving from the left-most column of the sheet to just next to the boards left column
    For i = 1 To UL_Col - 1
        LeftCount = LeftCount + Cells(i, 1).Width
    Next
    
    ' Move through columns until we get one that is the same distance from the left edge of the sheet as the passed through value
    ' A limit of 9 is set to prevent an infinite loop
    Do While LeftCount <> Left And Curr_Col <= 9
        LeftCount = LeftCount + Cells(i, 1).Width
        Curr_Col = Curr_Col + 1
    Loop
    
    GetCellCol = Curr_Col ' Return our value
    

End Function
Function IsBlack(ByVal shp As Shape) As Boolean
    
    ' This function will return a boolean of true if the piece is black and false if white, used to determine friend or foe
    ' This function isn't really necessary, just kind of cleans up the code
    IsBlack = shp.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    
End Function

Function TogglePotentialLocations(ByVal CallingShape As Shape) As Boolean

    ' This function will toggle potential locations on and off
    ' It does this by finding yellow shapes and checking who they belong to
    ' if they belong to the calling shape then return true to exit the sub, else return false to repopulate yellows

    Dim deleted  As Boolean ' A boolean indicating if any yellow shapes have been deleted
    Dim OwnLocs As Boolean ' A boolean indicating if the yellow shapes deleted are belong to the calling piece

    deleted = False
    OwnLocs = False
 
    ' This will check to see if there are any yellow shapes, deleting them if found
    For Each shp In ActiveSheet.Shapes
        If shp.Line.ForeColor.RGB = RGB(255, 255, 0) Then

            If CallingShape.ID = shp.TextFrame2.TextRange.Characters.Text Then ' if the yellow shape belongs to the calling shape
                OwnLocs = True
            End If
            
            shp.Delete
            deleted = True
        End If
    Next shp
    
    ' If yellow shapes were deleted that belonged to the calling shape, then end it here
    ' This allows users to "toggle" the yellow on and off
    TogglePotentialLocations = deleted And OwnLocs

End Function

Function HasMovedFromStart(ByVal shpID As Integer) As Boolean

    Dim srcID As Integer ' source ID, the shape ID of the NewBoard button

    ' Loop through shapes to find the Newboard button and get it's shape ID
    For Each shp In ActiveSheet.Shapes
        If shp.OnAction = ActiveWorkbook.Name + "!NewBoard" Then ' If is NewBoard Button, Save ID
            srcID = shp.ID
        End If
    Next shp
    
    ' the shape ID is the nth shape created in the worksheet so far, so the 1452nd shape has an ID of 1452
    ' If the ID of the shape in question is more the 32 digits higher the source shape, then it isn't the original
    HasMovedFromStart = CInt(shpID) > CInt(srcID) + 32

End Function

Sub PotentialLocation()

    ' This is the macro assigned to all pieces on the board and will run the correct macro based on the name of the piece
    ' I don't know that this macro is necessarily necessary, but I'll keep it for now
    ' If nothing else it just keeps the other subs a little cleaner
    
    Dim CallingShape As Shape ' Shape object variable for the calling shape
    Dim ShpPiece As String ' String holding which piece is making the call i.e. rook, pawn, knight etc
    Dim ShpTop, ShpLeft As Double ' top and left values of the calling shape
    Dim Curr_Row, Curr_Col As Integer ' Current row and column of the calling shape relative to the board
 
    ' Set variables
    Set CallingShape = ActiveSheet.Shapes(Application.Caller)
    
    ShpPiece = CallingShape.TextFrame2.TextRange.Characters.Text
    
    ShpTop = CallingShape.Top
    ShpLeft = CallingShape.Left
    
    Curr_Row = GetCellRow(ShpTop)
    Curr_Col = GetCellCol(ShpLeft)

      
    ' Call the correct sub based on the piece, passing through a few helpful parameters
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

    ' Toggle Yellow Locations
    If TogglePotentialLocations(CallingShape) Then
        Exit Sub
    End If
    'MsgBox Curr_Row
    'MsgBox Curr_Col
    
    
    ' This sub will create all of the yellow shapes for a Pawn indicating a potential move location

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Row, UL_Col As Integer ' Upper Left Row, Upper Left Column
    Dim hasMoved As Boolean
    Dim F1, F2 As Boolean ' Used to represent blocked space for Forward 1 and Forward 2
    Dim DL, DR As Boolean ' Used to represent available attack for Diagonal Left and Diagonal Right
    
    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    UL_Row = Range(ULRS).Row ' Getting the top-most row of the board
    UL_Col = Range(ULRS).Column ' Getting the left-most column of the board
    
    hasMoved = HasMovedFromStart(CallingShape.ID) ' A boolean for whether the pawn has moved since she start
    
    ' Adjustmant depending on which direction 'forward' is
    i = -1
    If IsBlack(CallingShape) Then
        i = 1
    End If
    
    ' Loop through the shapes looking for a match
    For Each shp In ActiveSheet.Shapes
        'If shp.Left = CallingShape.Left And shp.Top <> CallingShape.Top Then
        If Curr_Row + i = GetCellRow(shp.Top) And CallingShape.Left = shp.Left Then ' is in front of the pawn, block
            F1 = True
            F2 = True
        ElseIf (Curr_Row + (2 * i) = GetCellRow(shp.Top) And CallingShape.Left = shp.Left) Or hasMoved Then ' is 2 in front of the pawn, block
            F2 = True
        ElseIf Curr_Row + i = GetCellRow(shp.Top) And Curr_Col + i = GetCellCol(shp.Left) And IsBlack(shp) <> IsBlack(CallingShape) Then ' opposing to the diagonal left, attack
            DL = True
        ElseIf Curr_Row + i = GetCellRow(shp.Top) And Curr_Col - i = GetCellCol(shp.Left) And IsBlack(shp) <> IsBlack(CallingShape) Then ' opposing to the diagonal right, attack
            DR = True
        End If
    Next shp
    
    ' Add yellows depending on availability
    
    ' TODO: Figure out cells for yellows
    ' TODO: Create Cell to top, left, conversion
    
    If Not F1 Then
                
        Call AddYellowLocation(Cells(Curr_Row, i + UL_Col - 1).Left, CallingShape.Top, Cells(Curr_Row, i).Width, _
        Cells(Curr_Row, i).Height, CallingShape.ID) ' Create a shape at the available location
        
    ElseIf Not F2 Then
        
    ElseIf DL Then
    
    ElseIf DR Then
    
    End If
    

End Sub
Sub Rook(ByVal Curr_Row, Curr_Col, CallingShape)

    ' This sub will create all of the yellow shapes for a Rook indicating a potential move location
    
    ' Toggle Yellow Locations
    If TogglePotentialLocations(CallingShape) Then
        Exit Sub
    End If

    Dim ULRS As String ' Upper Left Row String
    Dim BR As Range ' Board Range
    Dim UL_Row, UL_Col As Integer ' Upper Left Row, Upper Left Column
    Dim LeB, RiB, UpB, LoB As Integer ' Left, Right, Upper, and Lower bound
    
    ' Set up variables
    ULRS = CStr(ActiveSheet.Cells(1, 1)) 'Getting the board range from A1

    Set BR = Range(ULRS).Resize(8, 8) ' Setting a range object for the board size
    
    UL_Row = Range(ULRS).Row ' Getting the top-most row of the board
    UL_Col = Range(ULRS).Column ' Getting the left-most column of the board
    
    ' indicates the bound of how far the rook can move from its current position
    LeB = 0 ' Moving left from rook
    RiB = 9 ' Moving right from rook
    UpB = 0 ' Moving up from rook
    LoB = 9 ' Moving down (lower) from rook
    
    
    
    ' This loop will find all of the bounds
    ' it will do this by looping through all of the shapes on the sheet, and checking parameters against them
    i = 0
    For Each shp In ActiveSheet.Shapes
        If shp.Left = CallingShape.Left And shp.Top <> CallingShape.Top Then ' If the current shape is in the same column and is not the calling shape
            
            ivert = GetCellRow(shp.Top) ' Get the row of the shape relative to the board
            
            If IsBlack(shp) <> IsBlack(CallingShape) Then ' If the shape is an opposing piece, make it takeable
                i = 1
            End If
        
            If ivert < Curr_Row And ivert > UpB Then ' If the shape is above the rook and is closer than the previous bound
        
                UpB = ivert - i
            
            End If
        
            If ivert > Curr_Row And ivert < LoB Then ' If the shape is below the rook and is closer than the previous bound
            
                LoB = ivert + i
            
            End If
            
            i = 0 ' Reset our takeable token
        End If
        
        If shp.Top = CallingShape.Top And shp.Left <> CallingShape.Left Then ' If the current shape is in the same row and is not the calling shape
            
            ivert = GetCellCol(shp.Left) ' Get the column of the shape relative to the board
            
            If IsBlack(shp) <> IsBlack(CallingShape) Then ' If the shape is an opposing piece, make it takeable
                i = 1
            End If
        
            If ivert < Curr_Col And ivert > LeB Then ' If the shape is to the left of the rook and is closer than the previous bound
        
                LeB = ivert - i
            
            End If
        
            If ivert > Curr_Col And ivert < RiB Then ' If the shape is to the right of the rook and is closer than the previous bound
            
                RiB = ivert + i
            
            End If
            
            i = 0 ' Reset our takeable token
        End If
        
    Next shp
    
    ' These next two loops will go about creating the potential location shapes based on the bounds calculated above
    
    
    For i = UpB + 1 To LoB - 1 ' Our bounds calculate where the bounding shapes are, not the farthest available square, so that needs accounted for
        If i <> Curr_Row Then ' Don't create a potential location ontop of our current piece
        
            Call AddYellowLocation(CallingShape.Left, Cells(i + UL_Row - 1, Curr_Col).Top, Cells(i, Curr_Col).Width, _
            Cells(i, Curr_Col).Height, CallingShape.ID) ' Create a shape at the available location
            
        End If
    Next
    
    For i = LeB + 1 To RiB - 1 ' Our bounds calculate where the bounding shapes are, not the farthest available square, so that needs accounted for
        If i <> Curr_Col Then ' Don't create a potential location ontop of our current piece
                
            Call AddYellowLocation(Cells(Curr_Row, i + UL_Col - 1).Left, CallingShape.Top, Cells(Curr_Row, i).Width, _
            Cells(Curr_Row, i).Height, CallingShape.ID) ' Create a shape at the available location
            
        End If
    Next
  
    
Range("A1").Select ' This is just to move the selection box our of the way of the board when we're finished
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

Sub AddYellowLocation(ByVal Yleft, Ytop, Ywidth, Yheight, Yid)

    ActiveSheet.Shapes.AddShape(msoShapeRectangle, Yleft, Ytop, Ywidth, Yheight).Select ' Create a shape at the available location
                
    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Yid ' Write the calling shape's id into the shape for future reference
            
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 1 ' Make the font as small as possible to be hopefully near invisible
            
    Selection.ShapeRange.Fill.Visible = msoFalse ' Making the color of the shape transparent so the colors of the board will show through
                
    Selection.OnAction = "MovePiece" ' Set shapes to be able to move the piece
                
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 255, 0) ' Give it a yellow border

End Sub

Sub MovePiece()

    Dim CallingShape As Shape
    Set CallingShape = ActiveSheet.Shapes(Application.Caller)
    
    For Each shp In ActiveSheet.Shapes
        If shp.ID = CallingShape.TextFrame2.TextRange.Characters.Text Then
            
            CallingShape.OnAction = shp.OnAction
            
            CallingShape.Line.ForeColor.RGB = shp.Line.ForeColor.RGB
            
            CallingShape.TextFrame2.TextRange.Characters.Text = shp.TextFrame2.TextRange.Characters.Text
            
            CallingShape.TextFrame2.TextRange.Font.Size = shp.TextFrame2.TextRange.Font.Size
            
            CallingShape.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = shp.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB
            
            With CallingShape.TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                .HorizontalAnchor = msoAnchorCenter
            End With
            
            shp.Delete
            
        ElseIf shp.Line.ForeColor.RGB = RGB(255, 255, 0) Then
            shp.Delete
        ElseIf shp.Left = CallingShape.Left And shp.Top = CallingShape.Top And shp.ID <> CallingShape.ID Then
            shp.Delete
        End If
    Next shp

End Sub

Sub buffer()






































































End Sub
