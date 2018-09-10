'The core of the domain terms visualization tool
'Tool author: Bennett Kankuzi
'PhD Student in Computer Science, School of Computing, University of Eastern Finland (Joensuu Campus)
'Email: bfkankuzi@gmail.com
'Supervised by: Prof. Jorma Sajaniemi
'Email: saja@cs.uef.fi
'Date: 10th June 2014
'revised in 2017 and 2018

Option Explicit
Option Compare Text

Dim State As Integer

Public Function ReturnState() As Integer
    'State = 1
    ReturnState = State
End Function


Function ColNo2ColRef(ColNo As Integer) As String
    If ColNo < 1 Or ColNo > 256 Then
        ColNo2ColRef = "#VALUE!"
        Exit Function
    End If
    ColNo2ColRef = Cells(1, ColNo).Address(True, False, xlA1)
    ColNo2ColRef = Left(ColNo2ColRef, InStr(1, ColNo2ColRef, "$") - 1)
End Function


Function ColRef2ColNo(ColRef As String) As Integer
    ColRef2ColNo = 0
    On Error Resume Next
    ColRef2ColNo = Range(ColRef & "1").Column
End Function

Function HasDirectDependents(ByVal target As Excel.Range) As Boolean
    On Error Resume Next
        HasDirectDependents = target.DirectDependents.Count
End Function
Function HasDirectPrecedents(ByVal target As Excel.Range) As Boolean
    On Error Resume Next
        HasDirectPrecedents = target.DirectPrecedents.Count
End Function

Function RowHeader(rowNumber As Integer, colNumber As Integer) As String
    Dim counter As Integer
    Dim testCounter As Integer
    Dim myCell As Range
    Dim myCell2 As Range
    
    Dim numberLabel As Boolean
    RowHeader = "---"
    For counter = colNumber To 1 Step -1
        
        'If Not IsEmpty(Cells(rowNumber, counter)) Then
                    'If Application.IsText(Cells(rowNumber, counter)) Then
                        'MsgBox "Text is row --" & Cells(rowNumber, counter).Text
                        'RowHeader = Cells(rowNumber, counter).Text
                        ''Cells(rowNumber, counter).Interior.ColorIndex = 20
                        'Exit For
                        
                    'End If
        'End If
        
        If Not IsEmpty(Cells(rowNumber, counter)) Then
        
                    numberLabel = False
                    Set myCell2 = Nothing
                                
                                
                                
                                
                    'original text test string
                    'If Application.IsText(Cells(rowNumber, counter)) Or Cells(rowNumber, counter).NumberFormat = "@" Or numberLabel = True Then
                     If Not (HasDirectPrecedents(Cells(rowNumber, counter))) And Not (HasDirectDependents(Cells(rowNumber, counter))) Then
                        
                        RowHeader = Cells(rowNumber, counter).Text
                        
                        
                        If counter > 1 Then
                            
                            testCounter = counter - 1
                            
                             For testCounter = counter - 1 To 1 Step -1
                                'numberLabel = False
                                'If Cells(rowNumber, testCounter).HasFormula Then
                                   'MsgBox "has formula"
                                   'If Cells(rowNumber, testCounter).DirectPrecedents.Count = 1 Then
                                       'MsgBox "has one precedent"
                                       'For Each myCell In Cells(rowNumber, testCounter).DirectPrecedents.Cells
                                           'MsgBox myCell.Address(True, True)
                                           'If myCell.Address(True, True) = Cells(rowNumber - 1, testCounter).Address(True, True) Then
                                               'MsgBox "am a number label" & Cells(rowNumber, testCounter).Address(False, False)
                                               'RowHeader = Cells(rowNumber, testCounter).Value
                                               'numberLabel = True
                                           'End If
                                       'Next
                                       
                                   'End If
                                 'End If
                              
                              
                              
                              
                              
                              'Or IsDate(Cells(rowNumber, testCounter)) removed Or
                              
                              ' Original test text string
                              'If Application.IsText(Cells(rowNumber, testCounter)) Or Cells(rowNumber, testCounter).NumberFormat = "@" Or numberLabel = True Then
                               If Not (HasDirectPrecedents(Cells(rowNumber, testCounter))) And Not (HasDirectDependents(Cells(rowNumber, testCounter))) And Not IsEmpty(Cells(rowNumber, testCounter)) Then
                                If Trim(Cells(rowNumber, testCounter).Text) = "-" Then
                                    'MsgBox "wawwa"
                                Else
                                    RowHeader = Cells(rowNumber, testCounter).Text & " ~ " & RowHeader
                                End If
                                    'MsgBox RowHeader & " " & Cells(rowNumber, testCounter).Address(True, True)
                                    'If RowHeader = " " Then
                                        'MsgBox "am empty"
                                    'End If
                                'End If
                              Else
                                Exit For
                              End If
                              
                              
                             Next
                             
                             
                             RowHeader = RowHeader
                           'Exit For
                        End If
                        
                        
                        
                        Exit For
                        
                    End If
        End If
    Next
    RowHeader = RowHeader
End Function


Function colHeader(rowNumber As Integer, colNumber As Integer) As String
    Dim counter As Integer
    Dim testCounter As Integer
    Dim myCell As Range
    Dim myCell2 As Range
    Dim numberLabel As Integer
    colHeader = "---"
    
    For counter = rowNumber To 1 Step -1
        
        If Not IsEmpty(Cells(counter, colNumber)) Then
        
        
                    numberLabel = False
                                
                   
                    'original text cell testing
                    'If Application.IsText(Cells(counter, colNumber)) Or Cells(counter, colNumber).NumberFormat = "@" Or numberLabel = True Then
                        
                    
                    If Not (HasDirectPrecedents(Cells(counter, colNumber))) And Not (HasDirectDependents(Cells(counter, colNumber))) Then
                     
                     'MsgBox Cells(counter, colNumber).DirectPrecedents.Count
                     'MsgBox Cells(counter, colNumber).Text
                     
                        colHeader = Cells(counter, colNumber).Text
                        
                        If counter > 1 Then
                            
                            testCounter = counter - 1
                            
                             For testCounter = counter - 1 To 1 Step -1
                             
                              'If Cells(testCounter, colNumber).HasFormula Then
                                'MsgBox "has formula"
                                'If Cells(testCounter, colNumber).DirectPrecedents.Count = 1 Then
                                    'MsgBox "has one precedent"
                                    'For Each myCell In Cells(testCounter, colNumber).DirectPrecedents.Cells
                                        'MsgBox myCell.Address(True, True)
                                        'If myCell.Address(True, True) = Cells(testCounter - 1, colNumber) Then
                                            'MsgBox "am a number label" & Cells(testCounter, colNumber).Address(False, False)
                                        'End If
                                    'Next
                                    
                                'End If
                              'End If
                             
                              'Or IsDate(Cells(testCounter, colNumber)) removed
                              
                              'original text cell testing
                             'If Application.IsText(Cells(testCounter, colNumber)) Or Cells(testCounter, colNumber).NumberFormat = "@" Then
                             If Not (HasDirectPrecedents(Cells(testCounter, colNumber))) And Not (HasDirectDependents(Cells(testCounter, colNumber))) And Not IsEmpty(Cells(testCounter, colNumber)) Then
                                If Trim(Cells(testCounter, colNumber).Text) = "-" Then
                                    'Do Nothing
                                    'MsgBox Cells(testCounter, colNumber).Address(True, True)
                                    'colHeader = colHeader
                                Else
                                    'MsgBox Cells(testCounter, colNumber).Text
                                    colHeader = Cells(testCounter, colNumber).Text & " ~ " & colHeader
                                    
                                End If
                              Else
                                Exit For
                              End If
                              
                              
                             Next
                             
                             
                             colHeader = colHeader
                           'Exit For
                        End If
                        
                        
                        
                        Exit For
                        
                    End If
        End If
    
    Next
    'colHeader = colHeader
End Function

Public Sub DomainVisualizeAllColumnToRow()

    'visualizeType = 2 if its from column to row
    State = 2
    DomainVisualizeAllFormulaCells (State)
    
End Sub

Public Sub CallToDomainVisualizeSingleFormulaCellMafikeng()

    
    'visualizeType = 2 if its from column to row
    State = 2
    'DomainVisualizeAllFormulaCells (State)
    
         'MsgBox "hhhh"
         'MsgBox Application.ActiveCell.Cells.Address
        'DomainVisualizeSingleFormulaCellMafikeng (State)
        
    
End Sub

Sub DomainVisualizeAllNaturalExpressions()

    DomainVisualizeAllFormulaCellsNaturalLanguageExpressions (1)
    
End Sub

Sub DomainVisualizeAllRowToColumn()
    'visualizeType = 1 if its from row to column
    State = 2
    DomainVisualizeAllFormulaCells (State)
    
End Sub

Sub DomainVisualizeAllFormulaCells(visualizeType As Integer)

'Application.ScreenUpdating = False

'visualizeType = 1 if its from row to column
'visualizeType = 2 if its from column to row

Dim oWS As Worksheet
Dim oCell As Range
Dim cellAdd As String
Dim cellAddTokens() As String
Dim rowNumber As Integer
Dim colNumber As Integer

Set oWS = ActiveSheet

On Error Resume Next
 For Each oCell In oWS.Cells.SpecialCells(xlCellTypeFormulas)

    'oCell.Interior.ColorIndex = 36
    'MsgBox oCell.Formula
    cellAdd = oCell.Offset(0, 0).Address(True, True)
    
    cellAddTokens() = Split(cellAdd, "$")
    
    rowNumber = CInt(cellAddTokens(2))
    
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = oCell.Formula
    cellFormula = Replace(cellFormula, "=", "", 1, 1)
    cellFormula = Replace(cellFormula, "$", "")
    
    'MsgBox cellFormula
    
    
    
    'go column-wise
    Dim colHeaderString As String
    'colHeaderString = colHeader(rowNumber, colNumber)
    'MsgBox "column header " & colHeaderString
    
    'go row-wise
    Dim rowHeaderString As String
    'rowHeaderString = RowHeader(rowNumber, colNumber)
    'MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range, myCell As Range, myCell2 As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    Dim tempRow As Integer
    Dim tempCol As Integer
    
    Dim numberLabel As Boolean
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            
            'MsgBox "cel  " & cel.Address(False, False)
            numberLabel = False 'in cases like fact(n)
            
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            
            
            If numberLabel = False Then
                cellAdd = cel.Address(True, True)
                
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
            End If
            If visualizeType = 1 Then
            cellHeader = rowHeaderString & " | " & colHeaderString
            Else
                If visualizeType = 2 Then
                    cellHeader = colHeaderString & " | " & rowHeaderString
                End If
                
            End If
            
            cellFormula = Replace(cellFormula, cel.Address(False, False), " " & cellHeader & " ")
            cellFormula = Replace(cellFormula, ":", " ... ")
            
            'cellFormula = Replace(cellFormula, "=", "")
            
            cellFormula = Replace(cellFormula, "--- |", "")
            cellFormula = Replace(cellFormula, "| ---", "")
            cellFormula = Replace(cellFormula, "---", "unnamed")
            
            cellFormula = Trim(cellFormula)
            
            'MsgBox cellHeader
            'MsgBox cellAdd
            'Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependencies found."
    End If
    
    
    oCell.ClearComments
    oCell.AddComment cellFormula
    
    oCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    oCell.Borders(xlEdgeRight).ColorIndex = 26
    oCell.Borders(xlEdgeRight).Weight = xlThick
    
    If oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous Then
     Else
        oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
    End If
    
    With oCell.Comment
        '.Shape.TextFrame.AutoSize = True
        .Shape.Fill.ForeColor.SchemeColor = 7
        .Shape.TextFrame.Characters.Font.Name = "Arial"
        .Shape.TextFrame.Characters.Font.Size = 14
        '.Shape.TextFrame.Characters.Font.Bold = True
        '.Shape.TextEffect.FontBold = msoTrue

        '.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Shape.TextFrame2.WordWrap = msoTrue
        .Shape.Top = .Parent.Top + 18
        
        Dim CommentSize As Long
        .Shape.TextFrame.AutoSize = True
        If .Shape.Width > 250 Then
            CommentSize = .Shape.Width * .Shape.Height
            .Shape.Width = 250
            .Shape.Height = (CommentSize / 180) * 1.2
        End If

    End With
    
    
    
Next oCell

'Application.ScreenUpdating = True
End Sub

Sub DomainVisualizeAllFormulaCellsNaturalLanguageExpressions(visualizeType As Integer)

'Application.ScreenUpdating = False

'visualizeType = 1 if its from row to column
'visualizeType = 2 if its from column to row

Dim oWS As Worksheet
Dim oCell As Range
Dim cellAdd As String
Dim cellAddTokens() As String
Dim rowNumber As Integer
Dim colNumber As Integer


Set oWS = ActiveSheet

On Error Resume Next
 For Each oCell In oWS.Cells.SpecialCells(xlCellTypeFormulas)

    'oCell.Interior.ColorIndex = 36
    'MsgBox oCell.Formula
    cellAdd = oCell.Offset(0, 0).Address(True, True)
    
    cellAddTokens() = Split(cellAdd, "$")
    
    rowNumber = CInt(cellAddTokens(2))
    
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = oCell.Formula
    cellFormula = Replace(cellFormula, "=", "", 1, 1)
    cellFormula = Replace(cellFormula, "$", "")
    
    'MsgBox cellFormula
    
    
    
    'go column-wise
    Dim colHeaderString As String
    'colHeaderString = colHeader(rowNumber, colNumber)
    'MsgBox "column header " & colHeaderString
    
    'go row-wise
    Dim rowHeaderString As String
    'rowHeaderString = RowHeader(rowNumber, colNumber)
    'MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range, myCell As Range, myCell2 As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    Dim tempRow As Integer
    Dim tempCol As Integer
    
    Dim numberLabel As Boolean
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            
            'MsgBox "cel  " & cel.Address(False, False)
            numberLabel = False 'in cases like fact(n)
            
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            
            
            If numberLabel = False Then
                cellAdd = cel.Address(True, True)
                
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
            End If
            If visualizeType = 1 Then
            cellHeader = rowHeaderString & " for " & colHeaderString
            Else
                If visualizeType = 2 Then
                    cellHeader = colHeaderString & " for " & rowHeaderString
                End If
                
            End If
            
            cellFormula = Replace(cellFormula, cel.Address(False, False), " " & cellHeader & " ")
            cellFormula = Replace(cellFormula, ":", " ... ")
            
            'cellFormula = Replace(cellFormula, "=", "")
            
            cellFormula = Replace(cellFormula, "--- for ", "")
            cellFormula = Replace(cellFormula, "for ---", "")
            cellFormula = Replace(cellFormula, "---", "unnamed")
            
            cellFormula = Trim(cellFormula)
            
            'MsgBox cellHeader
            'MsgBox cellAdd
            'Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependencies found."
    End If
    
    
    cellFormula = TranslateFormula(cellFormula) 'mafikeng
    'MsgBox "Equivalent Formula is " & cellFormula
    oCell.ClearComments
    oCell.AddComment cellFormula
    
    oCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    oCell.Borders(xlEdgeRight).ColorIndex = 26
    oCell.Borders(xlEdgeRight).Weight = xlThick
    
    If oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous Then
     Else
        oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
    End If
    
    'mafikeng
    'UserForm1.Label1 = cellFormula
    'UserForm1.Show
    'UserForm1.Left = True
    
    'mafikeng
    
    With oCell.Comment
        '.Shape.TextFrame.AutoSize = True
        .Shape.Fill.ForeColor.SchemeColor = 7
        .Shape.TextFrame.Characters.Font.Name = "Arial"
        .Shape.TextFrame.Characters.Font.Size = 14
        '.Shape.TextFrame.Characters.Font.Bold = True
        '.Shape.TextEffect.FontBold = msoTrue

        '.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Shape.TextFrame2.WordWrap = msoTrue
        .Shape.Top = .Parent.Top + 18
        
        Dim CommentSize As Long
        .Shape.TextFrame.AutoSize = True
        If .Shape.Width > 250 Then
            CommentSize = .Shape.Width * .Shape.Height
            .Shape.Width = 250
            .Shape.Height = (CommentSize / 180) * 1.2
        End If
        
        Dim pos As Integer
        Dim startPos As Integer
        With .Shape.TextFrame
        
            If InStr(1, cellFormula, "if ", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "if ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "if ", 1), 3).Font
                            
                            '.ColorIndex = 7
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                            
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "otherwise", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "otherwise", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "otherwise", 1), 9).Font
                            
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            
            If InStr(1, cellFormula, "sum", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "sum", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "sum", 1), 3).Font
                            
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                            
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            
            If InStr(1, cellFormula, "average", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "average", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "average", 1), 7).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "plus", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "plus", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "plus", 1), 4).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "minus", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "minus", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "minus", 1), 5).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "multiply by", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "multiply by", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "multiply by", 1), 11).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
        End With
        
    End With
    
    
    
Next oCell

'Application.ScreenUpdating = True
End Sub

'mafikeng code
Sub DomainVisualizeSingleFormulaCellMafikeng(visualizeType As Integer)

'Application.ScreenUpdating = False

'visualizeType = 1 if its from row to column
'visualizeType = 2 if its from column to row

Dim oWS As Worksheet
Dim oCell As Range
Dim cellAdd As String
Dim cellAddTokens() As String
Dim rowNumber As Integer
Dim colNumber As Integer
Dim ActiveCellAddress As String


Set oWS = ActiveSheet
On Error Resume Next
' mafikeng original For Each oCell In oWS.Cells.SpecialCells(xlCellTypeFormulas)
    
'' mafikeng original For Each oCell In ActiveCell.Cells.Range

    'oCell.Interior.ColorIndex = 36
    'MsgBox oCell.Formula
    
    ' Set oCell = Range(ActiveCellAddress)
    ' cellAdd = oCell.Offset(0, 0).Address(True, True)
    
    'mafikeng code just inputting the active address cell
    Set oCell = Application.ActiveCell.Cells
    cellAdd = Application.ActiveCell.Cells.Address
    
    cellAddTokens() = Split(cellAdd, "$")
    
    rowNumber = CInt(cellAddTokens(2))
    
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = oCell.Formula
    'cellFormula = Replace(cellFormula, "=", "", 1, 1)
    cellFormula = Replace(cellFormula, "$", "")
    
    'MsgBox cellFormula
    
    
    
    'go column-wise
    Dim colHeaderString As String
    'colHeaderString = colHeader(rowNumber, colNumber)
    'MsgBox "column header " & colHeaderString
    
    'go row-wise
    Dim rowHeaderString As String
    'rowHeaderString = RowHeader(rowNumber, colNumber)
    'MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range, myCell As Range, myCell2 As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    Dim tempRow As Integer
    Dim tempCol As Integer
    
    Dim numberLabel As Boolean
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        If (Cells(rowNumber, colNumber).DirectPrecedents.Cells.Count > 100) Or IsEmpty(oCell) Then
            'MsgBox Cells(rowNumber, colNumber).DirectPrecedents.Cells.Count
            'MsgBox "More than 300 referenced cells in the formula ... Cannot translate formula."
           Exit Sub
        End If
        'mafikeng
        
        
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            
            'MsgBox "cel  " & cel.Address(False, False)
            numberLabel = False 'in cases like fact(n)
            
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            
            
            If numberLabel = False Then
                cellAdd = cel.Address(True, True)
                
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
            End If
            If visualizeType = 1 Then
            cellHeader = rowHeaderString & " for " & colHeaderString
            Else
                If visualizeType = 2 Then
                    cellHeader = colHeaderString & " for " & rowHeaderString
                End If
                
            End If
            
            cellFormula = Replace(cellFormula, cel.Address(False, False), " " & cellHeader & " ")
            'cellFormula = Replace(cellFormula, ":", " ... ")
            
            'cellFormula = Replace(cellFormula, "=", "")
            
            cellFormula = Replace(cellFormula, "--- for", "")
            cellFormula = Replace(cellFormula, "for ---", "")
            cellFormula = Replace(cellFormula, "---", "unnamed")
            cellFormula = Trim(cellFormula)
            
            'MsgBox cellHeader
            'MsgBox cellAdd
            'Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependencies found."
    End If
    
    'MsgBox "Equivalent Formula is " & cellFormula
    cellFormula = TranslateFormula(cellFormula) 'mafikeng
    'MsgBox "Equivalent Formula is " & cellFormula
    oCell.ClearComments
    oCell.AddComment cellFormula
    
    oCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    oCell.Borders(xlEdgeRight).ColorIndex = 26
    oCell.Borders(xlEdgeRight).Weight = xlThick
    
    If oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous Then
     Else
        oCell.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
    End If
    
    'mafikeng
    'UserForm1.Label1 = cellFormula
    'UserForm1.Show
    'UserForm1.Left = True
    
    'mafikeng
    
    With oCell.Comment
        '.Shape.TextFrame.AutoSize = True
        .Shape.Fill.ForeColor.SchemeColor = 7
        .Shape.TextFrame.Characters.Font.Name = "Arial"
        .Shape.TextFrame.Characters.Font.Size = 14
        '.Shape.TextFrame.Characters.Font.Bold = True
        '.Shape.TextEffect.FontBold = msoTrue

        '.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Shape.TextFrame2.WordWrap = msoTrue
        .Shape.Top = .Parent.Top + 18
        
        Dim CommentSize As Long
        .Shape.TextFrame.AutoSize = True
        If .Shape.Width > 250 Then
            CommentSize = .Shape.Width * .Shape.Height
            .Shape.Width = 250
            .Shape.Height = (CommentSize / 180) * 1.2
        End If
        
        Dim pos As Integer
        Dim startPos As Integer
        With .Shape.TextFrame
        
            If InStr(1, cellFormula, "if ", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "if ", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "if ", 1), 3).Font
                            
                            '.ColorIndex = 7
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                            
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "otherwise", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "otherwise", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "otherwise", 1), 9).Font
                            
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            
            If InStr(1, cellFormula, "sum", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "sum", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "sum", 1), 3).Font
                            
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                            
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            
            If InStr(1, cellFormula, "average", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "average", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "average", 1), 7).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "plus", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "plus", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "plus", 1), 4).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
            If InStr(1, cellFormula, "minus", 1) > 0 Then
                startPos = 1
                Do
                    pos = InStr(startPos, cellFormula, "plus", 1)
                    If pos > 0 Then
                        With .Characters(InStr(startPos, cellFormula, "minus", 1), 4).Font
                            
                            '.ColorIndex = 6
                            .color = RGB(255, 255, 0) ' yellow
                            .Italic = True
                        End With
                    End If
                    
                    startPos = pos + 1
                Loop While (pos > 0)
            End If
            
        End With
        
    End With
    
    
    
' mafikeng end of Next oCell

'Application.ScreenUpdating = True
End Sub

Public Sub DomainVisualizeSingleFormulaCell()

  Dim rCell As Range
    Dim rRng As Range
    Dim strippedCellAddressString As String
    Dim CellAddressString As String
    Dim cellAddressTokens() As String
    Dim activeCellAddressString As String
    Dim resultRange As String
    Dim precedentRange As Range
    Dim precedentCell As Range
    Dim LRandomNumber As Integer
    
    If ActiveCell.HasFormula Then
    
        Dim cellAdd As String
        cellAdd = ActiveCell.Offset(0, 0).Address(True, True)
        Dim cellAddTokens() As String
        cellAddTokens() = Split(cellAdd, "$")
        Dim rowNumber As Integer
        rowNumber = CInt(cellAddTokens(2))
        Dim colNumber As Integer
        colNumber = ColRef2ColNo(cellAddTokens(1))
        
        Dim cellFormula As String
        cellFormula = ActiveCell.Formula
        
        'MsgBox cellFormula
        
        
        
        'go column-wise
        Dim colHeaderString As String
        'colHeaderString = colHeader(rowNumber, colNumber)
        'MsgBox "column header " & colHeaderString
        
        'go row-wise
        Dim rowHeaderString As String
        'rowHeaderString = RowHeader(rowNumber, colNumber)
        'MsgBox "row header" & rowHeaderString
        Dim precedentsRange As Range, cel As Range
        
        Dim directPrecedentsString As String
        Dim cellHeader As String
        
        
        On Error Resume Next
        If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
            Cells(rowNumber, colNumber).DirectPrecedents.Cells
            For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
                'MsgBox "cel  " & cel.Address(False, False)
                cellAdd = cel.Address(True, True)
                cellAddTokens() = Split(cellAdd, "$")
                rowNumber = CInt(cellAddTokens(2))
                colNumber = ColRef2ColNo(cellAddTokens(1))
                colHeaderString = colHeader(rowNumber, colNumber)
                'MsgBox "column header " & colHeaderString
                rowHeaderString = RowHeader(rowNumber, colNumber)
                'MsgBox "row header" & rowHeaderString
                cellHeader = rowHeaderString & " \ " & colHeaderString
                cellFormula = Replace(cellFormula, cel.Address(False, False), cellHeader)
                cellFormula = Replace(cellFormula, ":", " to ")
                
                'MsgBox cellFormula
                'MsgBox cellAdd
                'Range(cel.Address(False, False)).Interior.color = 20
                directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
            Next
            
            'MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
        Else
            MsgBox "No dependencies found."
        End If
        
        'MsgBox "Equivalent Formula is " & cellFormula
        ActiveCell.ClearComments
        ActiveCell.AddComment cellFormula
        With ActiveCell.Comment
            .Shape.TextFrame.AutoSize = True
            
        End With
    Else
    
        MsgBox "Cell has no formula"
        
    End If
    
    
End Sub

Sub SelectedRange()


    Dim rCell As Range
    Dim rRng As Range
    Dim strippedCellAddressString As String
    Dim CellAddressString As String
    Dim cellAddressTokens() As String
    Dim activeCellAddressString As String
    Dim resultRange As String
    Dim precedentRange As Range
    Dim precedentCell As Range
    

    Dim LRandomNumber As Integer
    
    Dim cellAdd As String
    cellAdd = ActiveCell.Offset(0, 0).Address(True, True)
    Dim cellAddTokens() As String
    cellAddTokens() = Split(cellAdd, "$")
    Dim rowNumber As Integer
    rowNumber = CInt(cellAddTokens(2))
    Dim colNumber As Integer
    colNumber = ColRef2ColNo(cellAddTokens(1))
    
    Dim cellFormula As String
    cellFormula = ActiveCell.Formula
    
    MsgBox cellFormula
    
    
    
    'go column-wise
    Dim colHeaderString As String
    colHeaderString = colHeader(rowNumber, colNumber)
    MsgBox "column header " & colHeaderString
    
    'go row-wise
    Dim rowHeaderString As String
    rowHeaderString = RowHeader(rowNumber, colNumber)
    MsgBox "row header" & rowHeaderString
    Dim precedentsRange As Range, cel As Range
    
    Dim directPrecedentsString As String
    Dim cellHeader As String
    
    
    On Error Resume Next
    If Cells(rowNumber, colNumber).DirectPrecedents.Count Then
        Cells(rowNumber, colNumber).DirectPrecedents.Cells
        For Each cel In Cells(rowNumber, colNumber).DirectPrecedents.Cells
            MsgBox "cel  " & cel.Address(False, False)
            cellAdd = cel.Address(True, True)
            cellAddTokens() = Split(cellAdd, "$")
            rowNumber = CInt(cellAddTokens(2))
            colNumber = ColRef2ColNo(cellAddTokens(1))
            colHeaderString = colHeader(rowNumber, colNumber)
            MsgBox "column header " & colHeaderString
            rowHeaderString = RowHeader(rowNumber, colNumber)
            MsgBox "row header" & rowHeaderString
            cellHeader = colHeaderString & rowHeaderString
            cellFormula = Replace(cellFormula, cel.Address(False, False), cellHeader)
            MsgBox cellHeader
            MsgBox cellAdd
            Range(cel.Address(False, False)).Interior.color = 20
            directPrecedentsString = directPrecedentsString & "," & cel.Address(False, False)
        Next
        
        MsgBox Cells(rowNumber, colNumber).Precedents.Count & " dependancies found."
    Else
        MsgBox "No dependancies found."
    End If
    
    MsgBox "Replaced Cell Formula is " & cellFormula
    
    
    
    

    

    
    
    Dim UsedRng As Range
    Dim FirstRow As Long, LastRow As Long, FirstCol As Long, LastCol As Integer
    
    Set UsedRng = ActiveSheet.UsedRange
     
    FirstRow = UsedRng(1).Row
    FirstCol = UsedRng(1).Column
    LastRow = UsedRng(UsedRng.Cells.Count).Row
    LastCol = UsedRng(UsedRng.Cells.Count).Column
     
    'MsgBox "First used row is: " & FirstRow
    'MsgBox "First used column is: " & FirstCol
    'MsgBox "Last used row is: " & LastRow
    'MsgBox "Last used column is: " & LastCol

    'MsgBox Selection.Address(False, False)
    'Set rRange = ActiveSheet.UsedRange.SpecialCells _
     '(xlCellTypeConstants, xlTextValues)
     Set rRange = ActiveSheet.UsedRange.SpecialCells _
     (xlCellTypeFormulas)
     
     'MsgBox "wawaw" & rRange
     
     For Each rCell In rRange

        'rCell.Interior.ColorIndex = 50
        'MsgBox "rrr" & rCell.Address

     Next rCell
     
     
    
    Dim mySelectedRangeString As String
    
    mySelectedRangeString = Selection.Address(False, False)
    
    'MsgBox "rrr" & mySelectedRangeString
    
    Dim mySelectedRange As Range
    
    Set mySelectedRange = ActiveSheet.Range(mySelectedRangeString)
    
    Dim myCell As Range
    Dim color As Integer
    
    Set rRange = mySelectedRange
    
    For Each myCell In rRange
            LRandomNumber = Int((56 - 2 + 1) * Rnd + 2)
            color = LRandomNumber
    
            'Debug.Print rCell.Address, rCell.Value
            'myCell.ShowPrecedents (True)
            'myCell.ShowPrecedents
            
            cellAddressTokens() = Split(myCell.Address, "$")
            'MsgBox (cellAddressTokens(2))
            CellAddressString = myCell.Address
            
            'MsgBox "tta" + CellAddressString
            strippedCellAddressString = Application.WorksheetFunction.Substitute(CellAddressString, "$", "")
        
        
            activeCellAddressString = strippedCellAddressString
            'MsgBox (activeCellAddressString)
            
            
            Dim result1 As Long
            
            'result1 = cellAddressTokens(2) + 1
            
            result1 = FirstRow '4/2/2013
            
            
            'resultRange = ActiveCell + LastRow
            'MsgBox "result 1" & result1
            
            Dim result2 As String
            
            Dim firstCellColumnwise As String
            
            firstCellColumnwise = cellAddressTokens(1) & CStr(result1)
            
            'MsgBox "First cell columnwise " & firstCellColumnwise
            
            Dim lastCellColumnwise As String
            
            'lastCellColumnwise = cellAddressTokens(1) & CStr(LastRow) '04/02/2012
            lastCellColumnwise = cellAddressTokens(1) & cellAddressTokens(2)
            
            'MsgBox "Last cell columnwise " & lastCellColumnwise
            
            Dim myRangeString As String
            
            myRangeString = lastCellColumnwise + ":" + firstCellColumnwise
            
            MsgBox myRangeString
            
            
            Set rRng = ActiveSheet.Range(myRangeString).End(xlUp)
        
            For Each rCell In rRng.Cells
                'Debug.Print rCell.Address, rCell.Value
                rCell.ShowPrecedents (True)
                rCell.ShowPrecedents
                'MsgBox "wawaw" & rCell.Precedents
                'If (rCell.HasFormula) Then
                'Set precedentRange = rCell.DirectPrecedents
                
                '    For Each precedentCell In precedentRange.Cells
                        
                '            precedentCell.Interior.ColorIndex = color
                        
                        'MsgBox "sddd" & precedentCell.Address
                 '   Next precedentCell
                'End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                If Not IsEmpty(rCell) Then
                    If Application.IsText(rCell) Then
                        MsgBox "Text" & rCell.Text
                    End If
                    
                    If IsNumeric(rCell) Then
                        MsgBox "Text" & rCell.Text
                        If rCell.HasFormula Then
                            MsgBox "Formula" & rCell.Text
                        End If
                    End If
                    
                    
                    rCell.Interior.ColorIndex = color
                End If
                
            Next rCell
            
            
        
            Dim firstCellRowwise As String
            
            MsgBox "token" + cellAddressTokens(2)
            
            firstCellRowwise = cellAddressTokens(1) & CStr(result1)
            
            Dim columnNumber As Integer
            
            columnNumber = ColRef2ColNo(cellAddressTokens(1))
            
            'columnNumber = columnNumber + 1 ' 04/02/2012
            
            'columnNumber = columnNumber + 1 '04/02/2012
            
            Dim columnNumberString As String
            
            'columnNumberString = ColNo2ColRef(columnNumber)'04/02/2012
            
            
            'firstCellRowwise = columnNumberString & cellAddressTokens(2)'04/02/2012
            
            firstCellRowwise = "A" & cellAddressTokens(2)
            
            MsgBox "First cell rowwise " & firstCellRowwise
            
            
            Dim lastCellRowwise As String
            
            'lastCellRowwise = CStr(ColNo2ColRef(LastCol)) & cellAddressTokens(2) '04/02/2012
            
            lastCellRowwise = cellAddressTokens(1) & cellAddressTokens(2)
            
            'MsgBox "Last cell rowwise " & lastCellRowwise
            
            Dim myRowRangeString As String
            
            myRowRangeString = firstCellRowwise + ":" + lastCellRowwise
            
            MsgBox "Row" & myRowRangeString
            
            'MsgBox ColNo2ColRef(LastCol)
            
            Set rRng = ActiveSheet.Range(myRowRangeString)
        
            For Each rCell In rRng.Cells
                'Debug.Print rCell.Address, rCell.Value
                'rCell.ShowPrecedents (True)
                'rCell.ShowPrecedents
                If Not IsEmpty(rCell) Then
                
                    If Application.IsText(rCell) Then
                        MsgBox "row" & rCell.Text
                    End If
                    
                    If IsNumeric(rCell) Then
                        MsgBox "row" & rCell.Text
                        If rCell.HasFormula Then
                            MsgBox "rowFormula" & rCell.Text
                        End If
                    End If
                    rCell.Interior.ColorIndex = color
                End If
                
                
                

                
                
            Next rCell
                
            
        Next myCell






End Sub


Sub RemoveAllShapes()
    Dim shp As Shape
                
                For Each shp In ActiveSheet.Shapes
                
                    If shp.AutoShapeType = _
                        msoShapeRectangle Then
                        shp.Delete
                    End If
                
                Next
End Sub




Sub Remove_All_Comments_From_Worksheet()

Cells.ClearComments

End Sub

Function TranslateFormula(cellFormula As String) As String

    Dim MyRE As Object
    Dim match As Object
    Dim matches As Object
    Dim subMatch As String
    Dim i As Integer
    Dim v As Variant
    Dim w As Variant
    Dim myFormula As String
    Dim subMatchTokens() As String
    Dim subMatchTokenZero As String
    Dim subMatchTokenTwo As String
    Dim subMatchTokenLast As String
    
    'the following is used to extract column header for current cell in order to use it
    ' in the if statement translation particularly for token 1
    Dim cellAdd As String
    Dim cellAddTokens() As String
    Dim rowNumber As Integer
    Dim colNumber As Integer
    Dim colHeaderString As String
    cellAdd = Application.ActiveCell.Cells.Address(True, True)
    cellAddTokens() = Split(cellAdd, "$")
    rowNumber = CInt(cellAddTokens(2))
    colNumber = ColRef2ColNo(cellAddTokens(1))
    colHeaderString = colHeader(rowNumber, colNumber)
    
    'back to our regex task
    Set MyRE = CreateObject("vbscript.regexp")
    MyRE.ignorecase = True
    MyRE.Global = True
    


    ''MsgBox cellFormula
    'MyRE.Pattern = "(IF\(\s[^,\(\)]+,[^,\(\)]+,[^,\(\)]+\))" 'Pattern IF(argument, argument, argument)
                                                              ' arguments should not contain, ( and )
    v = Array( _
    "(IF\([^\(\)]+\))", _
    "(AVERAGE\([^\(\)]+\))", _
    "(SUM\([^\(\)]+\))")
    
    MyRE.Pattern = Join(v, "|")
    'If MyRE.Test(cellFormula) Then
    While (MyRE.Test(cellFormula))
        ''MsgBox "current cell Formula " & cellFormula
        
        MyRE.Pattern = "(IF\([^\(\)]+\))"
        If MyRE.Test(cellFormula) Then ' for the IF excel function
            Set matches = MyRE.Execute(cellFormula)
            For Each match In matches
                'MsgBox "matches " + match.Value
            
                If match.SubMatches.Count > 0 Then
                    
                    For i = 0 To match.SubMatches.Count - 1
                         
                        ''MsgBox " submatch " & i & " " & match.SubMatches(i)
                        subMatch = match.SubMatches(i)
                        subMatchTokens() = Split(match.SubMatches(i), ",")
                        ''MsgBox " submatch token " & 0 & " " & subMatchTokens(0)
                        ''MsgBox " submatch token " & 1 & " " & subMatchTokens(1)
                        ''MsgBox " submatch token " & 2 & " " & subMatchTokens(2)
                        subMatchTokenZero = subMatchTokens(0)
                        subMatchTokenZero = Replace(subMatchTokenZero, "IF(", "if")
                        subMatch = Replace(subMatch, subMatchTokens(0), subMatchTokenZero & "~ " & colHeaderString & " is ")
                        'MsgBox "," & subMatchTokens(1)
                        subMatch = Replace(subMatch, "," & subMatchTokens(1), subMatchTokens(1))
                        If UBound(subMatchTokens) = 2 Then
                            subMatchTokenTwo = subMatchTokens(2)
                            subMatchTokenTwo = Replace(subMatchTokenTwo, ")", "")
                            subMatch = Replace(subMatch, "," & subMatchTokens(2), "; otherwise " & subMatchTokenTwo)
                        End If
                        
                        subMatch = Replace(subMatch, ",", "~") ' Replace all commas so that we can show that we are done with the subMatch
                        subMatch = Replace(subMatch, ")", "")
                        cellFormula = Replace(cellFormula, match.SubMatches(i), subMatch)
                    Next
                End If ' end match.SubMatches.Count > 0
            Next match
        End If ' end MyRE.Test(cellFormula)
        
        MyRE.Pattern = "(AVERAGE\([^\(\)]+\))"
        If MyRE.Test(cellFormula) Then ' for the AVERAGE excel function
            
            Set matches = MyRE.Execute(cellFormula)
            For Each match In matches
                ''MsgBox "matches average" + match.Value
            
                If match.SubMatches.Count > 0 Then
                    
                    For i = 0 To match.SubMatches.Count - 1
                        
                        ''MsgBox " submatch " & i & " " & match.SubMatches(i)
                        subMatch = match.SubMatches(i)
                        subMatchTokens() = Split(match.SubMatches(i), ",")
                        ''MsgBox " submatch token " & 0 & " " & subMatchTokens(0)
                        
                        If UBound(subMatchTokens) = 0 Then 'if there is one argument only
                            subMatchTokenZero = subMatchTokens(0)
                            subMatchTokenZero = Replace(subMatchTokenZero, "AVERAGE(", "the average of ")
                            subMatchTokenZero = Replace(subMatchTokenZero, ")", "")
                            subMatch = Replace(subMatch, subMatchTokens(0), subMatchTokenZero)
                            cellFormula = Replace(cellFormula, match.SubMatches(i), subMatch)
                        End If
                        
                        If UBound(subMatchTokens) > 0 Then      'if there is more than one argument
                            subMatchTokenZero = subMatchTokens(0)
                            subMatchTokenZero = Replace(subMatchTokenZero, "AVERAGE(", "the average of ")
                            subMatch = Replace(subMatch, subMatchTokens(0), subMatchTokenZero)
                            
                            subMatchTokenLast = subMatchTokens(UBound(subMatchTokens))
                            subMatchTokenLast = Replace(subMatchTokenLast, ")", "")
                            subMatch = Replace(subMatch, subMatchTokens(UBound(subMatchTokens)), " and " & subMatchTokenLast)
                            subMatch = Replace(subMatch, ",", "~") ' Replace all commas so that we can show that we are done with the subMatch
                            cellFormula = Replace(cellFormula, match.SubMatches(i), subMatch)
                        End If
                        
                    Next
                End If ' end match.SubMatches.Count > 0
            Next match
        End If ' end MyRE.Test(cellFormula)
        
        MyRE.Pattern = "(SUM\([^\(\)]+\))"
        If MyRE.Test(cellFormula) Then ' for the SUM excel function
            
            Set matches = MyRE.Execute(cellFormula)
            For Each match In matches
                ''MsgBox "matches average" + match.Value
            
                If match.SubMatches.Count > 0 Then
                    
                    For i = 0 To match.SubMatches.Count - 1
                        
                        ''MsgBox " submatch " & i & " " & match.SubMatches(i)
                        subMatch = match.SubMatches(i)
                        subMatchTokens() = Split(match.SubMatches(i), ",")
                        ''MsgBox " submatch token " & 0 & " " & subMatchTokens(0)
                        
                        If UBound(subMatchTokens) = 0 Then 'if there is one argument only
                            subMatchTokenZero = subMatchTokens(0)
                            subMatchTokenZero = Replace(subMatchTokenZero, "SUM(", "the sum of ")
                            subMatchTokenZero = Replace(subMatchTokenZero, ")", "")
                            subMatch = Replace(subMatch, subMatchTokens(0), subMatchTokenZero)
                            cellFormula = Replace(cellFormula, match.SubMatches(i), subMatch)
                        End If
                        
                        If UBound(subMatchTokens) > 0 Then      'if there is more than one argument
                            subMatchTokenZero = subMatchTokens(0)
                            subMatchTokenZero = Replace(subMatchTokenZero, "SUM(", "the sum of ")
                            subMatch = Replace(subMatch, subMatchTokens(0), subMatchTokenZero)
                            
                            subMatchTokenLast = subMatchTokens(UBound(subMatchTokens))
                            subMatchTokenLast = Replace(subMatchTokenLast, ")", "")
                            subMatch = Replace(subMatch, subMatchTokens(UBound(subMatchTokens)), " and " & subMatchTokenLast)
                            subMatch = Replace(subMatch, ",", "~") ' Replace all commas so that we can show that we are done with the subMatch
                            cellFormula = Replace(cellFormula, match.SubMatches(i), subMatch)
                        End If
                        
                    Next
                End If ' end match.SubMatches.Count > 0
            Next match
        End If ' end MyRE.Test(cellFormula)
        
        MyRE.Pattern = Join(v, "|")
    Wend
    
    If InStr(cellFormula, "+") > 0 Then
        cellFormula = Replace(cellFormula, "+", " plus ")
    End If
    
    If InStr(cellFormula, "-") > 0 Then
        cellFormula = Replace(cellFormula, "-", " minus ")
    End If
    
    If InStr(cellFormula, "*") > 0 Then
        cellFormula = Replace(cellFormula, "*", " multiply by ")
    End If
    
    If InStr(cellFormula, ":") > 0 Then
        cellFormula = Replace(cellFormula, ":", " to ")
    End If
    
    If InStr(cellFormula, "/") > 0 Then
        cellFormula = Replace(cellFormula, "/", " divided by ")
    End If
    
    If InStr(cellFormula, ">") > 0 Then
        cellFormula = Replace(cellFormula, ">", " is greater than ")
    End If
    
    If InStr(cellFormula, "<") > 0 Then
        cellFormula = Replace(cellFormula, "<", " is less than ")
    End If
    
    If InStr(cellFormula, "=") > 0 Then
        cellFormula = Replace(cellFormula, "=", "")
    End If
    
    
    
    
    If InStr(cellFormula, "CONCATENATE") > 0 Then
        cellFormula = Replace(cellFormula, "CONCATENATE(", " the concatenation of ")
    End If
    
    cellFormula = Replace(cellFormula, "~", ",") ' replace our delimiters back with commas
  TranslateFormula = cellFormula
End Function

Sub ExtractWorkBookDataFlowGraph()
    Dim thisWorkBook As Workbook
    Dim currentWorkSheet As Worksheet
    Dim formulaCellsInWorkSheet As Range
    Dim formulaCell As Range
    Dim precedentsRange As Range
    Dim precedentCell As Range
    Dim filePath As String
    Dim thisWorkBookName As String
    Dim textToWriteToFile As String
    Dim precedentsVar As Variant
    Dim qCount As Integer
    Dim shellString As String
    
    
    Set thisWorkBook = ActiveWorkbook
    
    thisWorkBookName = thisWorkBook.Name
    thisWorkBookName = Replace(thisWorkBookName, ".xlsm", "") ' replace extension
    
    'MsgBox thisWorkBook.Name
    'MsgBox thisWorkBook.Path
    
    filePath = thisWorkBook.Path & "\GeneratedGraphs\" & thisWorkBookName & ".dot"
    'MsgBox filePath
    
    Open filePath For Output As #1 ' open file for writing
    Print #1, "digraph G{"
    Print #1, "overlap=false"
    Print #1, "splines=true"
    On Error Resume Next
    For Each currentWorkSheet In thisWorkBook.Worksheets
        
        
        If IsEmpty(currentWorkSheet.Cells.SpecialCells(xlCellTypeFormulas)) Then
            'do nothing
        Else
            'MsgBox currentWorkSheet.Name
            Set formulaCellsInWorkSheet = currentWorkSheet.Cells.SpecialCells(xlCellTypeFormulas)
        
            For Each formulaCell In formulaCellsInWorkSheet
                If InStr(formulaCell.Formula, "!") > 0 Then
                    'MsgBox formulaCell.Formula
                        qCount = 1
                    Do
                        formulaCell.ShowPrecedents
                        'On Error Resume Next
                        precedentsVar = formulaCell.NavigateArrow(True, 1, qCount)
                        
                        'MsgBox CStr(precedentsVar)
                        'MsgBox Selection.Address(True, True)
                        'MsgBox Selection.Parent.Name
                        'MsgBox "formula" & formulaCell.Parent.Name
                        
                        
                        'MsgBox "error " & Err.Number 1044
                        If Err.Number <> 0 Then
                            Exit Do
                        End If
                        
                        If Selection.Parent.Name <> formulaCell.Parent.Name Then
                            
                            Print #1, Chr(34) & Selection.Parent.Name & "!" & Selection.Address(False, False) & Chr(34) _
                                    & " -> " & Chr(34) & formulaCell.Parent.Name & "!" & formulaCell.Address(False, False) & Chr(34)
                    
                        End If
                        
                        qCount = qCount + 1
                        'MsgBox qCount
                    Loop
                    
                End If
                
                If formulaCell.DirectPrecedents.Cells.Count > 0 Then
                    Set precedentsRange = formulaCell.DirectPrecedents.Cells
                    For Each precedentCell In precedentsRange
                        
                        'MsgBox formulaCell.Address(True, True) & "->" & precedentCell.Address(True, True)
                        
                        Print #1, Chr(34) & currentWorkSheet.Name & "!" & precedentCell.Address(False, False) & Chr(34) _
                                    & " -> " & Chr(34) & currentWorkSheet.Name & "!" & formulaCell.Address(False, False) & Chr(34)
                        
                        
                    Next precedentCell
                    
                End If
                
            Next formulaCell
        End If
        currentWorkSheet.ClearArrows
    Next currentWorkSheet
    
    Print #1, "}"
    
    Close #1 ' close file for writing
    
    shellString = "dot -Tpng " & thisWorkBook.Path & "\GeneratedGraphs\" & thisWorkBookName & ".dot -o " _
                    & thisWorkBook.Path & "\GeneratedGraphs\" & thisWorkBookName & ".png"
    
    Shell (shellString)
    
End Sub

Sub OpenWorkBookFilesForGraphExtraction()

    Dim strFile As String
    Dim srcWorkBook As Workbook
    Dim strFileCompletePath As String
    'testing opening workbooks in a directory
    
    
    strFile = Dir("C:\dot-work\CorpusKooker\")
    Do While Len(strFile) > 0
        'MsgBox strFile
        strFileCompletePath = "C:\dot-work\CorpusKooker\" & strFile
        Set srcWorkBook = Workbooks.Open(strFileCompletePath, True, True)
        Call ExtractWorkBookDataFlowGraph
        srcWorkBook.Close False
        Set srcWorkBook = Nothing
        strFile = Dir
        
    Loop
    
End Sub

Sub SmellDetectionHardCodedValues()
    Dim currentWorkSheet As Worksheet
    Dim formulaCellsInWorkSheet As Range
    Dim formulaCell As Range
    Dim formulaCellComment As Comment
    
    Set currentWorkSheet = ActiveSheet
    On Error Resume Next
    Set formulaCellsInWorkSheet = currentWorkSheet.Cells.SpecialCells(xlCellTypeFormulas)
    
    On Error Resume Next
    For Each formulaCell In formulaCellsInWorkSheet
        Set formulaCellComment = formulaCell.Comment
        If formulaCellComment Is Nothing Then
            'Do nothing
        Else
            If formulaCell.DirectPrecedents.Count = 0 Then
                
                'Do nothing
                formulaCell.Interior.ColorIndex = 17 ' light purple
                
            End If
        End If
        
    Next formulaCell
    
    
End Sub

Sub SmellDetectionOverWrittenWithConstants()
    Dim currentWorkSheet As Worksheet
    Dim constantCellsInWorkSheet As Range
    Dim constantCell As Range
    
    
    Set currentWorkSheet = ActiveSheet
    On Error Resume Next
    Set constantCellsInWorkSheet = currentWorkSheet.Cells.SpecialCells(xlCellTypeConstants)
    
    On Error Resume Next
    For Each constantCell In constantCellsInWorkSheet
    
        If IsNumeric(constantCell.Cells.Value) Then
            
            'top comment and bottom comment not empty
            If Not (Cells(constantCell.Row - 1, constantCell.Column).Comment Is Nothing) And _
            Not (Cells(constantCell.Row + 1, constantCell.Column).Comment Is Nothing) Then
                
                If Err.Number = 1004 Then
                    'do nothing
                    Err.Clear
                Else
                    'MsgBox Err.Number & " " & constantCell.Address
                    constantCell.Interior.ColorIndex = 27 ' yellow colour
                End If
                
            End If
            
            'left comment and right comment not empty
            If Not (Cells(constantCell.Row, constantCell.Column - 1).Comment Is Nothing) And _
            Not (Cells(constantCell.Row, constantCell.Column + 1).Comment Is Nothing) Then
                
                If Err.Number = 1004 Then
                    'do nothing
                    Err.Clear
                Else
                    
                    constantCell.Interior.ColorIndex = 27
                End If
                
                
            End If
            'MsgBox Cells(constantCell.Row + 1, constantCell.Column).Address
            
            'top comment and left comment not empty
            If Not (Cells(constantCell.Row - 1, constantCell.Column).Comment Is Nothing) And _
            Not (Cells(constantCell.Row, constantCell.Column - 1).Comment Is Nothing) Then
                
                If Err.Number = 1004 Then
                    'do nothing
                    Err.Clear
                Else
                    
                    constantCell.Interior.ColorIndex = 27
                End If
                
            End If
            
            'top comment and right comment not empty
            If Not (Cells(constantCell.Row - 1, constantCell.Column).Comment Is Nothing) And _
            Not (Cells(constantCell.Row, constantCell.Column + 1).Comment Is Nothing) Then
                
                constantCell.Interior.ColorIndex = 27
                
            End If
            
            'bottom comment and left comment not empty
            If Not (Cells(constantCell.Row + 1, constantCell.Column).Comment Is Nothing) And _
            Not (Cells(constantCell.Row, constantCell.Column - 1).Comment Is Nothing) Then
                
                If Err.Number = 1004 Then
                    'do nothing
                    Err.Clear
                Else
                    constantCell.Interior.ColorIndex = 27
                End If
                
            End If
            
            'bottom comment and right comment not empty
            If Not (Cells(constantCell.Row - 1, constantCell.Column).Comment Is Nothing) And _
            Not (Cells(constantCell.Row, constantCell.Column + 1).Comment Is Nothing) Then
                
                
                If Err.Number = 1004 Then
                    'do nothing
                    Err.Clear
                Else
                    
                    constantCell.Interior.ColorIndex = 27
                End If
                
            End If
            
            
        End If
        
    Next constantCell
    
    
End Sub

Sub SmellDetectionSameAsNeighbouringRange()
    Dim currentWorkSheet As Worksheet
    Dim formulaCellsInWorkSheet As Range
    Dim formulaCell As Range
    
    
    Set currentWorkSheet = ActiveSheet
    On Error Resume Next
    Set formulaCellsInWorkSheet = currentWorkSheet.Cells.SpecialCells(xlCellTypeFormulas)
    
    On Error Resume Next
    For Each formulaCell In formulaCellsInWorkSheet
    
            
            'top neigbour comment and formula comment are the same
            If Not (Cells(formulaCell.Row - 1, formulaCell.Column).Comment Is Nothing) Then
            
                If (StrComp(Cells(formulaCell.Row - 1, formulaCell.Column).Comment.Text, _
                    formulaCell.Comment.Text, vbTextCompare) = 0) Then
                    If Err.Number = 1004 Then
                        'do nothing
                        Err.Clear
                    Else
                        Cells(formulaCell.Row - 1, formulaCell.Column).Interior.ColorIndex = 46
                        formulaCell.Interior.ColorIndex = 46 'deep orange color
                    End If
                End If
                
            End If
            
            'bottom neigbour comment and formula comment are the same
            If Not (Cells(formulaCell.Row + 1, formulaCell.Column).Comment Is Nothing) Then
            
                If (StrComp(Cells(formulaCell.Row + 1, formulaCell.Column).Comment.Text, _
                    formulaCell.Comment.Text, vbTextCompare) = 0) Then
                    If Err.Number = 1004 Then
                        'do nothing
                        Err.Clear
                    Else
                        Cells(formulaCell.Row + 1, formulaCell.Column).Interior.ColorIndex = 46
                        formulaCell.Interior.ColorIndex = 46
                    End If
                End If
                
            End If
            
            'left neigbour comment and formula comment are the same
            If Not (Cells(formulaCell.Row, formulaCell.Column - 1).Comment Is Nothing) Then
            
                If (StrComp(Cells(formulaCell.Row, formulaCell.Column - 1).Comment.Text, _
                    formulaCell.Comment.Text, vbTextCompare) = 0) Then
                    If Err.Number = 1004 Then
                        'do nothing
                        Err.Clear
                    Else
                        Cells(formulaCell.Row, formulaCell.Column - 1).Interior.ColorIndex = 46
                        formulaCell.Interior.ColorIndex = 46
                    End If
                End If
                
            End If
            
            'right neigbour comment and formula comment are the same
            If Not (Cells(formulaCell.Row, formulaCell.Column + 1).Comment Is Nothing) Then
            
                If (StrComp(Cells(formulaCell.Row, formulaCell.Column + 1).Comment.Text, _
                    formulaCell.Comment.Text, vbTextCompare) = 0) Then
                    
                    If Err.Number = 1004 Then
                        'do nothing
                        Err.Clear
                    Else
                        Cells(formulaCell.Row, formulaCell.Column + 1).Interior.ColorIndex = 46
                        formulaCell.Interior.ColorIndex = 46
                    End If
                End If
                
            End If
            
            
        
    Next formulaCell
    
    
End Sub

Sub SmellDetectionReferencingLabel()
    Dim currentWorkSheet As Worksheet
    Dim formulaCellsInWorkSheet As Range
    Dim formulaCell As Range
    Dim formulaCellComment As Comment
    Dim directPrecedentsRange As Range
    Dim precedentCell As Range
    
    Set currentWorkSheet = ActiveSheet
    On Error Resume Next
    Set formulaCellsInWorkSheet = currentWorkSheet.Cells.SpecialCells(xlCellTypeFormulas)
    
    On Error Resume Next
    For Each formulaCell In formulaCellsInWorkSheet
        Set formulaCellComment = formulaCell.Comment
        If Not (formulaCellComment Is Nothing) Then
        
            If formulaCell.DirectPrecedents.Count = 0 Then
                'Do nothing
            Else
                Set directPrecedentsRange = formulaCell.DirectPrecedents.Cells
                For Each precedentCell In directPrecedentsRange
                    
                    If Not (IsNumeric(Cells(precedentCell.Row, precedentCell.Column))) Then
                        If Err.Number = 1004 Then
                            'do nothing
                            Err.Clear
                        Else
                            formulaCell.Interior.ColorIndex = 23 'light blue colour
                        End If
                    End If
                Next precedentCell
            End If
            
        End If
        
    Next formulaCell
End Sub

Sub DetectSmells()
    
    Cells.ClearComments

    Call DomainVisualizeAllRowToColumn
    Call SmellDetectionHardCodedValues 'purple
    Call SmellDetectionOverWrittenWithConstants 'yellow
    Call SmellDetectionSameAsNeighbouringRange 'deep orange
    Call SmellDetectionReferencingLabel 'light blue
    
End Sub

Sub DetectSmellsOpenWorkBookFilesForSmellDetection()

    Dim strFile As String
    Dim srcWorkBook As Workbook
    Dim strFileCompletePath As String
    Dim currentWorkSheet As Worksheet
    'testing opening workbooks in a directory
    
    
    strFile = Dir("C:\Users\hp\Documents\ASE_paper_v2_20180622_751\enron-samples\")
    Do While Len(strFile) > 0
        'MsgBox strFile
        strFileCompletePath = "C:\Users\hp\Documents\ASE_paper_v2_20180622_751\enron-samples\" & strFile
        Set srcWorkBook = Workbooks.Open(strFileCompletePath, True, True)
        For Each currentWorkSheet In srcWorkBook.Worksheets
            currentWorkSheet.Activate
            Call DetectSmells
        Next currentWorkSheet
        
        srcWorkBook.Close True
        Set srcWorkBook = Nothing
        strFile = Dir
        
    Loop
    
End Sub
