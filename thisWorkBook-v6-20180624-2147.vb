Option Explicit
Option Compare Text

Dim previouslyActiveCellAddress As String
Dim ActiveCellAddress As String
Dim previouslyActiveCell As Range
Dim myActiveCellAddress As String
Dim myActiveCellPrecedentsRange As Range
Dim targetChanged As Range
Dim colourIndexArray() As Variant
Dim indexint As Integer
Dim editedComment As String
Dim uneditedComment As String
Dim strippedEditedComment As String
Dim strippedunEditedComment As String
Dim rawEditedComment As String
Dim formulaSystemChanged As Boolean
Dim uneditedCommentPrecedents As Range
Dim myCellEditedComment As Range
Dim gCell As Range
Dim rowSearchNotErrorDetect As Integer
Dim colSearchNotErrorDetect As Integer
Dim gObject As Range
Dim resultCell As Range
Dim It_Is_A_Number As Boolean
Dim findTempRange As Range
Dim findTempRange2 As Range
Dim findAddress As String
Dim foundCount As Integer
Dim iRet As Integer
Dim strPrompt As String
Dim strTitle As String
Dim highlightedColumnLabelCell As Range

Private Sub Workbook_Open()

    Set previouslyActiveCell = Range("A1")
    Set highlightedColumnLabelCell = Nothing
    indexint = 1
    editedComment = ""
    uneditedComment = ""
    rawEditedComment = ""
    formulaSystemChanged = False
    
    'Dim usedCell As Range
    'On Error Resume Next
    'For Each usedCell In ActiveSheet.UsedRange
        'If usedCell.Interior.ColorIndex = 28 Then
           'usedCell.Interior.ColorIndex = xlNone
        'End If
    'Next
    
End Sub



Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
        
        Dim myComment As Object
        Set myComment = ActiveCell.Comment
        If myComment Is Nothing Then
        Else
            ActiveCell.Comment.Visible = False
        End If
        
        Application.EnableEvents = False
        
        
        Dim myCell As Range
        
        On Error Resume Next
        If ActiveCell.DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            
            On Error Resume Next
            For Each myCell In ActiveCell.DirectPrecedents.Cells
                
                myCell.Interior.ColorIndex = xlNone
            Next
              
        End If
        
        Application.EnableEvents = True
        
    
End Sub



Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    
    'mafikeng
        Application.EnableEvents = False
       Dim myTargetCell As Range
       If Target.Cells.Count = 1 Then
            For Each myTargetCell In Target.Cells
                
                On Error Resume Next
                If myTargetCell.HasFormula Then
                   If myTargetCell.DirectPrecedents.Cells.Count > 0 Then
                        'MsgBox "I have a formula with precedents"
                        
                        'Application.Run ("'DomainRealWorldVisualize.xlam'!CallToDomainVisualizeSingleFormulaCellMafikeng")
                    End If
                End If
            Next myTargetCell
       End If
       Application.EnableEvents = True
    'end mafikeng
    
    Application.DisplayCommentIndicator = xlNoIndicator
    
    formulaSystemChanged = False
    
    
    ' If the whole row or column has been selected, stop further processing
    Dim myCols As Integer
    myCols = Selection.Columns.Count
    
    If myCols >= 16384 Then
        'MsgBox myCols
        Exit Sub
    End If
    
    Dim myRows As Long
    myRows = Selection.Rows.Count
    
    If myRows >= 1048576 Then
        'MsgBox myRows
        Exit Sub
    End If
    
    ' End - If the whole row or column has been selected, stop further processing
    
    ' we hereafter, do further processing
    
    If previouslyActiveCellAddress = "" Then
    
        'MsgBox previouslyActiveCellAddress
        Set previouslyActiveCell = Range("A1")
        previouslyActiveCellAddress = Trim(previouslyActiveCell.Offset(0, 0).Address(False, False))
    
        Exit Sub
    
    End If
     
    
    'MsgBox previouslyActiveCell.Address
    'Set previouslyActiveCell = Range(previouslyActiveCellAddress)
    On Error Resume Next
    previouslyActiveCellAddress = Trim(previouslyActiveCell.Offset(0, 0).Address(False, False))
    Set previouslyActiveCell = Range(previouslyActiveCellAddress)
    ActiveCellAddress = Trim(ActiveCell.Offset(0, 0).Address(False, False))
    
    Dim myCell As Range
    Dim rRange As Range
    Dim myComment As Object
    
    
     
    With Range(previouslyActiveCellAddress)
        
        Set myComment = Range(previouslyActiveCellAddress).Comment
        If myComment Is Nothing Then
        Else
            Range(previouslyActiveCellAddress).Comment.Visible = False
            'MsgBox Range(previouslyActiveCellAddress).Comment.Text
            
            
            
            
            
            Dim intRow As Long
            intRow = 0
            'Dim searchRange As Range
            'searchRange = ActiveSheet.UsedRange
            Dim cellAddStripped As String
            Dim cellAddTokensStripped() As String
            Dim rowNumberStripped As String
            Dim colNumberStripped As String
            
            Dim intIndex As Integer
            
            
            ' mafikeng rawEditedComment = "=" & Replace(rawEditedComment, "|", "")
            ' mafikeng rawEditedComment = Replace(rawEditedComment, "...", ":")
            
            'rawEditedComment = Replace(rawEditedComment, " ", " ")
            'rawEditedComment = Trim(rawEditedComment)
            'MsgBox rawEditedComment
            
            Application.EnableEvents = False
            
            formulaSystemChanged = True
            
            Set uneditedCommentPrecedents = Range(previouslyActiveCellAddress).DirectPrecedents.Cells
            'MsgBox "pre" & uneditedCommentPrecedents.Count
            On Error Resume Next
            If uneditedCommentPrecedents.Count > 0 Then
                
                'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
                indexint = 0
                'Dim colorPreviousCellsEditedComment As Variant
                On Error Resume Next
                For Each myCell In uneditedCommentPrecedents.Cells
                    'myCell.Interior.ColorIndex = xlNone
                    'colorPreviousCells = colourIndexArray(indexint)
                    'if previous color was the comment light green colour (28), then change to none
                    
                    If colourIndexArray(indexint) = 28 Then
                        myCell.Interior.ColorIndex = xlNone
                    Else
                        myCell.Interior.ColorIndex = colourIndexArray(indexint)
                    End If
                    
                    indexint = indexint + 1
                Next
                  
            End If
            
            
            'Range(previouslyActiveCellAddress).ClearContents
            If highlightedColumnLabelCell Is Nothing Then
                'do nothing
            Else
                highlightedColumnLabelCell.Interior.ColorIndex = xlNone
            End If
            
            'mafikeng taken out Range(previouslyActiveCellAddress).Formula = rawEditedComment
            
            'Range(previouslyActiveCellAddress).Formula = "=IF(A1<3000,""Small hjj"", ""Large jkkk"")"
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            Application.EnableEvents = True
            
            'Range(previouslyActiveCellAddress).Formula = rawEditedComment
            
      
            
            
            
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            'Range("B18").Formula = "=IF(A1<3000,""Small hjj"", ""Large jkkk"")"
            
            'MsgBox "striEdi" & strippedEditedComment
            'MsgBox "striUnEdi" & strippedunEditedComment
            
            
            
            
        End If ' myComment Is Nothing
        
        Application.EnableEvents = False
        
        On Error Resume Next
        If Range(previouslyActiveCellAddress).DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            indexint = 0
            'Dim colorPreviousCells As Variant
            On Error Resume Next
            For Each myCell In Range(previouslyActiveCellAddress).DirectPrecedents.Cells
                'myCell.Interior.ColorIndex = xlNone
                'colorPreviousCells = colourIndexArray(indexint)
                'if previous color was the comment light green colour (28), then change to none
                
                If colourIndexArray(indexint) = 28 Then
                    myCell.Interior.ColorIndex = xlNone
                Else
                    myCell.Interior.ColorIndex = colourIndexArray(indexint)
                End If
                
                indexint = indexint + 1
            Next
              
        End If
        
        
             
        
        ' if the previous cell is empty after deleting its contents,
        ' remove its corresponding comment if any
        
        If IsEmpty(Range(previouslyActiveCellAddress)) Then
            'MsgBox "nothing"
            Set myComment = Range(previouslyActiveCellAddress).Comment
            If myComment Is Nothing Then
            Else
                
                
                Range(previouslyActiveCellAddress).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                Range(previouslyActiveCellAddress).ClearComments
                
            End If
            
        End If
        
        '08-06-2017: to make sure that a newly created formula is indicated as such with a label
        If Range(previouslyActiveCellAddress).HasFormula Then
            Range(previouslyActiveCellAddress).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            Range(previouslyActiveCellAddress).Borders(xlEdgeRight).ColorIndex = 26
            Range(previouslyActiveCellAddress).Borders(xlEdgeRight).Weight = xlThick
        End If
        
        Application.EnableEvents = True
        
    End With
    
    With Range(ActiveCellAddress)
        
        Set myComment = Range(ActiveCellAddress).Comment
        If myComment Is Nothing Then
        Else
            
            Range(ActiveCellAddress).Comment.Visible = True
            uneditedComment = Range(ActiveCellAddress).Comment.Text
            'MsgBox "un" & uneditedComment
            
        End If
        
        Application.EnableEvents = False
        
        On Error Resume Next
        If Range(ActiveCellAddress).DirectPrecedents.Count > 0 Then
            
            'Set rRange = Range(previouslyActiveCellAddress).DirectPrecedents
            
            
            
            Set myActiveCellPrecedentsRange = Range(ActiveCellAddress).DirectPrecedents
            myActiveCellAddress = ActiveCellAddress
            indexint = 0
            Dim precedentCurrentColor As Variant
            ReDim Preserve colourIndexArray(0 To Range(ActiveCellAddress).DirectPrecedents.Count)
            On Error Resume Next
            For Each myCell In Range(ActiveCellAddress).DirectPrecedents.Cells
                
                precedentCurrentColor = myCell.Interior.ColorIndex
                colourIndexArray(indexint) = CInt(precedentCurrentColor)
                indexint = indexint + 1
                myCell.Interior.ColorIndex = 28
            Next
            
        End If
        
        Application.EnableEvents = True
        
    End With
    
    
    Set previouslyActiveCell = ActiveCell
    
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    ' If the whole row or column has been selected, stop further processing
    
    'whole row selected
    
    Dim myCell As Range
    Dim myComment As Object
    
    Dim myCols As Integer
    myCols = Selection.Columns.Count
    
    If myCols >= 16384 Then
        'MsgBox myCols
            Application.EnableEvents = False
                
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    myCell.Interior.ColorIndex = xlNone
                    Set myComment = myCell.Comment
                    If myComment Is Nothing Then
                    Else
                        myCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                        myCell.ClearComments
                        
                    End If
                Next
                Set previouslyActiveCell = Range("A1")
                Application.EnableEvents = True
        Exit Sub
    End If
    
    'whole column selected
    Dim myRows As Long
    myRows = Selection.Rows.Count
    
    If myRows >= 1048576 Then
        'MsgBox myRows
        
        Application.EnableEvents = False
                
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    myCell.Interior.ColorIndex = xlNone
                    Set myComment = myCell.Comment
                    If myComment Is Nothing Then
                    Else
                        myCell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
                        myCell.ClearComments
                    End If
                Next
                Set previouslyActiveCell = Range("A1")
                Application.EnableEvents = True
        Exit Sub
    End If
    
    ' End - If the whole row or column has been selected, stop further processing

    'making sure that once a group of cells has been deleted, the corresponding
    'highlighted precedents are de-highlighted
    
    Dim aCell As Range
   
    For Each aCell In Target.Cells '-- Target may contains more than one cells.
        If aCell.Formula = "" Then
            'MsgBox "Cell " & aCell.Address & " in " & Sh.Name & " has been cleared or deleted."
            'MsgBox "My active cell address" & myActiveCellAddress
            'MsgBox " acell.address" & aCell.Address
            If myActiveCellAddress = aCell.Offset(0, 0).Address(False, False) Then
                'MsgBox "Same"
                Application.EnableEvents = False
                'Dim myCell As Range
                On Error Resume Next
                For Each myCell In myActiveCellPrecedentsRange.Cells
                    
                    myCell.Interior.ColorIndex = xlNone
                Next
                Application.EnableEvents = True
                
            End If
            
        Else
            'MsgBox "Cell " & aCell.Address & " in " & Sh.Name & " has been changed."
        End If
    Next
   
   
   ' Executing the domain visualization to regenerate all the comments once  changes
   ' have been made to the spreadsheet - dont regenerate for number inputs
    'If Not (IsNumeric(Target.Value)) Or Target.Formula = "" Then
    
    Dim cel As Range
    Dim dcel As Range
    Dim directDependents As Range
    
    If Target.Cells.Count >= 1 Then
        For Each cel In Target
            'MsgBox cel.Address(False, False)
            On Error Resume Next
            If cel.directDependents.Count > 0 Then
                Set directDependents = cel.directDependents
                For Each dcel In directDependents
                    'MsgBox "direct" & dcel.Address(False, False)
                Next
            End If 'cel.DirectDependents.Count > 0
        Next
    End If 'Target.Cells.Count >= 1
        'MsgBox Target
    
    If Target.Cells.Count = 1 Then
        'MsgBox Target
    
        If IsNumeric(Target) Then
            'do nothing
            If Target.HasFormula Or Target.Formula = "" Then
                If Target.Formula = "" Then
                    ' original  Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
                    'Application.Run ("'DomainRealWorldVisualize.xlam'!CallToDomainVisualizeSingleFormulaCellMafikeng")
                Else 'Target.HasFormula
                'MsgBox "wawwa"
                    'For Each cel In Target
                    'GetTargetChanged (cel)
                    If formulaSystemChanged = True Then
                        'MsgBox "system changed"
                    Else
                        ' original mafikeng Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
                        'Application.Run ("'DomainRealWorldVisualize.xlam'!CallToDomainVisualizeSingleFormulaCellMafikeng")
                    End If
                    'Next
                End If
                
            End If 'Target.HasFormula
        Else
            ' original mafikeng Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            'Application.Run ("'DomainRealWorldVisualize.xlam'!CallToDomainVisualizeSingleFormulaCellMafikeng")
            
        End If ' IsNumeric(Target)
            
    Else
        If Target.Cells.Count > 1 Then
            ' original mafikeng Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            
            'Application.Run ("'DomainRealWorldVisualize.xlam'!CallToDomainVisualizeSingleFormulaCellMafikeng")
        End If ' Target.Cells.Count > 1
    End If 'Target.Cells.Count = 1
    
    
        'Dim ResultState As Integer
        
        'ResultState = Application.Run("'DomainRealWorldVisualize.xlam'!ReturnState")
        'ResultState = 2 'make column to row as default
        'If ResultState = 1 Then
            'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllRowToColumn")
        'Else
            'If ResultState = 2 Then
                'run the default column to row
                'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllColumnToRow")
            'Else
                
                    'Application.Run ("'DomainRealWorldVisualize.xlam'!DomainVisualizeAllRowToColumn")
            'End If
        'End If
    
    
    
End Sub

