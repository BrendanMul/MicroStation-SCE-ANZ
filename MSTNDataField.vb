<ComClass(MSTNDataField.ClassId, MSTNDataField.InterfaceId, MSTNDataField.EventsId)> _
Public Class MSTNDataField

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "06850354-6358-4a77-af61-e944e07b0c02"
    Public Const InterfaceId As String = "4cb1b51a-e2bd-4097-acc6-5b240e403f7e"
    Public Const EventsId As String = "66ce8263-a9c4-49ec-8c91-d9a6538bec07"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

Function GetDataFields(ByVal sCell As String, Optional ByVal oModel As ModelReference)

        'get text from specified cell

        Dim i As Long
        Dim ee As ElementEnumerator
        Dim eeCell As ElementEnumerator
        Dim oCell As CellElement
        Dim oText As TextElement
        Dim aData() As String

        i = 0

        If oModel Is Nothing Then
            oModel = ActiveModelReference
        End If

        oCell = FindCell(sCell, True, oModel)
        If oCell Is Nothing Then
            ShowError("Cell: " & Chr(34) & sCell & " was not found. (from SCECore\DataField\GetDataFields)")
            GoTo NoCell
        End If
        eeCell = oCell.GetSubElements

        Do While eeCell.MoveNext

            If TypeOf eeCell.Current Is TextElement Then

                i = i + 1
                ReDim Preserve aData(i)
                aData(i) = eeCell.Current.AsTextElement.Text

            End If

        Loop
NoCell:
        If i = 0 Then
            ShowError("Cell " & Chr(34) & sCell & Chr(34) & " was not found.")
            ReDim aData(0)
        End If
        'Return value
        GetDataFields = CVar(aData)

    End Function

Function GetDataField(ByVal sCell As String, ByVal iIndex As Long, Optional ByVal oModel As ModelReference)

        'get text from specified cell

        Dim i As Long
        Dim ee As ElementEnumerator
        Dim eeCell As ElementEnumerator
        Dim oCell As CellElement
        Dim oText As TextElement
        Dim aData() As String

        i = 0

        If oModel Is Nothing Then
            oModel = ActiveModelReference
        End If

        oCell = FindCell(sCell, True, oModel)
        If oCell Is Nothing Then
            ShowError("Cell: " & Chr(34) & sCell & " was not found. (from SCECore\DataField\GetDataFields)")
            GoTo NoCell
        End If
        eeCell = oCell.GetSubElements

        Do While eeCell.MoveNext

            If TypeOf eeCell.Current Is TextElement Then

                i = i + 1
                If i = iIndex Then
                    GetDataField = eeCell.Current.AsTextElement.Text
                    Exit Do
                End If
            End If

        Loop
NoCell:
        If i = 0 Then
            ShowError("Cell " & Chr(34) & sCell & Chr(34) & " was not found.")
            ReDim aData(0)
        End If
        'Return value
        'GetDataFields = CVar(aData)

    End Function

    Sub SetDataFields(ByVal sCell As String, ByVal vDataFields As Object)

        'set text from specified datafields

        Dim i As Long
        Dim ee As ElementEnumerator
        Dim eeCell As ElementEnumerator
        Dim oCell As CellElement
        Dim oText As TextElement
        Dim aData() As String

        CadInputQueue.SendCommand("MARK") 'Set undo mark

        i = 0
        oCell = FindCell(sCell, True)
        eeCell = oCell.GetSubElements

        Do While eeCell.MoveNext

            If TypeOf eeCell.Current Is TextElement Then

                i = i + 1

                If Len(vDataFields(i)) < Len(eeCell.Current.AsTextElement.Text) Then

                    vDataFields(i) = vDataFields(i) & Space(Len(eeCell.Current.AsTextElement.Text) - Len(vDataFields(i)))

                End If

                eeCell.Current.AsTextElement.Text = vDataFields(i)
                eeCell.Current.AsTextElement.Rewrite()
                eeCell.Current.AsTextElement.Redraw()

            End If

        Loop

    End Sub

    Sub SetDataField(ByVal sCell As String, ByVal iIndex As Long, ByVal sValue As Object)

        'set text from specified datafield

        Dim i As Long
        Dim ee As ElementEnumerator
        Dim eeCell As ElementEnumerator
        Dim oCell As CellElement
        Dim oText As TextElement
        Dim aData() As String

        i = 0
        oCell = FindCell(sCell, True)
        If oCell Is Nothing Then
            ShowError("Cell: " & Chr(34) & sCell & " was not found. (from SCECore\DataField\SetDataField)")
            Exit Sub
        End If
        eeCell = oCell.GetSubElements

        Do While eeCell.MoveNext

            If TypeOf eeCell.Current Is TextElement Then

                i = i + 1

                If i = iIndex Then

                    If Len(sValue) < Len(eeCell.Current.AsTextElement.Text) Then

                        sValue = sValue & Space(Len(eeCell.Current.AsTextElement.Text) - Len(sValue))

                    End If

                    eeCell.Current.AsTextElement.Text = sValue
                    eeCell.Current.AsTextElement.Rewrite()
                    eeCell.Current.AsTextElement.Redraw()

                    Exit Do

                End If

            End If

        Loop

        oCell = Nothing
        eeCell = Nothing

    End Sub

Sub IndexDataFields(Optional ByVal bReverse As Boolean)
        'Index all datafields
        Dim i As Long
        Dim j As Long
        Dim ee As ElementEnumerator
        Dim eeCell As ElementEnumerator
        Dim oCell As CellElement
        Dim oText As TextElement
        Dim aData() As String
        Dim sCells As String
        Dim vSplit As Object

        CadInputQueue.SendCommand("MARK") 'Set undo mark

        i = 0
        sCells = TitleBorders.DataFieldCells(TitleBorders.GetBorderIndex)
        If InStr(sCells, ";") > 0 Then
            vSplit = Split(sCells, ";")
            For j = 0 To UBound(vSplit)
                oCell = FindCell(vSplit(j), True)
                If Not oCell Is Nothing Then
                    Exit For
                End If
            Next
        Else
            oCell = FindCell(TitleBorders.DataFieldCells(TitleBorders.GetBorderIndex), True)
        End If

        If oCell Is Nothing Then
            MsgBox("Cell not found.", vbCritical, "SCE - Error")
            Exit Sub
        End If

        eeCell = oCell.GetSubElements

        Do While eeCell.MoveNext
            If TypeOf eeCell.Current Is TextElement Then
                If bReverse = True Then
                    eeCell.Current.AsTextElement.Text = ""
                    eeCell.Current.AsTextElement.Rewrite()
                Else
                    i = i + 1
                    eeCell.Current.AsTextElement.Text = i
                    eeCell.Current.AsTextElement.Rewrite()
                End If
            End If
        Loop
        CadInputQueue.SendCommand("UPDATE ALL")
    End Sub

End Class


