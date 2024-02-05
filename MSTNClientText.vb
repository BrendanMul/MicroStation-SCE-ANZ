Imports MicroStationSCE.SCEAccess1

<ComClass(MSTNClientText.ClassId, MSTNClientText.InterfaceId, MSTNClientText.EventsId)> _
Public Class MSTNClientText

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4e687c61-a5d8-4c8f-8d91-63627e4bb927"
    Public Const InterfaceId As String = "5e612790-0045-4b01-baf9-1dd419c168c0"
    Public Const EventsId As String = "4fb6fa0d-6cd0-43bb-9d8f-94818556b36a"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Heights As Array
    Public Widths As Array
    Public Weights As Array
    Public Colours As Array
    Public Levels As Array
    Public Fonts As Array
    Public Styles As Array
    Public SymbStyles As Array
    Public Loaded As Boolean
    Public Count As Integer

    '    Public Sub Load(Optional ByVal sDBFullName As String = "")
    '        'Load projects
    '        Dim vDB As Object
    '        Dim vFields As Object
    '        Dim aHeights() As String
    '        Dim aWidths() As String
    '        Dim aWeights() As String
    '        Dim aColours() As String
    '        Dim aLevels() As String
    '        Dim aFonts() As String
    '        Dim aStyles() As String
    '        Dim aSymbStyles() As String
    '        Dim i As Long
    '        Dim j As Long
    '        Dim k As Long

    '        If sDBFullName = "" Then
    '            sDBFullName = SCE.GetDatabase("LUText")
    '        End If

    '        If sDBFullName = "" Then GoTo SkipLoad

    '        vDB = GetAllDB(sDBFullName, "LUText", "Height")
    '        vFields = GetFieldsDB(sDBFullName, "LUText")

    '        If vDB(0) = "" And UBound(vDB) = 0 Then GoTo SkipLoad

    '        i = -1

    '        For j = 0 To UBound(vDB)
    '            i = i + 1
    '            ReDim Preserve aHeights(i)
    '            ReDim Preserve aWidths(i)
    '            ReDim Preserve aWeights(i)
    '            ReDim Preserve aColours(i)
    '            ReDim Preserve aLevels(i)
    '            ReDim Preserve aFonts(i)
    '            ReDim Preserve aStyles(i)
    '            ReDim Preserve aSymbStyles(i)

    '            For k = 0 To UBound(vFields)
    '                If UCase(vFields(k)) = "HEIGHT" Then
    '                    If IsNull(vDB(j + k)) = False Then aHeights(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "WIDTH" Then
    '                    If IsNull(vDB(j + k)) = False Then aWidths(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "WEIGHT" Then
    '                    If IsNull(vDB(j + k)) = False Then aWeights(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "COLOUR" Then
    '                    If IsNull(vDB(j + k)) = False Then aColours(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "LEVELNAME" Then
    '                    If IsNull(vDB(j + k)) = False Then aLevels(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "FONT" Then
    '                    If IsNull(vDB(j + k)) = False Then aFonts(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "STYLE" Then
    '                    If IsNull(vDB(j + k)) = False Then aStyles(i) = Trim(vDB(j + k))
    '                ElseIf UCase(vFields(k)) = "SYMBSTYLE" Then
    '                    If IsNull(vDB(j + k)) = False Then aSymbStyles(i) = Trim(vDB(j + k))
    '                End If
    '            Next 'vFields
    '            j = j + UBound(vFields)
    '        Next 'vDB

    '        Heights = aHeights
    '        Widths = aWidths
    '        Weights = aWeights
    '        Colours = aColours
    '        Levels = aLevels
    '        Fonts = aFonts
    '        Styles = aStyles
    '        SymbStyles = aSymbStyles

    '        Erase vDB
    '        Erase vFields
    '        Erase aHeights
    '        Erase aWidths
    '        Erase aWeights
    '        Erase aColours
    '        Erase aLevels
    '        Erase aFonts
    '        Erase aStyles
    '        Erase aSymbStyles

    '        Count = UBound(Heights) + 1
    '        Loaded = True

    '        Exit Sub
    'SkipLoad:
    '        Loaded = True
    '        Count = 0
    '    End Sub

    Function GetIndex(ByVal sHeight As String) As Long
        'Return index of specified text height
        Dim i As Long
        Dim iIndex As Long

        iIndex = -1
        For i = 0 To (Count - 1)
            If UCase(sHeight) = UCase(Heights(i)) Then
                iIndex = i
                Exit For
            End If
        Next

        GetIndex = iIndex
    End Function

End Class


