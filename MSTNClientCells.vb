Imports MicroStationSCE.MSTNSCECore
Imports MicroStationSCE.SCEAccess1

<ComClass(MSTNClientCells.ClassId, MSTNClientCells.InterfaceId, MSTNClientCells.EventsId)> _
Public Class MSTNClientCells

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e77ddba3-3b62-4114-9af3-1dc88e34efdf"
    Public Const InterfaceId As String = "38be7994-9278-436c-bbf1-32468fe00a4a"
    Public Const EventsId As String = "22fcfd6c-609a-42e3-b747-d0d43f59c3fe"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public CellNames As Array
    Public CellLibraries As Array
    Public Regions As Array
    Public Clients As Array
    Public Keywords As Array
    Public XPos As Array
    Public YPos As Array
    Public Origins As Array
    Public BorderSizes As Array
    Public Scales As Array
    Public Rotations As Array
    Public Descriptions As Array
    Public Count As Integer
    Public Loaded As Boolean
    Public FullName As String

    Sub Load(Optional ByVal sRegion As String = "", Optional ByVal sClient As String = "")

        'Load client cells

        Dim vDB As Object
        Dim vFields As Object
        Dim aCellNames() As String
        Dim aCellLibraries() As String
        Dim aRegions() As Object
        Dim aClients() As String
        Dim aKeywords() As String
        Dim aXPos() As String
        Dim aYPos() As String
        Dim aOrigins() As String
        Dim aBorderSizes() As String
        Dim aScales() As String
        Dim aRotations() As String
        Dim aDescriptions() As String
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim sDBFullName As String

        If sRegion = "" Then
            sDBFullName = ClientDB
        ElseIf UCase(sRegion) = "SCE" And UCase(sClient) = "SCE" Then
            sDBFullName = CoreDB
        Else
            sDBFullName = SCERegionPath
            sDBFullName = sDBFullName & Region & "\"
            sDBFullName = sDBFullName & Client & "\"
            sDBFullName = sDBFullName & "Database\SCE Client - " & Client & "." & SCESettings.DBExtension
        End If

        If My.Computer.FileSystem.FileExists(sDBFullName) = False Then GoTo SkipLoad

        FullName = sDBFullName

        vDB = GetAllDB(sDBFullName, "LUCell", "CellName")
        vFields = GetFieldsDB(sDBFullName, "LUCell")

        If vDB(0) = "" And UBound(vDB) = 0 Then GoTo SkipLoad

        i = -1

        For j = 0 To UBound(vDB)

            i = i + 1

            ReDim Preserve aCellNames(i)
            ReDim Preserve aCellLibraries(i)
            ReDim Preserve aRegions(i)
            ReDim Preserve aClients(i)
            ReDim Preserve aKeywords(i)
            ReDim Preserve aXPos(i)
            ReDim Preserve aYPos(i)
            ReDim Preserve aOrigins(i)
            ReDim Preserve aBorderSizes(i)
            ReDim Preserve aScales(i)
            ReDim Preserve aRotations(i)
            ReDim Preserve aDescriptions(i)

            For k = 0 To UBound(vFields)

                If UCase(vFields(k)) = "CELLNAME" Then

                    If IsNull(vDB(j + k)) = False Then aCellNames(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "CELLLIBRARY" Then

                    If IsNull(vDB(j + k)) = False Then aCellLibraries(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "REGION" Then

                    If IsNull(vDB(j + k)) = False Then aRegions(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "CLIENT" Then

                    If IsNull(vDB(j + k)) = False Then aClients(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "KEYWORD" Then

                    If IsNull(vDB(j + k)) = False Then aKeywords(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "XPOS" Then

                    If IsNull(vDB(j + k)) = False Then aXPos(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "YPOS" Then

                    If IsNull(vDB(j + k)) = False Then aYPos(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "ORIGIN" Then

                    If IsNull(vDB(j + k)) = False Then aOrigins(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "BORDERSIZE" Then

                    If IsNull(vDB(j + k)) = False Then aBorderSizes(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "SCALE" Then

                    If IsNull(vDB(j + k)) = False Then aScales(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "ROTATION" Then

                    If IsNull(vDB(j + k)) = False Then aRotations(i) = Trim(vDB(j + k))

                ElseIf UCase(vFields(k)) = "DESCRIPTION" Then

                    If IsNull(vDB(j + k)) = False Then aDescriptions(i) = Trim(vDB(j + k))

                End If

            Next 'vFields

            j = j + UBound(vFields)

        Next 'vDB

        CellNames = aCellNames
        CellLibraries = aCellLibraries
        Regions = aRegions
        Clients = aClients
        Keywords = aKeywords
        XPos = aXPos
        YPos = aYPos
        Origins = aOrigins
        BorderSizes = aBorderSizes
        Scales = aScales
        Rotations = aRotations
        Descriptions = aDescriptions

        Erase vDB
        Erase vFields
        Erase aCellNames
        Erase aCellLibraries
        Erase aRegions
        Erase aClients
        Erase aKeywords
        Erase aXPos
        Erase aYPos
        Erase aOrigins
        Erase aBorderSizes
        Erase aScales
        Erase aRotations
        Erase aDescriptions

        Count = UBound(CellNames) + 1
        Loaded = True

        Exit Sub
SkipLoad:
        Loaded = True
        Count = 0

    End Sub

    Function GetIndex(ByVal sCellName As String) As Long

        'Return index of specified cell
        'i=-1 means no match found

        Dim i As Long

        i = -1

        For i = 0 To (Count - 1)

            If UCase(CellNames(i)) = UCase(sCellName) Then

                Exit For

            End If

        Next

        'Return value
        GetIndex = i

    End Function

Function GetIndexWithKeyword(ByVal sKeyword As String, Optional ByVal sBordersize As String) As Variant

        'Return all cells with specified keyword

        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim aIndexes() As String
        Dim vSplit As Array

        k = -1

        For i = 0 To (Count - 1)

            vSplit = Split(Keywords(i), ";")

            For j = 0 To UBound(vSplit)

                If UCase(vSplit(j)) = UCase(sKeyword) Then

                    If Trim(BorderSizes(i)) = "" Or UCase(BorderSizes(i)) = UCase(sBordersize) Or Trim(sBordersize) = "" Then

                        k = k + 1
                        ReDim Preserve aIndexes(k)
                        aIndexes(k) = i

                        'Exit For

                    End If

                End If

            Next 'vSplit

        Next 'ClientCells

        If k = -1 Then 'No cells found

            ReDim aIndexes(0)

        End If

        'Return value
        GetIndexWithKeyword = aIndexes

    End Function

    Function HasKeyword(ByVal iIndex As Long, ByVal sKeyword As String) As Boolean

        'Return boolean if specified cell has keyword

        HasKeyword = InSplit(Keywords(iIndex), sKeyword)

    End Function

End Class


