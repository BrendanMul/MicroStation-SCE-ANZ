'Imports System
'Imports System.Net
'Imports System.IO
'Imports System.Text
'Imports System.Data.SqlClient
'Imports System.Threading
'Imports System.Xml
'Imports Microsoft.Win32
'Imports SKM.Common.xml
'Imports System.Configuration.ConfigurationSettings

'Imports System.DirectoryServices

'Public Class SCEActiveDirectory

'    Public Const CATEGORY_CORPORATE_GROUP As String = "Group CAD System Administrators"
'    Public Const CATEGORY_CLIENT_GROUP As String = "Group CAD Project Leaders"

'    Dim arGroupMembership As Collection = New Collection

'    ''Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
'    ''    Dim sPassword As String

'    ''    sUsername = txtUsername.Text
'    ''    sPassword = txtPassword.Text

'    ''    ' Username must contain something
'    ''    If sUsername.Trim <> "" Then
'    ''        Try
'    ''            Dim bAuthenticated As Boolean = False

'    ''            ' Authenticate user against Active Directory
'    ''            Me.Cursor = Cursors.WaitCursor
'    ''            bAuthenticated = ADAuthenticate(sUsername, sPassword)
'    ''            Me.Cursor = Cursors.Default

'    ''            If bAuthenticated Then
'    ''                ' Obtain group name to which user belongs (user must belong to at least one CAD security group to proceed)
'    ''                If (ADGetUserGroups()) Then
'    ''                    If IsMemberOfGroup(CATEGORY_CORPORATE_GROUP) Then
'    ''                        ' Get a list of category item to select from
'    ''                        If (GetConfigurationCategories()) Then
'    ''                            ' Proceed to options panel
'    ''                            nCurrentForm += 1
'    ''                            ShowCurrentForm(True)
'    ''                        Else
'    ''                            MsgBox("Cannot obtain connection to database.")
'    ''                        End If
'    ''                    ElseIf IsMemberOfGroup(CATEGORY_CLIENT_GROUP) Then

'    ''                        If (GetConfigurationCategories()) Then
'    ''                            PopulateClientConfigurationOptions("", "2")
'    ''                            nCurrentForm += 1
'    ''                            ShowCurrentForm(True)
'    ''                        Else
'    ''                            MsgBox("Cannot obtain connection to database.")
'    ''                        End If
'    ''                    Else
'    ''                        MsgBox("You do not have access to create content in the SCE please contact your Content Administrator if it is required. Not a member of the required groups.")
'    ''                    End If
'    ''                Else
'    ''                    MsgBox("You do not have access to create content in the SCE please contact your Content Administrator if it is required." & vbCrLf & "User " & sUsername & " could not be Authenticated against AD")
'    ''                End If
'    ''            Else
'    ''                'MsgBox("For the SKMCAD Environment ContentWizard to work you must have access to the SKM network." & vbCr & "Check that you are connected to the network and try again.")
'    ''                MsgBox("Check your user name and password.")
'    ''            End If
'    ''        Catch ex As Exception
'    ''            Me.Cursor = Cursors.Default
'    ''            MsgBox("For the SKMCAD Environment ContentWizard to work you must have access to the SKM network." & vbCr & "Check that you are connected to the network and try again.")
'    ''            'MsgBox("Error authenticating. " & ex.Message)
'    ''        End Try
'    ''    Else
'    ''        MsgBox("Please specify a valid username and password.")
'    ''    End If
'    ''End Sub








'#Region " Active directory communication "
'    Private Function ADAuthenticate(ByVal sUsername As String, ByVal sPassword As String) As Boolean
'        'Try
'        Dim ws As New ActiveDirectory.svcDirectory
'        Dim bAuthenticated As Boolean = ws.Authenticate(sUsername, sPassword)

'        If (bAuthenticated) Then
'            ' User is authenticated, save username
'            sAuthenticatedUID = sUsername

'            ' Get this person's email address
'            sAuthenticatedEmail = ADEmailAddress(sUsername)
'        End If

'        Return bAuthenticated

'        'Catch ex As Exception
'        '    '  MsgBox("Error Happened: " & ex.Message & " " & ex.StackTrace)
'        '    Return False
'        'End Try

'    End Function

'    Private Function ADEmailAddress(ByVal sUsername As String) As String
'        Dim sEmailAddress As String = ""
'        Dim ws As New ActiveDirectory.svcDirectory
'        ws.Timeout = -1

'        'GAD Added "ActiveDirectory." to the line below
'        Dim pThisPerson As ActiveDirectory.Person = ws.GetUserByUID(sUsername, 0, True)
'        If Not (pThisPerson Is Nothing) Then
'            sEmailAddress = pThisPerson.Mail
'        End If

'        Return sEmailAddress
'    End Function

'    Private Function ADGetNameMatches(ByVal sCriteria As String) As Collection
'        Dim ws As New ActiveDirectory.svcDirectory
'        Dim cMatches As Collection = New Collection

'        'GAD Added "ActiveDirectory." to the line below
'        Dim Matches() As ActiveDirectory.Person

'        Try
'            Matches = ws.GetUsersByName(sCriteria, False)

'            'GAD Added "ActiveDirectory." to the line below
'            For Each pThisPerson As ActiveDirectory.Person In Matches
'                cMatches.Add(pThisPerson.UID)
'            Next
'        Catch ex As Exception

'        End Try

'        Return cMatches
'    End Function

'    Private Function ADGetUserGroups() As Boolean
'        Dim ws As New ActiveDirectory.svcDirectory
'        Dim nLastIndex As Integer = 0

'        'GAD Removed line below as variable no longer used
'        'Dim sGroup As String

'        Try
'            Me.Cursor = Cursors.WaitCursor
'            'GAD Added "ActiveDirectory." to the line below
'            Dim oUser As ActiveDirectory.User = ws.GetUser(sAuthenticatedUID)
'            Me.Cursor = Cursors.Default
'            If (oUser Is Nothing) Then
'                Return False
'            Else
'                'GAD Added "ActiveDirectory." to the line below
'                For Each thisGroup As ActiveDirectory.Group In oUser.GroupMembership
'                    arGroupMembership.Add(thisGroup.Name)
'                Next
'                Return True
'            End If

'        Catch ex As Exception
'            'MsgBox(ex.Message)
'            ws.Dispose()
'            Return False
'        End Try

'        ws.Dispose()
'    End Function

'    Private Function IsMemberOfGroup(ByVal sGroupName As String) As Boolean
'        Dim bIsMember As Boolean = False

'        If Not (arGroupMembership Is Nothing) Then

'            For Each sGroup As String In arGroupMembership
'                If sGroup.ToLower.CompareTo(sGroupName.ToLower) = 0 Then
'                    bIsMember = True
'                    Exit For
'                End If
'            Next
'        End If

'        Return bIsMember
'    End Function


'#End Region

'End Class
