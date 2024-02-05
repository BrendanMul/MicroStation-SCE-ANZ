Public Class SCETask

    'log tasks to be executed in MicroStation

    Public Shared Tasks() As Task
    Public Shared LastTaskID As Integer
    Public Shared DrawingList() As stDrawingList
    Public Shared AreaList() As stAreaList
    Public Shared Count As Integer

    Public Structure Task
        Dim TaskID As Integer
        Dim TaskName As String
        Dim Action As String
        Dim ProjectName As String
        Dim DrawingListName As String
        Dim AreaList As String
        Dim Status As String
        Dim State As String
        Dim Author As String
        Dim Updated As String
        Dim OpenReadOnly As Boolean
        Dim Schedule As Boolean
        Dim Time As String
        Dim Frequency As String
        Dim ConfigVars As String
        Dim Type As String
    End Structure

    Public Structure stDrawingList
        Dim DrawingNumber As String
        Dim Area As String
        Dim ProjectName As String
        Dim Revision As String
        Dim Path As String
        Dim Status As String
        Dim OutputFile As String
        Dim Updated As String
        Dim DBFullName As String
        Dim PrintIssued As Boolean
    End Structure

    Public Structure stAreaList
        Dim Area As String
        Dim ProjectName As String
        Dim DrawingType As String
        Dim DrawingPath As String
        Dim DBPath As String
        Dim AreaTable As String
        Dim Updated As String
        Dim Drawings() As stDrawings
        Dim sDBFullName As String
    End Structure

    Public Structure stDrawings
        Dim DrawingNumber As String
        Dim Path As String
        Dim Area As String
    End Structure

    Public Shared Function IsTasksScheduled() As Boolean
        Dim dsDataset As DataSet
        Dim bReturn As Boolean = False

        If My.Computer.FileSystem.FileExists(JEGCore.UserDatabase) = True Then
            If SCEAccess.TableExists(JEGCore.UserDatabase, "RGTasks") = True Then
                dsDataset = SCEAccess.GetRowsDB(JEGCore.UserDatabase, "RGTasks", "TaskName", "Schedule", "True")
                If SCEVB.IsNullDataSet(dsDataset) = False Then
                    bReturn = True
                End If
            End If
        End If

        Return bReturn
    End Function

    ''' <summary>
    ''' Add new task to task list
    ''' </summary>
    ''' <param name="oTask"></param>
    ''' <remarks></remarks>
    Public Shared Sub Add(ByVal oTask As Task, Optional ByVal sDBFullName As String = "", Optional ByVal sPassword As String = "")
        Dim sFields As String = SCEAccess.GetTableColumns("RGTasks", "User")
        Dim sValues As String = ""

        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserDatabase
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Count = Count + 1
        ReDim Preserve Tasks(Count - 1)
        Tasks(Count - 1) = oTask

        sValues = oTask.TaskID
        sValues = sValues & "|" & oTask.TaskName
        sValues = sValues & "|" & oTask.Action
        sValues = sValues & "|" & oTask.ProjectName
        sValues = sValues & "|" & oTask.DrawingListName
        sValues = sValues & "|" & oTask.AreaList
        sValues = sValues & "|" & oTask.Status
        sValues = sValues & "|" & oTask.State
        sValues = sValues & "|" & oTask.Author
        sValues = sValues & "|" & oTask.Updated
        sValues = sValues & "|" & oTask.OpenReadOnly
        sValues = sValues & "|" & oTask.Schedule
        sValues = sValues & "|" & oTask.Time
        sValues = sValues & "|" & oTask.Frequency

        If SCEAccess.TableExists(sDBFullName, "RGTasks", sPassword) = False Then
            SCEAccess.CreateTable(sDBFullName, "RGTasks", sFields, sPassword)
        End If

        SCEAccess.SetDB(sDBFullName, "RGTasks", sFields, sValues, sPassword)
    End Sub

    ''' <summary>
    ''' Add new drawing list
    ''' </summary>
    ''' <param name="oDrawingList"></param>
    ''' <remarks></remarks>
    Public Shared Sub AddDrawingList(ByVal sDrawingListName As String, ByVal oDrawingList As stDrawingList, Optional ByVal sPassword As String = "", Optional ByVal sTemplateDB As String = "")
        Dim sFields As String = SCEAccess.GetTableColumns("RGDrawingList", "User")
        Dim sValues As String = ""

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        Count = Count + 1
        sValues = oDrawingList.DrawingNumber
        sValues = sValues & "|" & oDrawingList.Area
        sValues = sValues & "|" & oDrawingList.ProjectName
        sValues = sValues & "|" & oDrawingList.Revision
        sValues = sValues & "|" & oDrawingList.Path
        sValues = sValues & "|" & oDrawingList.Status
        sValues = sValues & "|" & oDrawingList.OutputFile
        sValues = sValues & "|" & oDrawingList.Updated

        'Create database
        If My.Computer.FileSystem.FileExists(oDrawingList.DBFullName) = False Then
            SCEAccess.CreateDB(oDrawingList.DBFullName, sTemplateDB)
        End If

        'create table 
        If SCEAccess.TableExists(oDrawingList.DBFullName, sDrawingListName, sPassword) = False Then
            SCEAccess.CreateTable(oDrawingList.DBFullName, sDrawingListName, sFields, sPassword)
        End If

        SCEAccess.SetDB(oDrawingList.DBFullName, sDrawingListName, sFields, sValues, sPassword)
    End Sub

    ''''' <summary>
    ''''' update current Task
    ''''' </summary>
    ''''' <param name="oTask"></param>
    ''''' <remarks></remarks>
    ''Public Shared Sub Update(ByVal oTask As Task)
    ''    SCEAccess.UpdateField(SCECore.UserDatabase, "RGTasks", "Status", oTask.Status, "DrawingNumber", oTask.DrawingNumber)
    ''End Sub

    ''' <summary>
    ''' update current Drawing List
    ''' </summary>
    ''' <param name="oTask"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateDrawingList(ByVal oTask As Task, ByVal oDrawingListItem As stDrawingList, Optional ByVal sPassword As String = "")
        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If
        SCEAccess.UpdateField(oDrawingListItem.DBFullName, oTask.DrawingListName, "Status", oDrawingListItem.Status, "DrawingNumber", oDrawingListItem.DrawingNumber, sPassword)
        SCEAccess.UpdateField(oDrawingListItem.DBFullName, oTask.DrawingListName, "Updated", Now, "DrawingNumber", oDrawingListItem.DrawingNumber, sPassword)
    End Sub

    Public Shared Sub UpdateTask(ByVal oTask As Task, ByVal oDrawingListItem As stDrawingList, Optional ByVal sPassword As String = "", Optional ByVal sDBFullName As String = "")
        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If
        'If sDBFullName = "" Then
        '    sDBFullName = SCECore.UserDatabase
        'Else
        '    sDBFullName = SCECore.UserPath & "\Database\TEMP - Tasks.mdb"
        'End If
        SCEAccess.UpdateField(oDrawingListItem.DBFullName, "RGTasks", "Status", oTask.Status, "TaskName", oTask.TaskName, sPassword)
        SCEAccess.UpdateField(oDrawingListItem.DBFullName, "RGTasks", "Updated", Now, "TaskName", oTask.TaskName, sPassword)
    End Sub

    ''' <summary>
    ''' remove task from list
    ''' </summary>
    ''' <param name="oTask"></param>
    ''' <remarks></remarks>
    Public Shared Sub Remove(ByVal oTask As Task, Optional ByVal sDBFullName As String = "", Optional ByVal sPassword As String = "")
        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserDatabase
        Else
            'sDBFullName = SCECore.UserPath & "\Database\TEMP - Tasks.mdb"
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If
        SCEAccess.DelDB(sDBFullName, "RGTasks", "TaskName", oTask.TaskName, sPassword)
    End Sub

    ''' <summary>
    ''' clear task list
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Clear(Optional ByVal sDBFullName As String = "")
        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserDatabase
        Else
            'sDBFullName = SCECore.UserPath & "\Database\TEMP - Tasks.mdb"
        End If
        SCEAccess.DeleteTable(sDBFullName, "RGTasks")
    End Sub

    ''' <summary>
    ''' Retrieve task collection
    ''' </summary>
    ''' <returns>Task collection</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTasks(Optional ByVal sPassword As String = "", Optional ByVal sDBFullName As String = "")
        Dim dsDataSet As DataSet
        Dim oRow As DataRow
        Dim oTask As Task = Nothing
        Dim i As Integer = -1

        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserDatabase
        Else
            'sDBFullName = SCECore.UserPath & "\Database\TEMP - Tasks.mdb"
        End If

        Tasks = Nothing

        'Check whether any tasks have been assigned 
        If SCEAccess.TableExists(sDBFullName, "RGTasks", sPassword) = False Then
            GoTo ExitSub
        End If

        dsDataSet = SCEAccess.GetAllDB(sDBFullName, "RGTasks", "TaskID", sPassword)
        LastTaskID = 0

        If SCEVB.IsNullDataSet(dsDataSet) = True Then
            GoTo ExitSub
        End If

        For Each oRow In dsDataSet.Tables(0).Rows
            oTask = New Task
            oTask.TaskID = oRow.Item("TaskID")
            oTask.TaskName = oRow.Item("TaskName")
            oTask.Action = oRow.Item("Action")
            oTask.ProjectName = oRow.Item("ProjectName")
            oTask.DrawingListName = oRow.Item("DrawingList")
            If Not oRow.Item("AreaList") Is System.DBNull.Value Then
                oTask.AreaList = oRow.Item("AreaList")
            End If
            If Not oRow.Item("Status") Is System.DBNull.Value Then
                oTask.Status = oRow.Item("Status")
            End If
            If Not oRow.Item("State") Is System.DBNull.Value Then
                oTask.State = oRow.Item("State")
            End If
            oTask.Author = oRow.Item("Author")
            oTask.Updated = oRow.Item("Updated")
            If Not oRow.Item("OpenReadOnly") Is System.DBNull.Value Then
                If Trim(oRow.Item("OpenReadOnly")) <> "" Then
                    oTask.OpenReadOnly = oRow.Item("OpenReadOnly")
                End If
            End If
            If Not oRow.Item("Schedule") Is System.DBNull.Value Then
                oTask.Schedule = oRow.Item("Schedule")
            End If
            If Not oRow.Item("Time") Is System.DBNull.Value Then
                oTask.Time = oRow.Item("Time")
            End If
            If Not oRow.Item("Frequency") Is System.DBNull.Value Then
                oTask.Frequency = oRow.Item("Frequency")
            End If
            i = i + 1
            ReDim Preserve Tasks(i)
            Tasks(i) = oTask

            If oTask.TaskID > LastTaskID Then
                LastTaskID = oTask.TaskID
            End If
        Next
ExitSub:
        Return Tasks
    End Function

    Public Shared Function GetTask(ByVal sTaskName As String, Optional ByVal sDBFullName As String = "") As Task
        'Return task
        Dim oTask As Task = Nothing
        Dim Tasks() As Task

        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserDatabase
        Else
            'sDBFullName = SCECore.UserPath & "\Database\TEMP - Tasks.mdb"
        End If

        Tasks = GetTasks(, sDBFullName)

        For Each oTask In Tasks
            If oTask.TaskName = sTaskName Then
                Exit For
            End If
        Next

        Return oTask
    End Function

    ''' <summary>
    ''' Retrieve Drawing List 
    ''' </summary>
    ''' <returns>Drawing List</returns>
    ''' <remarks></remarks>
    Public Shared Function GetDrawingList(ByVal sProjectName As String, ByVal sDrawingListName As String, Optional ByVal sPassword As String = "", Optional ByVal sDBFullName As String = "")
        Dim dsDataSet As DataSet
        Dim oRow As DataRow
        Dim oDrawingList As stDrawingList = Nothing
        Dim i As Integer = -1

        DrawingList = Nothing
        If sDBFullName = "" Then
            sDBFullName = JEGCore.UserPath & "\Database\" & "Drawing List - " & sProjectName & ".mdb"
        End If

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Check any Tasks have been registered")
        'Check whether any tasks have been assigned 
        If SCEAccess.TableExists(sDBFullName, sDrawingListName, sPassword) = False Then
            GoTo ExitSub
        End If

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Get Drawing List")
        dsDataSet = SCEAccess.GetAllDB(sDBFullName, sDrawingListName, "DrawingNumber", sPassword)

        If SCEVB.IsNullDataSet(dsDataSet) = True Then
            If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Drawing list is empty\null")
            GoTo ExitSub
        End If

        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Storing Drawing List")
        For Each oRow In dsDataSet.Tables(0).Rows
            oDrawingList = New stDrawingList
            oDrawingList.DrawingNumber = oRow.Item("DrawingNumber")
            oDrawingList.Area = oRow.Item("Area")
            oDrawingList.ProjectName = oRow.Item("ProjectName")
            If Not oRow.Item("Revision") Is System.DBNull.Value Then
                oDrawingList.Revision = oRow.Item("Revision")
            End If
            oDrawingList.Path = oRow.Item("Path")
            If Not oRow.Item("Status") Is System.DBNull.Value Then
                oDrawingList.Status = oRow.Item("Status")
            End If
            If Not oRow.Item("OutputFile") Is System.DBNull.Value Then
                oDrawingList.OutputFile = oRow.Item("OutputFile")
            End If
            oDrawingList.Updated = oRow.Item("Updated")
            oDrawingList.DBFullName = sDBFullName
            i = i + 1
            ReDim Preserve DrawingList(i)
            DrawingList(i) = oDrawingList
        Next
ExitSub:
        If JEGCore.bDebug = True Then JEGCore.RecordSequence(System.Reflection.MethodBase.GetCurrentMethod.Module.Name, System.Reflection.MethodBase.GetCurrentMethod.Name(), "Get Drawing List complete")
        Return DrawingList
    End Function

    ''' <summary>
    ''' Retrieve Area List
    ''' </summary>
    ''' <param name="sProjectName"></param>
    ''' <param name="sAreaListName"></param>
    ''' <param name="sPassword"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAreaList(ByVal sProjectName As String, ByVal sAreaListName As String, Optional ByVal sPassword As String = "")
        Dim dsDataSet As DataSet
        Dim oRow As DataRow
        Dim oAreaList As stAreaList = Nothing
        Dim i As Integer = -1
        Dim sDBFullName As String

        sDBFullName = JEGCore.UserPath & "\Database\" & "Area List - " & sProjectName & ".mdb"

        If sPassword = "" Then
            sPassword = JEGCore.Password
        End If

        'Check whether any tasks have been assigned 
        If SCEAccess.TableExists(sDBFullName, sAreaListName, sPassword) = False Then
            GoTo ExitSub
        End If

        dsDataSet = SCEAccess.GetAllDB(sDBFullName, sAreaListName, "Area", sPassword)

        If SCEVB.IsNullDataSet(dsDataSet) = True Then
            GoTo ExitSub
        End If

        For Each oRow In dsDataSet.Tables(0).Rows
            oAreaList = New stAreaList
            oAreaList.Area = oRow.Item("Area")
            oAreaList.ProjectName = oRow.Item("ProjectName")
            oAreaList.DrawingType = oRow.Item("DrawingType")
            If Not oRow.Item("DrawingPath") Is System.DBNull.Value Then
                oAreaList.DrawingPath = oRow.Item("DrawingPath")
            End If
            If Not oRow.Item("DBPath") Is System.DBNull.Value Then
                oAreaList.DBPath = oRow.Item("DBPath")
            End If
            If Not oRow.Item("AreaTable") Is System.DBNull.Value Then
                oAreaList.AreaTable = oRow.Item("AreaTable")
            End If
            oAreaList.Updated = oRow.Item("Updated")
            i = i + 1
            ReDim Preserve areaList(i)
            AreaList(i) = oAreaList
        Next
ExitSub:
        Return AreaList
    End Function

End Class
