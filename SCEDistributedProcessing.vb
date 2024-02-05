Public Class SCEDistributedProcessing

    'Distributed Processing allows the ability to issue tasks to remote machines simultaneously to perform a task

    Public oDistributedProcessing As DataTable

    Public Sub AdvertiseForProcessing()
        'Advertise current PC for remote processing
        'Ability to nominate Master user for submitting task

    End Sub

    Public Sub DeployTasks()
        Dim oRow As DataRow = Nothing

        'Collate task for remote machine(s)
        oRow.Table.Columns.Add("Master") 'Asset Number of master PC issuing task
        oRow.Table.Columns.Add("UserName") 'user invoking task
        oRow.Table.Columns.Add("AssetNumber") 'Remote PC asset number
        oRow.Table.Columns.Add("Software") 'software
        oRow.Table.Columns.Add("Region") 'Region
        oRow.Table.Columns.Add("Client") 'Client
        oRow.Table.Columns.Add("ProjectName") 'Project name (if required)
        oRow.Table.Columns.Add("File") ' file to run task on
        oRow.Table.Columns.Add("TaskName") 'Name of task being executed
        oRow.Table.Columns.Add("Command") 'Command to run 
        oRow.Table.Columns.Add("Status") 'status of task - idle, running, complete
        oRow.Table.Columns.Add("Comment") 'comment for process of running task

        'Deploy task to remote PC

        'Check and report status of task, use timer to update

        'Release remote machine from processing task

    End Sub

End Class
