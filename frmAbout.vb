Public Class frmAbout

    Public Shared tblAbout As DataTable

    ''    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
    ''        Me.Dispose()
    ''    End Sub

    ''    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ''        lsvAbout.Items.Add("MicroStation SCE Version")
    ''        lsvAbout.Items.Add("Client Name")
    ''        lsvAbout.Items.Add("Client Version")
    ''        lsvAbout.Items.Add("Client Owner")

    ''        If SCECore.Loaded = False Then SCE.Load()
    ''        lblSCEVersion.Caption = SCESettings.SCEVersion
    ''        lblClientName.Caption = ClientSettings.Client
    ''        lblVersion.Caption = ClientSettings.BuildVersion
    ''        lblOwner.Caption = ClientSettings.Owner

    ''        Communicator.InitialiseCommunicatorImageList(imlImages)
    ''        oContact = Communicator.GetUserContact(lblOwner.Caption)

    ''        If oContact Is Nothing Then
    ''            imgPresence.Picture = LoadPicture(ImagePath & "presence_16-off.JPG")
    ''            lblOwner.ForeColor = vbBlack
    ''            lblOwner.Font.Underline = False
    ''        Else
    ''            imgPresence.Picture = LoadPicture(ImagePath & SCECommunicator.GetStatusImage(oContact.Status))
    ''            lblOwner.Tag = oContact.SigninName
    ''            lblOwner.Caption = Communicator.GetRealName(oContact.FriendlyName)
    ''            lblOwner.ForeColor = vbBlue
    ''            lblOwner.Font.Underline = True
    ''            imgIM.Picture = LoadPicture(ImagePath & "send_instant_message_16x16.JPG")
    ''            imgIM.Visible = True
    ''            imgEmail.Picture = LoadPicture(ImagePath & "email.JPG")
    ''            imgEmail.Visible = True
    ''        End If

    ''    End Sub

    Private Sub frmAbout_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If Not tblAbout Is Nothing Then
            dgvAbout.DataSource = tblAbout
        End If
    End Sub

End Class