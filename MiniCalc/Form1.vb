
Public Class Form1

#Region "Constants"
    Const MENU_Break As Integer = 100
    Const MENU_AlwaysOnTop As Integer = 101
    Const MENU_CopyToClipboard As Integer = 102
    Const MENU_About As Integer = 103
#End Region

    Private WithEvents mobjSubclassedSystemMenu As SubclassedSystemMenu

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.WindowLocation = Me.Location
            My.Settings.AlwaysOnTop = Me.TopMost
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Check to see if this is the first time we've loaded the my.settings for this version, if so then upgrade
        If My.Settings.CallUpgrade = True Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
        End If

        Me.Location = My.Settings.WindowLocation
        Me.TopMost = My.Settings.AlwaysOnTop

        'Add menu items
        mobjSubclassedSystemMenu = New SubclassedSystemMenu(Me.Handle.ToInt32)
        mobjSubclassedSystemMenu.AppendSeperatorToSystemMenu(mobjSubclassedSystemMenu.SysMenuHandle, MENU_Break)
        mobjSubclassedSystemMenu.AppendToSystemMenu(mobjSubclassedSystemMenu.SysMenuHandle, MENU_AlwaysOnTop, "Always On Top")
        mobjSubclassedSystemMenu.AppendToSystemMenu(mobjSubclassedSystemMenu.SysMenuHandle, MENU_CopyToClipboard, "Copy Result to Clipboard")
        mobjSubclassedSystemMenu.AppendToSystemMenu(mobjSubclassedSystemMenu.SysMenuHandle, MENU_About, "About")

        'Set check state of menu items
        If Me.TopMost Then
            mobjSubclassedSystemMenu.Checked(mobjSubclassedSystemMenu.SysMenuHandle, MENU_AlwaysOnTop) = True
        End If

        If My.Settings.CopyResultsToClipboard Then
            mobjSubclassedSystemMenu.Checked(mobjSubclassedSystemMenu.SysMenuHandle, MENU_CopyToClipboard) = True
        End If

    End Sub

    Private Sub btnCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalc.Click
        DoCalc()
    End Sub

    Sub DoCalc()
        Try

            Dim CalcSuccess As Boolean = False

            If txtEntry.Text.ToUpper.StartsWith("&H") Then
                'Its a hex value, just try and convert it to decimal
                txtResult.Text = CLng(txtEntry.Text.Trim).ToString

                CalcSuccess = True
            Else

                If txtEntry.Text.Trim.Length > 0 Then
                    txtResult.Text = EvalExpression(txtEntry.Text).ToString.Trim

                    CalcSuccess = True
                Else
                    txtResult.Text = ""
                End If
            End If

            If CalcSuccess Then
                CopyResultToClipboard()
            End If
        Catch
            txtResult.Text = "Calculation Error!"
        Finally
            txtEntry.SelectAll()
            txtEntry.Focus()
        End Try
    End Sub

    Private Function EvalExpression(ByVal expr As String) As Double
        ' create a new DataTable containing a calculated column
        Dim dt As New DataTable()
        dt.Columns.Add("Expr", GetType(Double), expr)
        ' add a dummy row
        dt.Rows.Add(dt.NewRow)
        ' return the value of the calculated column
        Return CDbl(dt.Rows(0).Item("Expr"))
    End Function

    Private Sub CopyResultToClipboard()
        If My.Settings.CopyResultsToClipboard Then
            My.Computer.Clipboard.Clear()
            My.Computer.Clipboard.SetText(txtResult.Text)
        End If
    End Sub

    Private Sub mobjSubclassedSystemMenu_SysMenuClick(ByVal menuItemID As Integer) Handles mobjSubclassedSystemMenu.SysMenuClick
        Select Case menuItemID
            Case MENU_AlwaysOnTop
                Me.TopMost = Not Me.TopMost
                mobjSubclassedSystemMenu.Checked(mobjSubclassedSystemMenu.SysMenuHandle, MENU_AlwaysOnTop) = Me.TopMost

            Case MENU_CopyToClipboard
                My.Settings.CopyResultsToClipboard = Not My.Settings.CopyResultsToClipboard
                mobjSubclassedSystemMenu.Checked(mobjSubclassedSystemMenu.SysMenuHandle, MENU_CopyToClipboard) = My.Settings.CopyResultsToClipboard

            Case MENU_About     'About
                AboutBox1.ShowDialog()
        End Select
    End Sub
End Class
