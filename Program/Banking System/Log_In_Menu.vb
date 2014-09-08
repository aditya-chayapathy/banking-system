Public Class Log_In_Menu

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        MessageBox.Show("Each employee is given an Employee ID and a password." + System.Environment.NewLine + "This is provided to the employee on the day of him/her joining the bank." + System.Environment.NewLine + "Please enter these credentials to proceed......", "Log In Help", MessageBoxButtons.OK)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Dim res As DialogResult
        res = MessageBox.Show("Are You Sure You Want To Exit ?", "EXIT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            MessageBox.Show("Created By C.Aditya !!!", "Thank You !!!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2)
            End
        Else
            Me.Show()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim des As String
        Dim SAPI
        Dim gv1 As New random
        SAPI = CreateObject("SAPI.spvoice")
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        conn.Open()
        Dim cmd As New OleDb.OleDbCommand("SELECT EMP_DES FROM EMPLOYEE WHERE EMP_NAME=? AND EMP_NO=? AND EMP_PWORD=?", conn)
        cmd.Parameters.AddWithValue("?", TextBox1.Text)
        cmd.Parameters.AddWithValue("?", TextBox2.Text)
        cmd.Parameters.AddWithValue("?", TextBox3.Text)
        Dim dr As OleDb.OleDbDataReader
        dr = cmd.ExecuteReader
        Global_Variables.random.user_name = TextBox1.Text
        Global_Variables.random.user_id = TextBox2.Text
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        If dr.Read() Then
            SAPI.Speak("Log In Successful")
            Global_Variables.random.log_in_datetime = String.Format("{0:dd/MM/yyyy" + System.Environment.NewLine + " HH:mm}", DateTime.Now)
            des = dr.GetString(0)
            If des = "Branch Manager" Then
                Me.Hide()
                Branch_Manager_Menu.Show()
            ElseIf des = "Employee" Then
                Me.Hide()
                Employee_Menu.Show()
            End If
        Else
            MessageBox.Show("LOGIN FAILED !!!" + System.Environment.NewLine + "Please ensure you have entered the correct credentials", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        conn.Close()
    End Sub

    Private Sub Log_In_Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class