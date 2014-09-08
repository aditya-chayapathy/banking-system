Public Class Employee_Details

    Private Sub Employee_Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        conn.Open()
        Dim cmd As New OleDb.OleDbCommand("SELECT * FROM EMPLOYEE WHERE EMP_NO=?", conn)
        cmd.Parameters.AddWithValue("?", Global_Variables.random.user_id)
        Dim da As New OleDb.OleDbDataAdapter
        da.SelectCommand = cmd
        Dim dt As New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        conn.Close()
    End Sub

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        Me.Hide()
        Employee_Menu.Show()
    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            Log_In_Menu.Show()
            SAPI = CreateObject("SAPI.spvoice")
            SAPI.Speak("Successfully Logged Out")
            Global_Variables.random.log_out_datetime = String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now)
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn.Open()
            Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO LOG_IN VALUES(?,?,?,?)", conn)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.user_name)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.log_in_datetime)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.log_out_datetime)
            cmd2.ExecuteNonQuery()
            conn.Close()
        Else
            Me.Show()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Edit_Employee_Details.Show()
    End Sub
End Class