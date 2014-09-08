Public Class Branch_Transactions

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

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        Me.Hide()
        Branch_Manager_Menu.Show()
    End Sub

    Private Sub Branch_Transactions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        conn.Open()
        Dim cmd1 As New OleDb.OleDbCommand("SELECT EMP_BRANCH FROM EMPLOYEE WHERE EMP_NO=?", conn)
        cmd1.Parameters.AddWithValue("?", Global_Variables.random.user_id)
        Dim da1 As New OleDb.OleDbDataAdapter
        da1.SelectCommand = cmd1
        Dim dt1 As New DataTable
        da1.Fill(dt1)
        Dim branch_name As String
        branch_name = dt1.Rows(0).Item(0)
        Dim cmd As New OleDb.OleDbCommand("SELECT T.CUS_ACC_NO,T.EMP_NAME,T.EMP_NO,T.TYPE,T.AMOUNT,T.TIME FROM TRANSACTION T,EMPLOYEE E WHERE T.EMP_NO=E.EMP_NO AND EMP_BRANCH=?", conn)
        cmd.Parameters.AddWithValue("?", branch_name)
        Dim da As New OleDb.OleDbDataAdapter
        da.SelectCommand = cmd
        Dim dt As New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        conn.Close()
    End Sub
End Class