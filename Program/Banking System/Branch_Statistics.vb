Public Class Branch_Statistics

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

    Private Sub Branch_Statistics_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Dim cmd As New OleDb.OleDbCommand("SELECT COUNT(*) FROM EMPLOYEE WHERE EMP_BRANCH=?", conn)
        cmd.Parameters.AddWithValue("?", branch_name)
        Dim da As New OleDb.OleDbDataAdapter
        da.SelectCommand = cmd
        Dim dt As New DataTable
        da.Fill(dt)
        Label6.Text = dt.Rows(0).Item(0)
        Dim cmd2 As New OleDb.OleDbCommand("SELECT COUNT(*) FROM CUSTOMER WHERE CUS_BRANCH=?", conn)
        cmd2.Parameters.AddWithValue("?", branch_name)
        Dim da2 As New OleDb.OleDbDataAdapter
        da2.SelectCommand = cmd2
        Dim dt2 As New DataTable
        da2.Fill(dt2)
        Label7.Text = dt2.Rows(0).Item(0)
        Dim cmd3 As New OleDb.OleDbCommand("SELECT COUNT(*) FROM CUSTOMER,BORROWER WHERE CUS_ACC_NO=BO_ACC_NO AND CUS_BRANCH=?", conn)
        cmd3.Parameters.AddWithValue("?", branch_name)
        Dim da3 As New OleDb.OleDbDataAdapter
        da3.SelectCommand = cmd3
        Dim dt3 As New DataTable
        da3.Fill(dt3)
        Label8.Text = dt3.Rows(0).Item(0)
        Dim cmd4 As New OleDb.OleDbCommand("SELECT SUM(T.AMOUNT) FROM TRANSACTION T, EMPLOYEE E WHERE E.EMP_NO=T.EMP_NO AND E.EMP_BRANCH=? AND TYPE='Deposit'", conn)
        cmd4.Parameters.AddWithValue("?", branch_name)
        Dim da4 As New OleDb.OleDbDataAdapter
        da4.SelectCommand = cmd4
        Dim dt4 As New DataTable
        da4.Fill(dt4)
        Label9.Text = dt4.Rows(0).Item(0)
        Dim cmd5 As New OleDb.OleDbCommand("SELECT SUM(T.AMOUNT) FROM TRANSACTION T, EMPLOYEE E WHERE E.EMP_NO=T.EMP_NO AND E.EMP_BRANCH=? AND TYPE='Withdraw'", conn)
        cmd5.Parameters.AddWithValue("?", branch_name)
        Dim da5 As New OleDb.OleDbDataAdapter
        da5.SelectCommand = cmd5
        Dim dt5 As New DataTable
        da5.Fill(dt5)
        Label10.Text = dt5.Rows(0).Item(0)
        conn.Close()
    End Sub
End Class