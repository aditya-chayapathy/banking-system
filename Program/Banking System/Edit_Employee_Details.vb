Public Class Edit_Employee_Details

    Private Sub Edit_Employee_Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        conn.Open()
        Dim cmd As New OleDb.OleDbCommand("SELECT * FROM EMPLOYEE WHERE EMP_NO=?", conn)
        cmd.Parameters.AddWithValue("?", Global_Variables.random.user_id)
        Dim da As New OleDb.OleDbDataAdapter
        da.SelectCommand = cmd
        Dim dt As New DataTable
        da.Fill(dt)
        Label10.Text = dt.Rows(0).Item(0)
        TextBox1.Text = dt.Rows(0).Item(1)
        TextBox2.Text = dt.Rows(0).Item(2)
        TextBox3.Text = dt.Rows(0).Item(3)
        Label11.Text = dt.Rows(0).Item(4)
        Label12.Text = dt.Rows(0).Item(5)
        Label13.Text = dt.Rows(0).Item(6)
        Label14.Text = dt.Rows(0).Item(7)
        TextBox4.Text = dt.Rows(0).Item(8)
        conn.Close()
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

    Private Sub HowToEditMyDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToEditMyDetailsToolStripMenuItem.Click
        MessageBox.Show("Make necessary changes for the fields provided and click on the 'Save' button", "How To Edit ?", MessageBoxButtons.OK)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        conn.Open()
        Dim cmd1 As New OleDb.OleDbCommand("UPDATE EMPLOYEE SET EMP_NAME=?,EMP_ADD=?,EMP_CONTACT_NO=?,EMP_PWORD=? WHERE EMP_NO=?", conn)
        cmd1.Parameters.AddWithValue("?", TextBox1.Text)
        cmd1.Parameters.AddWithValue("?", TextBox2.Text)
        cmd1.Parameters.AddWithValue("?", CDbl(TextBox3.Text))
        cmd1.Parameters.AddWithValue("?", TextBox4.Text)
        cmd1.Parameters.AddWithValue("?", Label10.Text)
        cmd1.ExecuteNonQuery()
        MessageBox.Show("Details Successfully Updated !!! " + System.Environment.NewLine + "Please Restart the application to reflect the changes !!!", "Details Updated !!!", MessageBoxButtons.OK)
        Me.Hide()
        Employee_Details.Show()
    End Sub

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        Me.Hide()
        Employee_Details.Show()
    End Sub
End Class