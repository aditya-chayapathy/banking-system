Public Class Dismiss_Employee

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
            Button2.Hide()
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
        TextBox1.Text = ""
        DataGridView1.DataSource = ""
        Button2.Hide()
        Branch_Manager_Menu.Show()
    End Sub

    Private Sub HowToDismissAnEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToDismissAnEmployeeToolStripMenuItem.Click
        MessageBox.Show("STEP 1  -  Enter the employee no of the customer who is to be dismissed and then click on the 'Search' button" + System.Environment.NewLine + "STEP 2  -  Click the 'Dismiss' button to dismiss the employee", "How To Dismiss An Employee ?", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Dismiss_Employee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim flag As New Integer
        flag = 0
        If IsNumeric(TextBox1.Text) = False Then
            flag = 1
        End If
        If flag = 1 Then
            MessageBox.Show("Please Enter Employee No !!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            Dim da As New OleDb.OleDbDataAdapter
            Dim dt As New DataTable
            Dim res As New DialogResult
            Dim y As String
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM EMPLOYEE WHERE EMP_NO=? AND EMP_MGR_NO=?", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            cmd.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            da.SelectCommand = cmd
            da.Fill(dt)
            DataGridView1.DataSource = dt
            y = TextBox1.Text
            If DataGridView1.RowCount = 2 Then
                Button2.Show()
            Else
                MessageBox.Show("Please Enter Valid Employee No !!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Text = ""
            End If
            conn.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim res As DialogResult
        res = MessageBox.Show("Are You Sure You Want To Dismiss This Employee ?", "Dismiss Employee ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Global_Variables.random.acc_no = TextBox1.Text
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("DELETE FROM EMPLOYEE WHERE EMP_NO=?", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            cmd.ExecuteNonQuery()
            MessageBox.Show("Employee Dismissed !!!", "Employee Dismissed", MessageBoxButtons.OK, MessageBoxIcon.None)
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
        Else
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
        End If
    End Sub
End Class