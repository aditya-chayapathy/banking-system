Public Class Hire_Employee

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            Log_In_Menu.Show()
            SAPI = CreateObject("SAPI.spvoice")
            SAPI.Speak("Successfully Logged Out")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox6.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            ComboBox1.ResetText()
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
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        ComboBox1.ResetText()
        Branch_Manager_Menu.Show()
    End Sub

    Private Sub HowToHireEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToHireEmployeeToolStripMenuItem.Click
        MessageBox.Show("Fill all the fields provided and click on the 'Hire' button", "How To Hire Employee ?", MessageBoxButtons.OK)
    End Sub

    Private Sub Hire_Employee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        Dim dr As OleDb.OleDbDataReader
        conn.Open()
        Dim cmd1 As New OleDb.OleDbCommand("SELECT BR_NAME FROM BRANCH", conn)
        dr = cmd1.ExecuteReader
        While dr.Read()
            ComboBox1.Items.Add(dr.GetValue(0))
        End While
        conn.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim emp_no As New Integer
        Dim emp_update As New Integer
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        Dim da As New OleDb.OleDbDataAdapter
        Dim dt As New DataTable
        conn.Open()
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Please ensure you have filled all the fields !!!", "ERROR !!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM EMP_NO", conn)
            da.SelectCommand = cmd
            da.Fill(dt)
            emp_no = dt.Rows(0).Item(0)
            emp_update = emp_no + 1
            Dim cmd1 As New OleDb.OleDbCommand("UPDATE EMP_NO SET NO=? WHERE NO=?", conn)
            cmd1.Parameters.AddWithValue("?", emp_update)
            cmd1.Parameters.AddWithValue("?", emp_no)
            cmd1.ExecuteNonQuery()
            Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO EMPLOYEE VALUES(?,?,?,?,?,?,?,?,?)", conn)
            cmd2.Parameters.AddWithValue("?", emp_no)
            cmd2.Parameters.AddWithValue("?", TextBox1.Text)
            cmd2.Parameters.AddWithValue("?", TextBox2.Text)
            cmd2.Parameters.AddWithValue("?", TextBox3.Text)
            cmd2.Parameters.AddWithValue("?", TextBox4.Text)
            cmd2.Parameters.AddWithValue("?", TextBox5.Text)
            cmd2.Parameters.AddWithValue("?", ComboBox1.Text)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            cmd2.Parameters.AddWithValue("?", TextBox1.Text)
            cmd2.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Employee Successfully Added !!! " + System.Environment.NewLine + "Employee No = " + String.Format(emp_no), "Employee Hired !!!", MessageBoxButtons.OK)
            Me.Hide()
            Branch_Manager_Menu.Show()
        End If
    End Sub
End Class