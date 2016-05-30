Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Data.SqlClient



Public Class Form2
    Public cnn As SqlConnection
    Public comm As SqlCommand
    Dim strconexion As String = "Data Source=GGEVOD\SQLEXPRESS;" & "Initial Catalog=formules;" & "Integrated Security = True"
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        cnn = New SqlConnection(strconexion) 'creem nova conexio 
        Dim comm As New SqlCommand
        'obrim
        cnn.Open()

        comm.Connection = cnn

        'If Me.TextBox1.Text = "" Then
        '    MessageBox.Show("Introdueix el CODI del Element")
        'Else
        '    Try
        '        'comm.CommandText = "INSERT INTO Elements(codi_element, nom_element, data_element) " & _
        '        '                " VALUES(" & Me.TextBox1.Text & ",'" & Me.TextBox2.Text & "','" & _
        '        '                Me.DateTimePicker1.Text & "')"

        '        comm.CommandText = "INSERT INTO Elements(codi_element, nom_element, data_element) " & _
        '                        " VALUES(" & Me.TextBox1.Text & ",'" & Me.TextBox2.Text & "','" & _
        '                        Me.DateTimePicker1.Text & "')"
        '        comm.ExecuteNonQuery()
        '    Catch ex As Exception

        '        MessageBox.Show("Error no s'ha guardat!!!")

        '    Finally
        '        cnn.Close()
        '    End Try
        'End If

        'creating instance of SsqlParameter

        Dim codi_element As New SqlParameter("@codi_element", SqlDbType.VarChar)
        Dim nom_element As New SqlParameter("@nom_element", SqlDbType.VarChar)
        Dim data_creacio As New SqlParameter("@data_creacio", SqlDbType.Date)

        'Adding parameter to SqlCommand

        comm.Parameters.Add(codi_element)
        comm.Parameters.Add(nom_element)
        comm.Parameters.Add(data_creacio)

        'Setting values 

        codi_element.Value = TextBox1.Text
        nom_element.Value = TextBox2.Text
        data_creacio.Value = DateTimePicker1.Text



        ' adding connection to SqlCommand
        comm.Connection = cnn
        ' Sql Statement
        comm.CommandText = "insert into Elements values(@codi_element,@nom_element,@data_creacio)"

        Try
            If TextBox1.Text = "" Then
                MessageBox.Show("Introdueix el CODI del Element!")

            Else
                comm.ExecuteNonQuery()
                MessageBox.Show("L'element introduit s'ha guardat a la BD.")
                TextBox1.Clear()
                TextBox2.Clear()
                DateTimePicker1.Refresh()

                Call Form1.Update_Combobox1_2()
            End If

        Catch ex As Exception

            MessageBox.Show("Error no s'ha guardat!!!")

        Finally
            cnn.Close()
        End Try





    End Sub
End Class