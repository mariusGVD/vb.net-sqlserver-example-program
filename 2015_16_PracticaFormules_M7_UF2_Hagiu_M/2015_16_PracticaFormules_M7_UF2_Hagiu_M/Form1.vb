Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
'H.M.C.
Public Class Form1
    ' recordo el ultim index del combobox
    Private PreviousComboBoxIndex As Integer = 0

    'Paramatres de la connexio
    Public cn As SqlConnection
    Dim strconexion As String = "Data Source=GGEVOD\SQLEXPRESS;" & "Initial Catalog=formules;" & "Integrated Security = True"

    'Boto per fer la connexio a la BD
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            ' Obrim la connexió PRINCIPAL (UNICA CONEXIO per treballa amb les dades (-1 conexio en funcio))
            cn = New SqlConnection(strconexion) 'creem nova conexio 
            cn.Open() 'obrim
            MessageBox.Show("La connexió a la BD s'ha realitzat amb exit!")
        Catch ex As Exception ' si hi ha un error treiem per pantalla el messagebox
            MessageBox.Show("Error al obrir la conexíó:" & vbCrLf & ex.Message)
            Exit Sub ' sortim
        End Try

        Dim sSel As String = "SELECT * FROM Elements"
        Dim da As New SqlDataAdapter(sSel, cn)
        Dim ds As New DataSet
        da.Fill(ds)
        ds.Tables(0).TableName = "Elements"

        'Omplo els combobox amb els elements de la BD
        ComboBox1.DataSource = ds.Tables(0)
        ComboBox1.DisplayMember = "nom_element" 'El campo nom_element se mostrara en el combo
        ComboBox1.ValueMember = "codi_element"

        ComboBox2.DataSource = ds.Tables(0)
        ComboBox2.DisplayMember = "nom_element" 'El campo nom_element se mostrara en el combo
        ComboBox2.ValueMember = "codi_element"

        ComboBox1.SelectedIndex = 0

        Console.WriteLine(ComboBox1.SelectedValue)

        ' desbloqueja els elements del Form1 despres de fer la connexió
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If ctrl.Name <> "Button1" And ctrl.Name <> "Button2" Then
                ctrl.Enabled = True
            End If
        Next
        
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Bloqueja els elements del Form1 fins que fem la connexió
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If ctrl.Name <> "Button1" And ctrl.Name <> "Button2" Then
                ctrl.Enabled = False
            End If
        Next

        ' Afegeix dades als Combobox del camp activaOno de la taula Formules
        Dim comboSource As New Dictionary(Of String, String)()
        comboSource.Add("1", "Si")
        comboSource.Add("0", "No")
        ComboBox3.DataSource = New BindingSource(comboSource, Nothing)
        ComboBox3.DisplayMember = "Value"
        ComboBox3.ValueMember = "Key"

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    End Sub

    Private Function DuplicateValue(key As String) As Boolean
        Dim Result As Boolean = False
        For Each rw As DataGridViewRow In DataGridView1.Rows
            If rw.Cells(0).Value = key Then
                Result = True
                Exit For
            End If
        Next
        Return Result
    End Function


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim frm2 As New Form2
        frm2.Show()

    End Sub

    'Boto per cercar avere si hi ha alguna formula amb el codi introduit
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (TextBox5.Text = "") Then
            MessageBox.Show("Introdueix el CODI de la Formula!!!")
        Else

            Dim comm As New SqlCommand()

            Dim codi_formula As New SqlParameter("@codi_formula", SqlDbType.VarChar)
            Dim nom_formula As New SqlParameter("@nom_formula", SqlDbType.VarChar)
            Dim data_creacio As New SqlParameter("@data_creacio", SqlDbType.VarChar)

            comm.Parameters.Add(codi_formula)
            comm.Parameters.Add(nom_formula)
            comm.Parameters.Add(data_creacio)

            codi_formula.Direction = ParameterDirection.Input
            nom_formula.Direction = ParameterDirection.Output
            data_creacio.Direction = ParameterDirection.Output

            nom_formula.Size = 50
            data_creacio.Size = 50


            codi_formula.Value = TextBox5.Text

            comm.Connection = cn
            comm.CommandText = "select @nom_formula=nom_formula,@data_creacio=data_creacio from Formules where codi_formula=@codi_formula"

            Try

                comm.ExecuteNonQuery()
                MessageBox.Show("La FORMULA amb el id: " + codi_formula.Value.ToString() + " esta creada el dia: " + data_creacio.Value.ToString() + " i s'anomena: " + nom_formula.Value.ToString())

            Catch ex As Exception
                Console.WriteLine(e.ToString)

                MessageBox.Show("Not Found")

            End Try


        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    End Sub

    'Boto del apartat Altes per afegir un element de la composicio de la formula al datagrid per despres fer el insert a la BD
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If TextBox3.Text = "" Then
            MessageBox.Show("Introdueix la quantitat en grams!!")
            Return
        End If

        ' Si el element existeix al datagrid no afegir-lo
        If VerificaSiElElementJaExisteixALaLlista(ComboBox1.SelectedValue) = 1 Then
            Return
        End If

        ' Selecciono nomes el element actula de la BD
        Dim sSel As String = "SELECT * FROM Elements WHERE codi_element = '" + ComboBox1.SelectedValue + "'"
        Console.WriteLine(sSel)
        Dim da As New SqlDataAdapter(sSel, cn)
        Dim ds As New DataSet
        da.Fill(ds)
        ds.Tables(0).TableName = "Elements"


        DataGridView1.ColumnCount = 4
        DataGridView1.Columns(0).Name = "ID Element"
        DataGridView1.Columns(1).Name = "Nom"
        DataGridView1.Columns(2).Name = "Data creació"
        DataGridView1.Columns(3).Name = "Quantitat grams"

        For Each fila As DataRow In ds.Tables(0).Rows
            Dim codi_element As String = fila.Field(Of String)("codi_element")
            Dim nom_element As String = fila.Field(Of String)("nom_element")
            Dim data_creacio As Date = fila.Field(Of Date)("data_creacio")
            Dim quantitat_grams As Integer = TextBox3.Text

            Dim row As String() = New String() {codi_element, nom_element, data_creacio, quantitat_grams}
            DataGridView1.Rows.Add(row)
            TextBox3.Text = ""
        Next

    End Sub

    'Boto per eliminar la fila seleccionada del datagrid1
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.DataGridView1.Rows.RemoveAt(Me.DataGridView1.CurrentRow.Index)
    End Sub

    'Boto per borrar la formula buscada pel codi_formula. Tambe es borra tots els element que pertanyen a la composicio de la formula
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim comm1 = New SqlCommand()
        Dim comm2 = New SqlCommand()

        comm1.Connection = cn
        comm2.Connection = cn

        comm1.Parameters.AddWithValue("@codi_formula", TextBox4.Text)
        comm1.CommandText = "DELETE FROM Composicio WHERE codi_formula = @codi_formula"

        comm2.Parameters.AddWithValue("@codi_formula", TextBox4.Text)
        comm2.CommandText = "DELETE FROM Formules WHERE codi_formula = @codi_formula"

        Try
                comm1.ExecuteNonQuery()
                comm2.ExecuteNonQuery()
                MessageBox.Show("La formula ha estat ELIMINADA....")

            TextBox4.Clear()

        Catch ex As SqlException

            MessageBox.Show("No borrat....")

        End Try

    End Sub

    'Boto per desactivar una formula per el codi_formula introduit
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim comm = New SqlCommand()

        comm.Connection = cn

        comm.Parameters.AddWithValue("@codi_formula", TextBox4.Text)
        comm.CommandText = "UPDATE Formules SET activaOno = 0 WHERE codi_formula=@codi_formula;"

        Try
            comm.ExecuteNonQuery()
            If (comm.ExecuteNonQuery() <> 0) Then

                MessageBox.Show("La formula ha estat desactivada....")

            Else
                MessageBox.Show("La formula introduida no existeix!!!")
            End If
            TextBox4.Clear()
        Catch ex As Exception

            MessageBox.Show("No s'ha pogut desactivar....")

        End Try
    End Sub

    'Boto per tancar el programa
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    'Boto per buscar formules a les qual pertanyen un element buscat pel codi_element
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        DataGridView3.DataSource = Nothing
        If (TextBox10.Text = "") Then

            MessageBox.Show("Introdueix el CODI del Element!!!")

        Else

            Dim sSel As String = "SELECT Elements.nom_element AS Element, Formules.nom_formula AS Formula, Formules.data_creacio AS Data, Formules.totalPes_grams AS Pes from Formules INNER JOIN Composicio ON Formules.codi_formula = Composicio.codi_formula INNER JOIN Elements ON Composicio.codi_element = Elements.codi_element WHERE Elements.codi_element = '" + TextBox10.Text + "'"
            Dim da As New SqlDataAdapter(sSel, cn)
            Dim ds As New DataSet
            da.Fill(ds)
            ds.Tables(0).TableName = "Formules"

            DataGridView3.DataSource = ds.Tables(0)
            TextBox10.Text = ""


        End If
    End Sub

    'Boto per llista les formules inactives
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DataGridView3.DataSource = Nothing
        Dim sSel As String = "SELECT codi_formula AS ID, nom_formula AS Formula, data_creacio AS Data, totalPes_grams AS Pes FROM Formules where activaOno = '0'"

        Try
            
            Dim da As New SqlDataAdapter(sSel, cn)
            Dim ds As New DataSet
            da.Fill(ds)
            ds.Tables(0).TableName = "Formules"

            DataGridView3.DataSource = ds.Tables(0)

        Catch ex As Exception

            MessageBox.Show("No hi han formules INACTIVES!!")

        End Try

    End Sub

    'Boto per llista formules creades en una data determinada
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        DataGridView3.DataSource = Nothing
        Dim sSel As String = "select codi_formula AS ID, nom_formula AS Formula, data_creacio AS Data, totalPes_grams AS Pes, activaOno AS Activa from Formules where data_creacio = '" + DateTimePicker3.Value.ToString("yyyy-MM-dd") + "'"
        Dim da As New SqlDataAdapter(sSel, cn)
        Dim ds As New DataSet
        da.Fill(ds)
        ds.Tables(0).TableName = "Formules"

        DataGridView3.DataSource = ds.Tables(0)
    End Sub

    'Boto per cercar la formula i carregar les dades per poder fer el update despres
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        DataGridView2.DataSource = Nothing
        If (TextBox7.Text = "") Then
            MessageBox.Show("Introdueix el CODI de la Formula!!!")

        Else

            Dim commando As SqlCommand = New SqlCommand("SELECT codi_formula, nom_formula, data_creacio, activaOno FROM Formules WHERE codi_formula ='" + TextBox7.Text + "'")
            commando.Connection = cn
            Dim reader As SqlDataReader
            reader = commando.ExecuteReader()

            While (reader.Read())
                If (TextBox7.Text = reader("codi_formula").ToString()) Then
                    TextBox6.Text = reader("codi_formula").ToString()
                    TextBox8.Text = reader("nom_formula").ToString()
                    TextBox9.Text = reader("activaOno").ToString()
                    DateTimePicker2.Value = reader("data_creacio")
                Else
                    MessageBox.Show("La formula no existeix!!")
                End If

            End While
            commando.Connection.Close()

            Update_DataGridView2()

        End If

    End Sub

    'Boto per afegir element a la composicio de la formula. fa un insert a la BD cridan el procediment Sub AfegeixElementFormula()
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim codi_formula As String = TextBox7.Text
        Dim quantitat_grams As String = TextBox12.Text
        Dim codi_element As String = ComboBox2.SelectedValue
        AfegeixElementALaFormula(codi_formula, quantitat_grams, codi_element)

        ' Actualitza la suma a la composicio de la formula
        Actualitza_pesTotal(codi_formula)
        'Actualitza les dades del datagrid2
        Update_DataGridView2()
    End Sub

    'Boto per borrar els element de la composicio de la formula. fa directament un DELETE a la BD
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim codi_formula As String = TextBox7.Text

        Dim selectedRowCount As Integer = DataGridView2.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount = 0 Then
            MessageBox.Show("Selecciona tota la FILA no nomes una cel·lula!!")
            Return
        ElseIf selectedRowCount > 1 Then
            MessageBox.Show("Selecciona UNA SOLA fila!!!")
            Return
        End If

        Dim data As String = DataGridView2.SelectedRows(0).Cells(0).Value.ToString()
        Dim comm1 = New SqlCommand()
        comm1.Connection = cn

        comm1.CommandText = "DELETE FROM Composicio WHERE codi_formula = '" + codi_formula + "' AND codi_element = '" + data + "'"

        Try
            comm1.ExecuteNonQuery()
            MessageBox.Show("Borrat....")
            Update_DataGridView2()

        Catch ex As Exception

            MessageBox.Show("No borrat....")

        Finally

        End Try
        'Accualitza la suma dels elements de la composicio de la formula
        Actualitza_pesTotal(codi_formula)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Introdueix nomes numeros!!!")
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Introdueix nomes numeros!!!")
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        TextBox9.MaxLength = 1
        Dim k As Byte = Asc(e.KeyChar)

        If Not (e.KeyChar = "0" Or e.KeyChar = "1" Or k = 8 Or k = 13) Then e.Handled = True
    End Sub

    'Boto per fer el update de les dades de la formula que vull modificar
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim codi_formula As String = TextBox6.Text
        Dim nom_formula As String = TextBox8.Text
        Dim activa As String = TextBox9.Text
        Dim data As String = DateTimePicker2.Value.ToString("yyyy-MM-dd")

        Dim comm = New SqlCommand()

        comm.Connection = cn

        comm.Parameters.AddWithValue("@codi_formula", codi_formula)
        comm.Parameters.AddWithValue("@nom_formula", nom_formula)
        comm.Parameters.AddWithValue("@activa", activa)
        comm.Parameters.AddWithValue("@data", data)
        comm.CommandText = "UPDATE Formules SET nom_formula = @nom_formula, activaOno = @activa, data_creacio = @data, totalPes_grams = (SELECT SUM(quantitat_grams) FROM Composicio WHERE codi_formula = @codi_formula) WHERE codi_formula=@codi_formula;"

        Try
            comm.ExecuteNonQuery()
            MessageBox.Show("La Formula s'ha actualitzat correctament!")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    'Boto per fer un insert d'una nova formula a la BD
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim codi_formula As String = TextBox1.Text
        Dim nom_formula As String = TextBox2.Text
        Dim data As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        Dim activa As String = ComboBox3.SelectedValue.ToString()

        If codi_formula = "" Then
            MessageBox.Show("Introdueix el CODI de la FORMULA!!")
            Return
        End If

        If nom_formula = "" Then
            MessageBox.Show("Introdueix el NOM de la Formula!!!")
            Return
        End If

        If data = "" Then
            MessageBox.Show("Selecciona la Data de creacio de la Formula!!!")
            Return
        End If

        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("La Formula no te elements a la composicio! Introdueix els element de la composicio!!")
            Return
        End If

        If (VerificaSiExisteixFormula(codi_formula) = 1) Then
            MessageBox.Show("El codi de Formula introduida ja existeix o pertany a una altra Formula!!")
            Return
        End If

        AfegeixFormula(codi_formula, nom_formula, data, activa)
        AfegeixElements(codi_formula)

        Actualitza_pesTotal(codi_formula)

        ' Neteja formulari
        TextBox1.Text = ""
        TextBox2.Text = ""
        DataGridView1.Rows.Clear()

        MessageBox.Show("La formula ha estat guardada! Introdueix una altra Formula.")
    End Sub

    


    'FUNCCIONS

    'Procediment Sub per introduir els elements afegits al datagrid1 a la BD
    Sub AfegeixElements(codi_formula)
        Dim cellValues As New List(Of Object)

        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            AfegeixElementALaFormula(codi_formula, DataGridView1(3, i).Value, DataGridView1(0, i).Value)
        Next

    End Sub

    'Funcio per verificar si ja existeix el element a la llista
    Function VerificaSiElElementJaExisteixALaLlista(element As String)
        Dim exista = 0

        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1(0, i).Value = element Then
                exista = 1
            End If
        Next

        Return exista
    End Function

    'Procediment per fer el insert de les dades introduides de la formula
    Sub AfegeixFormula(codi_formula As String, nom_formula As String, data As String, activa As String)

        Dim comm = New SqlCommand()
        comm.Connection = cn

        comm.Parameters.AddWithValue("@codi_formula", codi_formula)
        comm.Parameters.AddWithValue("@nom_formula", nom_formula)
        comm.Parameters.AddWithValue("@data", data)
        comm.Parameters.AddWithValue("@activa", activa)
        comm.CommandText = "INSERT INTO Formules (codi_formula, nom_formula, data_creacio, activaOno) VALUES (@codi_formula, @nom_formula, @data, @activa)"

        Try
            comm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine("Eo")
        End Try
    End Sub

    'Funcio per verificar si existeix la Formula 
    Function VerificaSiExisteixFormula(codi_formula As String)
        Dim exista As String = 0
        cn.Close()

        cn = New SqlConnection(strconexion) 'creem nova conexio 
        cn.Open() 'obrim

        Dim comm = New SqlCommand()
        comm.Connection = cn

        comm.Parameters.AddWithValue("@codi_formula", codi_formula)
        comm.CommandText = "SELECT COUNT(codi_formula) AS count FROM Formules where codi_formula = @codi_formula"

        Try
            exista = comm.ExecuteScalar()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return exista
    End Function

    'Procediment per actualizar el pes total de tots els elements de la composicio de la formula(UPDATE)
    Private Sub Actualitza_pesTotal(codi_formula As String)
        cn.Close()

        cn = New SqlConnection(strconexion) 'creem nova conexio 
        cn.Open() 'obrim

        Dim comm = New SqlCommand()
        comm.Connection = cn

        comm.Parameters.AddWithValue("@codi_formula", codi_formula)
        comm.CommandText = "UPDATE Formules SET totalPes_grams = (SELECT SUM(quantitat_grams) FROM Composicio WHERE codi_formula = @codi_formula) WHERE codi_formula=@codi_formula;"

        Try
            comm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    'Procediment per afegir element a la composicio de la formula(apartat Modificacions) UPDATE si existeix si no INSERT
    Sub AfegeixElementALaFormula(codi_formula, quantitat_grams, codi_element)
        Dim query As String = String.Empty
        query &= "begin tran if exists (   select *      from Composicio with (updlock,serializable)      where codi_element = @codi_element       and codi_formula = @codi_formula   )   begin     update Composicio       set quantitat_grams = @quantitat_grams       where codi_element = @codi_element         and codi_formula = @codi_formula   end else   begin     insert into Composicio (codi_formula, codi_element, quantitat_grams)       values (@codi_formula, @codi_element, @quantitat_grams);   end commit tran"

        Using cn
            Using comm As New SqlCommand()
                With comm
                    .Connection = cn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@codi_formula", codi_formula)
                    .Parameters.AddWithValue("@codi_element", codi_element)
                    .Parameters.AddWithValue("@quantitat_grams", quantitat_grams)
                End With
                'cn.Open()
                comm.ExecuteNonQuery()
            End Using
        End Using


        ' Actualitza el pes total de la composicio de la formula
        Actualitza_pesTotal(codi_formula)
    End Sub

    'Funcio per refrescar i tornar a mostrar les dades al datagrid2
    Private Function Update_DataGridView2()
        cn.Close()

        cn = New SqlConnection(strconexion) 'creem nova conexio 
        cn.Open() 'obrim

        Dim sSel As String = "SELECT Elements.codi_element AS ID, Elements.nom_element AS Element, Elements.data_creacio AS Data, Composicio.quantitat_grams AS Quantitat from Elements INNER JOIN Composicio ON Elements.codi_element = Composicio.codi_element INNER JOIN Formules ON Formules.codi_formula = Composicio.codi_formula WHERE Formules.codi_formula = '" + TextBox7.Text + "'"
        Dim da As New SqlDataAdapter(sSel, cn)
        Dim ds As New DataSet
        da.Fill(ds)
        ds.Tables(0).TableName = "Elements"

        DataGridView2.DataSource = ds.Tables(0)
    End Function

    'Funcio per refrescar els combobox1 i 2 al afegir un nou element
    Public Function Update_Combobox1_2()
        cn.Close()

        cn = New SqlConnection(strconexion) 'creem nova conexio 
        cn.Open() 'obrim

        Dim sSel As String = "SELECT * FROM Elements"
        Dim da As New SqlDataAdapter(sSel, cn)
        Dim ds As New DataSet
        da.Fill(ds)
        ds.Tables(0).TableName = "Elements"

        ComboBox1.DataSource = ds.Tables(0)
        ComboBox1.DisplayMember = "nom_element" 'El campo nom_element se mostrara en el combo
        ComboBox1.ValueMember = "codi_element"

        ComboBox2.DataSource = ds.Tables(0)
        ComboBox2.DisplayMember = "nom_element" 'El campo nom_element se mostrara en el combo
        ComboBox2.ValueMember = "codi_element"
    End Function
End Class
