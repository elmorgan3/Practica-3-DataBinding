Imports System.Data.SqlClient

Public Class Form1
    Private con As SqlConnection
    Private dts As DataSet
    Private ada As SqlDataAdapter

    Private bmb As BindingManagerBase

    Dim cambioPermitido = True

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Abrimos la conexion con la base de datos
        con = New SqlConnection
        con.ConnectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=COMPASSTRAVEL; Trusted_Connection=True;"
        con.Open()

        ' Creamos un adapter para hablar con la tabla de CONTACTO
        ada = New SqlDataAdapter("select * from MEMBERS order by MEMBERID", con)

        ' Generamos las sentencias de insert, update, delete, si no hacemos esto no funcionara el update del adapter
        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(ada)

        ' Creamos el dataset, que es la base de datos virtual, con la que trabajaremos cargada en memoria
        dts = New DataSet()

        ' Ejecutamos el adapter y cargamos en memoria la información y la estructura de la tabla
        ada.Fill(dts, "Members")


        AddHandler dts.Tables("Members").ColumnChanging, New DataColumnChangeEventHandler(AddressOf OnColumnChanging)
        AddHandler dts.Tables("Members").RowChanging, New DataRowChangeEventHandler(AddressOf OnRowChanging)

        ConstruirDataBinding()


    End Sub

    Private Sub ConstruirDataBinding()
        Dim oBind As Binding

        'los bildng tienen tres propiedades la propiedad a la que haremos referencia,
        'el DataSet, y la columna de la base datos cargada en memoria
        oBind = New Binding("Text", dts, "Members.MemberId")
        txtId.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.FirstName")
        txtFirstName.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.LastName")
        txtLastName.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.Address1")
        txtAddress1.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.Address2")
        txtAddress2.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.State")
        txtState.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.PostalCode")
        txtPostalCode.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.Country")
        txtCountry.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.UserName")
        txtUserName.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.Password")
        txtPassword.DataBindings.Add(oBind)
        oBind = Nothing


        oBind = New Binding("Text", dts, "Members.MemberId")
        TextBoxBuscar.DataBindings.Add(oBind)
        oBind = Nothing


        bmb = BindingContext(dts, "Members")

    End Sub

    '------------------------------------------------------
    'Este metodo se actoba cuando el foco cambia de COLUMNA
    '------------------------------------------------------
    Private Sub OnColumnChanging(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        'Este IF sirve para controlar que la contraseña tiene una longitud correctta
        If (e.Column.ColumnName = "PASSWORD") Then
            If txtPassword.Text.Count < 4 Then
                MsgBox("El campo de la contraseña debe tener como minimo 4 díjitos.")
                Throw New System.Exception("error")
                txtPassword.Select()
                'Throw New System.Exception("error")
                'e.Row.EndEdit()

                'El Throw New System.Exception("error") hace que se pare la ejecucion


            ElseIf txtPassword.Text.Count >= 4 And txtPassword.Text.Count <= 8 Then
                MsgBox("El campo de la contraseña, se recomienda que tenga un mínimo de 9 dijitos pero se acepta.")

            End If
        End If


        'Este IF comprueba que no exista ya este user name
        If (e.Column.ColumnName = "USERNAME") Then
            Dim userNameBuscada
            Dim userNameExistente
            Dim encontrado = False


            userNameBuscada = txtUserName.Text

            For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
                userNameExistente = dts.Tables("Members").Rows(index)("USERNAME").ToString()
                If userNameBuscada.Equals(userNameExistente) Then
                    encontrado = True
                    If encontrado = True Then
                        MsgBox("Este nombre de usuario ya existe")

                        Throw New System.Exception("error")
                    End If
                End If
            Next




        End If

    End Sub

    '------------------------------------------------------
    'Este metodo se actoba cuando el foco cambia de FILA
    '------------------------------------------------------
    Private Sub OnRowChanging(ByVal sender As Object, ByVal e As DataRowChangeEventArgs)

        'If txtPassword.Text = "" Then


        '    Throw New System.Exception("error")

        '    'El Throw New System.Exception("error") hace que se pare la ejecucion


        'End If
        'HacerPersistencia()
    End Sub


    Private Sub ButtonNew_Click(sender As Object, e As EventArgs) Handles ButtonNew.Click
        ButtonNew.Visible = False
        ButtonUpdate.Visible = False
        ButtonDelete.Visible = False
        ButtonPrimero.Enabled = False
        ButtonAnterior.Enabled = False
        ButtonSiguiente.Enabled = False
        ButtonUltimo.Enabled = False


        ButtonAceptarCambios.Visible = True
        ButtonCancelar.Visible = True
        ButtonSugerirNombre.Visible = True

        'txtId.ReadOnly = False
        txtFirstName.ReadOnly = False
        txtLastName.ReadOnly = False
        txtAddress1.ReadOnly = False
        txtAddress2.ReadOnly = False
        txtState.ReadOnly = False
        txtPostalCode.ReadOnly = False
        txtCountry.ReadOnly = False
        txtUserName.ReadOnly = False
        txtPassword.ReadOnly = False

        bmb.AddNew()

        'txtId.Text = bmb.Count
    End Sub

    Private Sub ButtonUpdate_Click(sender As Object, e As EventArgs) Handles ButtonUpdate.Click
        ButtonNew.Visible = False
        ButtonUpdate.Visible = False
        ButtonDelete.Visible = False
        ButtonPrimero.Enabled = False
        ButtonAnterior.Enabled = False
        ButtonSiguiente.Enabled = False
        ButtonUltimo.Enabled = False

        ButtonAceptarCambios.Visible = True
        ButtonCancelar.Visible = True
        ButtonSugerirNombre.Visible = True

        'txtId.ReadOnly = False
        txtFirstName.ReadOnly = False
        txtLastName.ReadOnly = False
        txtAddress1.ReadOnly = False
        txtAddress2.ReadOnly = False
        txtState.ReadOnly = False
        txtPostalCode.ReadOnly = False
        txtCountry.ReadOnly = False
        txtUserName.ReadOnly = False
        txtPassword.ReadOnly = False
    End Sub

    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click
        Select Case MsgBox("¿Estas seguro que quieres eliminar este registro?", MsgBoxStyle.YesNoCancel, "caption")
            Case MsgBoxResult.Yes
                bmb.RemoveAt(bmb.Position)
                HacerPersistencia()
            Case MsgBoxResult.Cancel

            Case MsgBoxResult.No

        End Select
    End Sub

    Private Sub ButtonAceptarCambios_Click(sender As Object, e As EventArgs) Handles ButtonAceptarCambios.Click
        bmb.EndCurrentEdit()

        ButtonNew.Visible = True
        ButtonUpdate.Visible = True
        ButtonDelete.Visible = True
        ButtonPrimero.Enabled = True
        ButtonAnterior.Enabled = True
        ButtonSiguiente.Enabled = True
        ButtonUltimo.Enabled = True

        ButtonAceptarCambios.Visible = False
        ButtonCancelar.Visible = False
        ButtonSugerirNombre.Visible = False

        'txtId.ReadOnly = True
        txtFirstName.ReadOnly = True
        txtLastName.ReadOnly = True
        txtAddress1.ReadOnly = True
        txtAddress2.ReadOnly = True
        txtState.ReadOnly = True
        txtPostalCode.ReadOnly = True
        txtCountry.ReadOnly = True
        txtUserName.ReadOnly = True
        txtPassword.ReadOnly = True

        HacerPersistencia()

    End Sub

    Private Sub ButtonCancelar_Click(sender As Object, e As EventArgs) Handles ButtonCancelar.Click
        bmb.CancelCurrentEdit()

        ButtonNew.Visible = True
        ButtonUpdate.Visible = True
        ButtonDelete.Visible = True
        ButtonPrimero.Enabled = True
        ButtonAnterior.Enabled = True
        ButtonSiguiente.Enabled = True
        ButtonUltimo.Enabled = True

        ButtonAceptarCambios.Visible = False
        ButtonCancelar.Visible = False
        ButtonSugerirNombre.Visible = False


        'txtId.ReadOnly = True
        txtFirstName.ReadOnly = True
        txtLastName.ReadOnly = True
        txtAddress1.ReadOnly = True
        txtAddress2.ReadOnly = True
        txtState.ReadOnly = True
        txtPostalCode.ReadOnly = True
        txtCountry.ReadOnly = True
        txtUserName.ReadOnly = True
        txtPassword.ReadOnly = True
    End Sub

    Private Sub ButtonPrimero_Click(sender As Object, e As EventArgs) Handles ButtonPrimero.Click
        bmb.Position = 0

    End Sub

    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        bmb.Position = bmb.Position - 1

    End Sub

    Private Sub ButtonSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonSiguiente.Click
        bmb.Position = bmb.Position + 1

    End Sub

    Private Sub ButtonUltimo_Click(sender As Object, e As EventArgs) Handles ButtonUltimo.Click
        bmb.Position = bmb.Count - 1

    End Sub

    '-----------------------------------------------------------------------
    'Metodo que guarda los cambios, en la tabla de la base de datos original 
    '-----------------------------------------------------------------------
    Private Sub HacerPersistencia()
        ' METODO 1. Actualitzar toda la tabla
        'ada.Update(ds, "CONTACTE")

        ' METODO 2. Actualitzar solo los cambios, es mas eficiente
        Dim dt As DataTable
        dt = dts.Tables("Members").GetChanges()
        If IsNothing(dt) Then
            'MsgBox("No hay nada que guardar")
        Else
            ada.Update(dt)
        End If

        dts.Tables("Members").AcceptChanges()

    End Sub

    Private Sub txtId_KeyDown(sender As Object, e As KeyEventArgs) Handles txtId.KeyDown

        If e.Alt And e.KeyCode = Keys.Up Then
            bmb.Position = 0
        End If

        If e.Alt And e.KeyCode = Keys.Left Then
            bmb.Position = bmb.Position - 1
        End If

        If e.Alt And e.KeyCode = Keys.Right Then
            bmb.Position = bmb.Position + 1
        End If

        If e.Alt And e.KeyCode = Keys.Down Then
            bmb.Position = bmb.Count - 1
        End If
    End Sub

    Private Sub ButtonBuscar_Click(sender As Object, e As EventArgs) Handles ButtonBuscar.Click
        Dim idBuscada
        Dim idExistente
        Dim encontrado = False

        If TextBoxBuscar.Text = "" Then
            MsgBox("El campo de busqueda esta vacio.")
        Else
            idBuscada = TextBoxBuscar.Text

            For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
                idExistente = dts.Tables("Members").Rows(index)("MEMBERID").ToString()
                If idBuscada.Equals(idExistente) Then
                    encontrado = True
                    bmb.Position = index

                End If
            Next
        End If


        If encontrado = False Then
            MsgBox("Esta ID, no existe.")
        End If
    End Sub

    Private Sub ButtonSugerirNombre_Click(sender As Object, e As EventArgs) Handles ButtonSugerirNombre.Click
        Dim nombre As String = txtFirstName.Text.Substring(0, 3)
        Dim apellido As String = txtLastName.Text.Substring(txtLastName.TextLength - 3)
        Dim username As String = nombre & apellido

        Dim result As Integer = MessageBox.Show("¿Quieres usar " & username & " como nombre de usuario ?", "Sugerencia", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            'MessageBox.Show("Cancel pressed")
        ElseIf result = DialogResult.No Then
            'MessageBox.Show("No pressed")
        ElseIf result = DialogResult.Yes Then
            txtUserName.Text = username
            txtUserName.Select()
        End If
    End Sub
End Class
