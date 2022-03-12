# tabla-registro

Private Sub btnnuevo_Click()
    configuracion.Cells(1, 2) = configuracion.Cells(1, 2) + 1
    configuracion.Cells(2, 2) = configuracion.Cells(1, 2)
    activar_controles (True)
    txtnombre.SetFocus
End Sub


Private Sub btnguardar_Click()
    nombre = txtnombre.Text
    fila = configuracion.Cells(1, 2)
    datos.Cells(fila, 1) = txtnombre.Text
    datos.Cells(fila, 2) = txtcedula.Text
    datos.Cells(fila, 3) = txtprograma.Text
    activar_controles (False)
    limpiar
    MsgBox "datos guardados con exito"
End Sub


Private Sub btnbuscar_Click()
    frmbuscar.Show
    
End Sub

Private Sub btnbuscar_Click()
    i = 3
    ultimo = configuracion.Cells(2, 2)
    encontrado = False
    While i <= ultimo And encontrado = False
        If datos.Cells(i, 2) = frmbuscar.txtbusqueda.Text Then
           encontrado = True
        Else
            i = i + 1
        End If
    Wend
    If encontrado Then
        lblnuevo.txtnombre.Text = datos.Cells(i, 1)
        lblnuevo.txtcedula.Text = datos.Cells(i, 2)
        lblnuevo.txtprograma.Text = datos.Cells(i, 3)
        configuracion.Cells(1, 2) = i
        frmbuscar.Hide
    Else
        MsgBox "Nada... no encontrado"
    End If
End Sub

Private Sub btnimagen_Click()
    archivo = Application.GetOpenFilename("imagenes(*.jpg;*.bmp), *.jpg;.bmp")
    imagen.Picture = LoadPicture(archivo)
    datos.Cells(configuracion.Cells(1, 2), 4) = archivo
End Sub


Private Sub btnconfig_Click()
Application.Visible = True
End Sub


Private Sub btnborrar_Click()
    datos.Rows(configuracion.Cells(1, 2)).EntireRow.Delete
End Sub

Private Sub btneditar_Click()
    activar_controles (True)
End Sub

