Imports System
Imports System.Windows.Forms
Imports Word = Microsoft.Office.Interop.Word
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dialogo As New SaveFileDialog()
        dialogo.Filter = "Documentos de Word (*.docx)|*.docx" ' Filtro para seleccionar solo archivos de Word
        dialogo.DefaultExt = ".docx" ' Extensión predeterminada
        dialogo.AddExtension = True ' Agregar automáticamente la extensión si no se proporciona

        If dialogo.ShowDialog() = DialogResult.OK Then
            Dim ruta As String = dialogo.FileName ' Obtener la ruta seleccionada por el usuario
            Dim dato As String = txtDato.Text

            Try
                Dim wordApp As New Word.Application()
                wordApp.Visible = True
                Dim doc As Word.Document = wordApp.Documents.Add()
                doc.Content.Text = dato
                doc.SaveAs2(ruta)
                doc.Close()

                MessageBox.Show("Documento guardado exitosamente.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error al guardar el documento:" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
End Class
