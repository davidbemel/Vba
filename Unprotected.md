Cerrar libro Excel (guardar cambios)

ActiveWorkbook.Close

ActiveWorkbook.Close Savechanges:=True

ActiveWorkbook.Close(True)

Cerrar libro Excel (sin guardar cambios)

ActiveWorkbook.Close(False)

ActiveWorkbook.Close Savechanges:=False

Cerrar libro Excel (variable, sin guardar cambios)

Application.DisplayAlerts = False

Windows(Libro_mayor).Close

Application.DisplayAlerts = True

Abrir libro Excel (ruta fija)

Workbooks.Open FileName:=”C:\Trabajo\Informe.xls”

Abrir libro Excel (diálogo)

Msg = MsgBox(“Elija archivo para abrir.”, vbOKOnly, (“”))

strArchivo = Application.GetOpenFilename

On Error GoTo 99

Workbooks.OpenText Filename: = strArchivo

If strArchivo = “” Then Exit Sub

strArchivo = ActiveWindow.Caption

99:

Exit sub

Devolver nombre del libro Excel

strNombre = ActiveSheet.Parent.FullName

MsgBox ActiveWorkbook.FullName

--------

1) Ocultar una hoja de Excel (para mostrarla, hay que cambiar "False" por "True".

Sub OcultarHoja1()

   Sheets("Hoja1").Visible=False

End Sub

2) Ocultar totalmente una hoja de Excel (no se podrá mostrar con el menú contextual que emerge de las pestañas de hojas).

Sub OcultarHoja1Totalmente()

   Sheets("Hoja1").Visible=xlVeryHidden

End Sub

PROTECCIÓN DE HOJAS DE EXCEL CON VBA

3) Proteger la hoja activa de Excel (no se podrán seleccionar las celdas bloqueadas).

Sub ProtegerHoja()

   ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

   ActiveSheet.EnableSelection = xlUnlockedCells

End Sub

4) Desproteger una hoja de Excel.

Sub DesprotegerHoja()

   ActiveSheet.Unprotect

End Sub

VARIAS ACCIONES CON HOJAS DE EXCEL CON VBA

5) Seleccionar una hoja de Excel (el nombre debe ser el que aparece en la pestaña de la hoja).

Sub SeleccionarHoja1()

   Sheets("Hoja1").Select

End Sub

6) Cambiar el zoom de una hoja de Excel.

Sub Zoom120()

   ActiveWindow.Zoom = 120

End Sub

7) Imprimir las hojas seleccionadas con la impresora y configuraciones por defecto.

Sub ImprimirHojaActiva()

   ActiveWindow.SelectedSheets.PrintOut copies:=1, collate:=True

End Sub

8) Ocultar la cuadrícula de la hoja activa (para mostrarlos, cambiar "False" por "True").

Sub OcultarCuadricula()

   ActiveWindow.DisplayGridlines = False

End Sub

9) Ocultar los títulos de encabezamiento de filas y columnas (para mostrarlos, cambiar "False" por "True").

Sub OcultarTitulos()

   ActiveWindow.DisplayHeadings = False

End Sub

SUPRIMIR HOJA

Application.DisplayAlerts = False

For i = 1 To Sheets.Count

  Sheets(i).Activate

    xxx = ActiveCell.Worksheet.Name

      If xxx = "Informe" Then

        ActiveWindow.SelectedSheets.Delete

      End If

Next

Application.DisplayAlerts = True

10) Mostrar las hojas del libro

# 11 Desasegurar

Sub breakit()

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
On Error Resume Next
    For i = 65 To 66
      For j = 65 To 66
        For k = 65 To 66
          For l = 65 To 66
            For m = 65 To 66
              For i1 = 65 To 66
                For i2 = 65 To 66
                  For i3 = 65 To 66
                    For i4 = 65 To 66
                      For i5 = 65 To 66
                        For i6 = 65 To 66
                          For n = 32 To 126
                             
                              ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
                                          Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                          Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                         
                                If ActiveSheet.ProtectContents = False Then
                                   MsgBox "One usable password is " & Chr(i) & Chr(j) & _
                                   Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) _
                                   & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                             
                                   Exit Sub
                                End If
                           Next
                         Next
                       Next
                     Next
                   Next
                 Next
               Next
             Next
          Next
        Next
      Next
    Next

End Sub

	
	
	
