Option Strict Off
Imports System.IO
Public Class Form1
    Dim directorio As String = My.Application.Info.DirectoryPath
    'SET,$PDOBL$§PTOS_FRANQUICIA_PORTON_SECCIONES_MPE&NET&ZFKA!ZK4!FKB&L&MX1!MX2&T19&D34&W50&IN1!IN2!IN3§,8
    Dim fila As String
    Dim Entrada As String = directorio & "\CodigosComercialesPorCategoriaTab.txt"
    Dim trozo(3) As String
    Dim nivelEnsamblado(3) As String
    Dim valorAtributo(9) As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        inicializaArays()
        TsetInicial.Focus()
        visualizarCodigos()
        TsetInicial.Focus()
    End Sub

    Private Sub cambioTexto(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tlevel.TextChanged, tFunction.TextChanged, tAligment.TextChanged
        componerSet()
    End Sub

    Private Sub cambioChek(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cVariantesPais.SelectedIndexChanged, cTransmision.SelectedIndexChanged, cTipo.SelectedIndexChanged, cNoAsignado.SelectedIndexChanged, cNivelCarga.SelectedIndexChanged, cMontajeBruto.SelectedIndexChanged, cConduccion.SelectedIndexChanged
        componerSet()
    End Sub

    Private Sub componerSet()
        Dim i As Integer = 0
        Dim cruceSet As String = "SET,"

        'Incluimos level
        cruceSet = cruceSet & "$" & tlevel.Text

        'Incluimos function
        cruceSet = cruceSet & "$§" & tFunction.Text

        'Incluimos Aligment
        cruceSet = cruceSet & "&" & tAligment.Text

        'Incluimos Conduccion
        cruceSet = cruceSet & "&"

        Dim hayuno As Boolean = False
        For i = 0 To cConduccion.Items.Count - 1
            If (cConduccion.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cConduccion.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cConduccion.Items.Count - 1
                    cConduccion.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet

        'Incluimos NoAsignado
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cNoAsignado.Items.Count - 1
            If (cNoAsignado.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cNoAsignado.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cNoAsignado.Items.Count - 1
                    cNoAsignado.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet

        'Incluimos Montaje Bruto
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cMontajeBruto.Items.Count - 1
            If (cMontajeBruto.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cMontajeBruto.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cMontajeBruto.Items.Count - 1
                    cMontajeBruto.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet

        'Incluimos Transmision
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cTransmision.Items.Count - 1
            If (cTransmision.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cTransmision.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cTransmision.Items.Count - 1
                    cTransmision.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet

        'Incluimos Nivel de Carga
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cNivelCarga.Items.Count - 1
            If (cNivelCarga.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cNivelCarga.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cNivelCarga.Items.Count - 1
                    cNivelCarga.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet

        'Incluimos Variante de Pais
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cVariantesPais.Items.Count - 1
            If (cVariantesPais.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cVariantesPais.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cVariantesPais.Items.Count - 1
                    cVariantesPais.SetItemChecked(i, False)
                Next
            End If
        Next
        Tset.Text = cruceSet


        'Incluimos Tipo
        cruceSet = cruceSet & "&"

        hayuno = False
        For i = 0 To cTipo.Items.Count - 1
            If (cTipo.GetItemChecked(i)) Then
                If (hayuno = True) Then
                    cruceSet = cruceSet & "!"
                End If
                cruceSet = cruceSet & cTipo.Items(i).ToString
                hayuno = True
            Else
                Dim x As Integer = 0
                For x = 0 To cTipo.Items.Count - 1
                    cTipo.SetItemChecked(i, False)
                Next
            End If
        Next

        'Incluimos Final clave Sol

        cruceSet = cruceSet & "§"

        Tset.Text = cruceSet
        'Copiar al Portapapeles
        My.Computer.Clipboard.SetText(Tset.Text)


    End Sub

    Private Sub visualizarCodigos()
        If nivelEnsamblado(1) > "" Then
            tlevel.Text = nivelEnsamblado(1)
        End If
        If valorAtributo(0) > "" Then
            tFunction.Text = valorAtributo(0)
        End If
        If valorAtributo(1) > "" Then
            tAligment.Text = valorAtributo(1)
        End If
        Try
            Dim Separo() As String
            Dim ficheroTxt() As String = File.ReadAllLines(Entrada)
            For Each linea In ficheroTxt
                If linea.Substring(0, 2) <> "1-" Then
                    Separo = Split(linea, vbTab)

                    If Separo(0) <> "" And valorAtributo(8).IndexOf(Separo(0)) <> -1 Then
                        cTipo.Items.Add(Separo(0), True)
                    ElseIf Separo(0) <> "" Then
                        cTipo.Items.Add(Separo(0))
                    End If

                    If Separo(1) <> "" And valorAtributo(2).IndexOf(Separo(1)) <> -1 Then
                        cConduccion.Items.Add(Separo(1), True)
                    ElseIf Separo(1) <> "" Then
                        cConduccion.Items.Add(Separo(1))
                    End If
                    If Separo(2) <> "" And valorAtributo(4).IndexOf(Separo(2)) <> -1 Then
                        cMontajeBruto.Items.Add(Separo(2), True)
                    ElseIf Separo(2) <> "" Then
                        cMontajeBruto.Items.Add(Separo(2))
                    End If
                    If Separo(3) <> "" And valorAtributo(5).IndexOf(Separo(3)) <> -1 Then
                        cTransmision.Items.Add(Separo(3), True)
                    ElseIf Separo(3) <> "" Then
                        cTransmision.Items.Add(Separo(3))
                    End If
                    If Separo(4) <> "" And valorAtributo(6).IndexOf(Separo(4)) <> -1 Then
                        cNivelCarga.Items.Add(Separo(4), True)
                    ElseIf Separo(4) <> "" Then
                        cNivelCarga.Items.Add(Separo(4))
                    End If
                    If Separo(5) <> "" And valorAtributo(7).IndexOf(Separo(5)) <> -1 Then
                        cVariantesPais.Items.Add(Separo(5), True)
                    ElseIf Separo(5) <> "" Then
                        cVariantesPais.Items.Add(Separo(5))
                    End If
                    If Separo(6) <> "" And valorAtributo(3).IndexOf(Separo(6)) <> -1 Then
                        cNoAsignado.Items.Add(Separo(6), True)
                    ElseIf Separo(6) <> "" Then
                        cNoAsignado.Items.Add(Separo(6))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Probablemente tengas el fichero abierto, Cierralo.")
        End Try

    End Sub

    Private Sub Beditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Beditar.Click
        Dim Arch = Shell("notepad " & Entrada, vbNormalFocus)
    End Sub

    Private Sub BRefrescar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BRefrescar.Click
        borrarListas()
        visualizarCodigos()
    End Sub

    Private Sub borrarListas()
        tlevel.Text = ""
        tFunction.Text = ""
        tAligment.Text = ""
        cTipo.Items.Clear()
        cConduccion.Items.Clear()
        cMontajeBruto.Items.Clear()
        cTransmision.Items.Clear()
        cNivelCarga.Items.Clear()
        cVariantesPais.Items.Clear()
        cNoAsignado.Items.Clear()
    End Sub

    'Descomponer la SET inicial
    Private Function DescomponerSetInicial()
        fila = TsetInicial.Text

        ' Descomponer SET en 3 trozos
        trozo = Split(fila, "§")
        'MsgBox("Trozo-0" & vbCrLf & trozo(0))
        'MsgBox("Trozo-1" & vbCrLf & trozo(1))
        'MsgBox("Trozo-2" & vbCrLf & trozo(2))

        If (UBound(trozo) = 2) Then
            'Separa los 2 atributos por $ de trozo(0)
            nivelEnsamblado = Split(trozo(0), "$")
            If (UBound(nivelEnsamblado) = 2) Then
                'MsgBox("Nivel Ensamblado-0" & vbCrLf & nivelEnsamblado(0))
                'MsgBox("Nivel Ensamblado-1" & vbCrLf & nivelEnsamblado(1))

                'Separa los 9 atributos por & de trozo(1)
                valorAtributo = Split(trozo(1), "&")
                If (UBound(valorAtributo) = 8) Then
                    Return True
                    'MsgBox(trozo(1) & vbCrLf &
                    '       "0." & valorAtributo(0) & vbCrLf &
                    '       "1." & valorAtributo(1) & vbCrLf &
                    '       "2." & valorAtributo(2) & vbCrLf &
                    '       "3." & valorAtributo(3) & vbCrLf &
                    '       "4." & valorAtributo(4) & vbCrLf &
                    '       "5." & valorAtributo(5) & vbCrLf &
                    '       "6." & valorAtributo(6) & vbCrLf &
                    '       "7." & valorAtributo(7) & vbCrLf &
                    '       "8." & valorAtributo(8) & vbCrLf)

                Else
                    'No tiene los 8 & atributos 
                    MsgBox("No tiene los 8 & atributos la set Inicial")
                    Return False
                End If
            Else
                'No tiene 2 niveles de $ ensamblado 
                MsgBox("No tiene 2 niveles de $ ensamblado la set Inicial")
                Return False
            End If
        Else
            'No tiene 2 trozos § la set Inicial
            MsgBox("No tiene 3 trozos § la set Inicial")
            Return False
        End If
    End Function

    Private Sub TsetInicial_TextChanged(sender As System.Object, e As System.EventArgs) Handles TsetInicial.TextChanged
        If DescomponerSetInicial() = True Then
            borrarListas()
            visualizarCodigos()
            componerSet()
        Else
            borrarListas()
            Tset.Text = ""
        End If

    End Sub

    Private Sub inicializaArays()
        trozo(0) = ""
        trozo(1) = ""
        trozo(2) = ""
        nivelEnsamblado(0) = ""
        nivelEnsamblado(1) = ""
        nivelEnsamblado(2) = ""
        valorAtributo(0) = ""
        valorAtributo(1) = ""
        valorAtributo(2) = ""
        valorAtributo(3) = ""
        valorAtributo(4) = ""
        valorAtributo(5) = ""
        valorAtributo(6) = ""
        valorAtributo(7) = ""
        valorAtributo(8) = ""
    End Sub
End Class
