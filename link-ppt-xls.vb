'Code to create a correspondence between PowerPoint and Excel, insert in a visual basic module in Ms Excel

Option Explicit
Sub create()
    Dim shtHoja1 As Worksheet

    Dim strEstudiante As String
    Dim strCurso As String
    Dim strFecha As String
    Dim strHoras As String
    Dim PrimeraFila As Long

    Dim objPPT As Object
    Dim objPres As Object
    Dim objSld As Object
    Dim objShp As Object

    'Set shtHoja1 = Worksheets("Nombres")
    Set objPPT = CreateObject("Powerpoint.Application")
    objPPT.Visible = True

    'Modify the file path: C:\Certificates\Certificates.pptx
    Set objPres = objPPT.presentations.Open("C:\Certificados\Certificados.pptx")
    objPres.SaveAs "C:\Certificados\CertificadosNombres.pptx"

    PrimeraFila = 2
 
  Do While Worksheets("Nombres").Cells(PrimeraFila, 1) <> ""
        strEstudiante = Worksheets("Nombres").Cells(PrimeraFila, 1)
        strCurso = Worksheets("Nombres").Cells(PrimeraFila, 2)
        strFecha = Worksheets("Nombres").Cells(PrimeraFila, 3)
        strHoras = Worksheets("Nombres").Cells(PrimeraFila, 4)
        Set objSld = objPres.slides(1).Duplicate
        
        For Each objShp In objSld.Shapes
            If objShp.HasTextFrame Then
                If objShp.TextFrame.HasText Then
                    objShp.TextFrame.TextRange.Replace "<Estudiante>", strEstudiante
                    objShp.TextFrame.TextRange.Replace "<Curso>", strCurso
                    objShp.TextFrame.TextRange.Replace "<Fecha>", strFecha
                    objShp.TextFrame.TextRange.Replace "<Horas>", strHoras
                End If
            End If
        Next
        PrimeraFila = PrimeraFila + 1
    Loop

    objPres.slides(1).Delete
    objPres.Save
End Sub