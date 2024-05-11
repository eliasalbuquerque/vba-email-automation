Sub EnviarEmail()
    ' Este código VBA automatiza o envio de um e-mail via Outlook com uma 
    '   planilha Excel e um PDF como anexos, ambos gerados a partir de uma 
    '   planilha específica no Excel.
    ' A planilha é desprotegida, manipulada, salva como PDF e Excel, e então 
    '   protegida novamente antes do e-mail ser enviado.
    ' Todos os locais de código que possuem " ---- comentario ---- " nesse 
    '   formato, necessitam configuraçao adicional para o correto funcionamento 
    '   desse código.
    ' Os seguintes trechos de código devem ser configurados antes do uso:
    '   1. ---- Inserir ----
    '   2. ---- Caminho da pasta para salvar os anexos ---- 
    '   3. ---- Definir o corpo do email ---- 
    '   4. ---- Configuracao do email ---- 

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim strbody As String
    Dim mes As String
    Dim saudacao As String
    Dim mes_pdf As String
    Dim ano As String
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wsCopy As Worksheet
    Dim rng As Range
    Dim CaminhoArquivo As String
    Dim CaminhoPDF As String
    Dim Path As String
    Dim nome_aba As String
    Dim senha As String
    Dim nome_pdf As String



    ' ---- Inserir ---- 
    nome_aba = <nome_aba_planilha>
    senha = <senha_planilha_protegida>
    nome_pdf = <nome_pdf_a_ser_salvo>

    ' ---- Caminho da pasta para salvar os anexos ---- 
    Path = "C:\Seu\Caminho\Dos\Anexos\"



    ' Desproteger a planilha
    ThisWorkbook.Worksheets(nome_aba).Unprotect Password:=senha
    
    ' Iniciar o outlook
    Shell ("Outlook.exe")
    Application.Wait (Now + TimeValue("0:00:05"))

    ' Determinar a saudação com base na hora do dia
    If Hour(Now) < 12 Then
        saudacao = "Bom dia!"
    ElseIf Hour(Now) < 18 Then
        saudacao = "Boa tarde!"
    Else
        saudacao = "Boa noite!"
    End If

    ' Determinar o mês e ano a partir da célula B1
    mes = Range("B1").Value
    mes_pdf = Format(Range("B1").Value, "mmmm")
    ano = Format(Range("B1").Value, "yy")



    '  ---- Definir o corpo do email ---- 
    strbody = saudacao & "<br><br>" & _
    "Segue a minha agenda referente ao mês de " & mes & ".<br><br>" & _
    "Att."



    ' Criar um novo email no Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    ' Definir a planilha que você deseja salvar como PDF
    Set ws = ThisWorkbook.Sheets(nome_aba)
    
    ' Salvar a planilha "<nome_aba_planilha>" como PDF no formato "MÊSANO.pdf"
    CaminhoPDF = Path & <nome_pdf_a_ser_salvo> & mes_pdf & ano & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CaminhoPDF, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, OpenAfterPublish:=False, From:=1, To:=1
     
    ' Criar nova planilha e definir as celulas a serem copiadas da plan original
    Set wb = Application.Workbooks.Add
    Set wsCopy = wb.Sheets(1)
    Set rng = ws.Range("A1:D32")
 
    ' Copiar os valores e a formatação das células da planilha 
    rng.Copy
    With wsCopy.Range("A1")
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteColumnWidths
        Application.CutCopyMode = False
    End With
 
    ' Salvar a nova pasta de trabalho como um arquivo Excel
    CaminhoArquivo = Path & <nome_pdf_a_ser_salvo> & mes_pdf & ano & ".xlsx"
    wb.SaveAs Filename:=CaminhoArquivo
    wb.Close SaveChanges:=False
 


    ' ---- Configuracao do email ---- 
    With OutlookMail
        .To = "fulano1@gmail.com; fulano2@hotmail.com"
        .CC = "fulano3@gmail.com"
        .Subject = "Agenda " & mes
        .BodyFormat = 2 ' Formato HTML
        .HTMLBody = strbody
 
        ' Anexos
        .Attachments.Add CaminhoPDF
        .Attachments.Add CaminhoArquivo
 
        ' Exibir o email
        .Display
    End With


 
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    ' Proteger a planilha novamente no final
    ThisWorkbook.Worksheets(nome_aba).Protect Password:=senha, _
    UserInterfaceOnly:=True
    
End Sub

Sub mês_atual()
    ' Este código VBA atualiza a célula B1 com o mês e o ano atual em 
    '   maiúsculas, desprotegendo e protegendo novamente a planilha durante o 
    '   processo.

    ThisWorkbook.Worksheets(nome_aba).Unprotect Password:=senha
    Range("B1").Formula = "=UPPER(TEXT(TODAY(),""mmmm/AA""))"
    ThisWorkbook.Worksheets(nome_aba).Protect Password:=senha, _
    UserInterfaceOnly:=True
End Sub

Sub avanca_mes()
    ' Este código VBA avança o mês na célula B1 em um mês, mantendo o formato 
    '   em maiúsculas, desprotegendo e protegendo novamente a planilha durante 
    '   o processo.

    ThisWorkbook.Worksheets(nome_aba).Unprotect Password:=senha
    
    Dim dateValue As Date
    dateValue = CDate(Range("B1").Value)
    dateValue = DateAdd("m", 1, dateValue)
    Range("B1").Formula = "=UPPER(TEXT(""" & Format(dateValue, "dd/mm/yyyy") _
    & """,""mmmm/AA""))"
    
    ThisWorkbook.Worksheets(nome_aba).Protect Password:=senha, _
    UserInterfaceOnly:=True
End Sub

Sub retrocede_mes()
    ' Este código VBA retrocede o mês na célula B1 em um mês, mantendo o formato 
    '   em maiúsculas, desprotegendo e protegendo novamente a planilha durante 
    '   o processo.
    ThisWorkbook.Worksheets(nome_aba).Unprotect Password:=senha
    
    Dim dateValue As Date
    dateValue = CDate(Range("B1").Value)
    dateValue = DateAdd("m", -1, dateValue)
    Range("B1").Formula = "=UPPER(TEXT(""" & Format(dateValue, "dd/mm/yyyy") _
    & """,""mmmm/AA""))"
    
    ThisWorkbook.Worksheets(nome_aba).Protect Password:=senha, _
    UserInterfaceOnly:=True
End Sub

Private Sub Workbook_Open()
    ' ----- colocar em EstaPastaDeTrabalho -----
    ' Este código VBA é um evento que é acionado quando a pasta de trabalho é 
    '   aberta. Ele atualiza a célula B1 com o mês e o ano atual em maiúsculas, 
    '   a menos que a fórmula já contenha a função HOJE(), desprotegendo e 
    '   protegendo novamente a planilha durante o processo.
    ' Este código deve ser colocado no módulo "EstaPastaDeTrabalho" 
    '   (ThisWorkbook) do Editor VBA.

    ThisWorkbook.Worksheets(nome_aba).Unprotect Password:=senha
    
    If InStr(1, Range("B1").Formula, "HOJE()") > 0 Then
    Else
        Range("B1").Formula = "=UPPER(TEXT(TODAY(),""mmmm/AA""))"
    End If
    
    ThisWorkbook.Worksheets(nome_aba).Protect Password:=senha, UserInterfaceOnly:=True
End Sub
