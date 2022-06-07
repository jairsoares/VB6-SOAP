Attribute VB_Name = "SOC"


Public Function ws_SOC() As Boolean
Dim parametros As String
Dim aspas As String
Dim header As String

aspas = """"
parametros = "{" & _
                aspas & "empresa" & aspas & ":" & aspas & "579077" & aspas & "," & _
                aspas & "codigo" & aspas & ":" & aspas & "27753" & aspas & "," & _
                aspas & "chave" & aspas & ":" & aspas & "5b1710ab99ec6732ab02" & aspas & "," & _
                aspas & "tipoSaida" & aspas & ":" & aspas & "csv" & aspas & "," & _
                aspas & "EmpresaSel" & aspas & ":" & aspas & "592390" & aspas & "," & _
                aspas & "funcionarioInicio" & aspas & ":" & aspas & "0" & aspas & "," & _
                aspas & "funcionarioFim" & aspas & ":" & aspas & "9999999999" & aspas & "," & _
                aspas & "dataInicio" & aspas & ":" & aspas & "01/06/2020" & aspas & "," & _
                aspas & "dataFim" & aspas & ":" & aspas & "13/07/2020" & aspas & "," & _
                aspas & "tpExame" & aspas & ":" & aspas & "1,2,3,4,5,6" & aspas & _
             "}"

header = "<Envelope xmlns=" & aspas & "http://schemas.xmlsoap.org/soap/envelope/" & aspas & ">" & vbNewLine & _
         "<body>" & vbNewLine & _
         "<exportaDadosWs xmlns=" & aspas & "http://services.soc.age.com/" & aspas & ">" & vbNewLine & _
         "   <!-- Optional --> " & vbNewLine & _
         "   <arg0 xmlns=" & aspas & aspas & ">" & vbNewLine & _
         "       <parametros>" & parametros & "</parametros>" & vbNewLine & _
         "   </arg0>" & vbNewLine & _
         "</exportaDadosWs>" & vbNewLine & _
         "</body>" & vbNewLine & _
         "</Envelope>"




Dim Connector As New MSSOAPLib30.HttpConnector30
Dim Serializer As New MSSOAPLib30.SoapSerializer30
Dim Reader As New MSSOAPLib30.SoapReader30

Connector.Property("EndPointURL") = "https://ws1.soc.com.br/WSSoc/services/ExportaDadosWs?wsdl"
Connector.Property("UseSSL") = True
Connector.Connect
Connector.BeginMessage
    Serializer.Init Connector.InputStream
    Serializer.StartEnvelope
    Serializer.StartBody
    Serializer.StartElement "exportaDadosWs", "http://services.soc.age.com/"
    Serializer.StartElement "arg0"
    Serializer.StartElement "parametros"
    Serializer.WriteString parametros
    Serializer.EndElement
    Serializer.EndElement
    Serializer.EndElement
    Serializer.EndBody
    Serializer.EndEnvelope
Connector.EndMessage
Reader.Load Connector.OutputStream
Dim ret As String
ret = Reader.Body.Text
If Dir(App.Path & "\retorno.csv") <> "" Then
   Kill App.Path & "\retorno.csv"
End If

Dim sBuffer As String
Dim wBuffer As String
Dim wArrayBytes() As String
Dim iBufferArrayBytes As Long

sBuffer = Mid(ret, InStr(1, ret, "}") + 1)
wBuffer = Mid(sBuffer, 1, Len(sBuffer) - 4)
wArrayBytes = Split(sBuffer, vbLf)
sBuffer = ""
For iBufferArrayBytes = 1 To UBound(wArrayBytes) - 1
    sBuffer = sBuffer & wArrayBytes(iBufferArrayBytes) & vbCrLf
Next

Close #1
Open App.Path & "\retorno.csv" For Output As #1
Print #1, sBuffer
Close #1


If Dir(App.Path & "\retorno.csv") <> "" Then
   Form1.txtStatus.Text = "Arquivo " & App.Path & "\retorno.csv criado com sucesso."
   ws_SOC = True
Else
   Form1.txtStatus.Text = "Falha na obtenção dos dados."
   ws_SOC = False
End If




End Function


