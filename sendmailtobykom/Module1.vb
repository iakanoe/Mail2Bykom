Imports System.Data.OleDb
Imports System.Threading.Thread
Imports System.IO.Path
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Sockets
Imports MySql.Data.MySqlClient
Module Module1
    Dim mailmail As New Class1
    Dim markfordelete As New POP3
    Dim serverip As String
    Dim conteo As Integer = 0
    Sub Main()
        Console.Title = "BUSCA SMS"
        escribirenlog("Buscando SMS")
        Console.WriteLine("Busca SMS")
        Console.WriteLine("Buscando SMS..." & vbCrLf & vbCrLf)
        Do
            Console.WriteLine("Ultimo chequeo: " & Now.ToString("dd/MM/yyyy HH:mm:ss"))
            Dim mailsrecibidos As Object(,)
            Dim mail1(2) As String
            mailsrecibidos = mailmail.email_recieve("mail.monitoreomayorista.com", "sms@monitoreomayorista.com", "Riva7906", 110, True)
            If mailsrecibidos IsNot Nothing Then
                For i = 0 To UBound(mailsrecibidos, 2)
                    mail1(0) = mailsrecibidos(0, i)
                    mail1(1) = mailsrecibidos(1, i)
                    mail1(2) = mailsrecibidos(2, i)
                    workmail(mail1)
                Next
            End If
            If conteo = 4 Then
                mandar("000000", 8001, "192.168.5.100")
                escribirenlog("Heartbit enviado")
                conteo = 0
            Else
                conteo = conteo + 1
            End If
            Sleep(30000)
        Loop
    End Sub
    Sub workmail(mail As String())
        Dim sender As String = mail(0)
        Dim tema As String = mail(1)
        Dim cuerpo As String = mail(2)
        Dim d1, d2, d3, d4, cuentahex, paqueteudp, abonado, rc As String
        Dim puerto As Integer
        Dim numero As String
        Dim porXhoras As String
        Dim evt As String
        Dim numhab As Boolean
        Dim channel As Integer
        If tema.Contains("origen") Then
            'numero = Replace(Replace(Mid$(tema, InStr(tema, "origen: ") + 8, tema.Length - 11 - InStr(tema, "origen: ") + 8), vbLf, Nothing), " ", Nothing)
            numero = Mid(tema, InStr(tema, "origen: ") + 8, (InStr(tema, "Canal:") - 1) - (InStr(tema, "origen: ") + 8))
            channel = CInt(Mid(tema, InStr(tema, "Canal: ") + 7, 1))
        Else
            numero = "desconocido"
            channel = 0
        End if
        cuerpo = Mid$(cuerpo, 3, InStr(cuerpo, "/") + 6)
        Dim mogf As String = UCase$(Left$(cuerpo, 6))
        If channel <> 8 Then
            If mogf = "PRUEBA" Or mogf = "NORMAL" Or mogf = "LLAMAR" Then
                rc = UCase$(Mid$(cuerpo, 8, 2))
                abonado = Mid$(cuerpo, 11, 4)
                Dim puertoip As Object() = buscarendb("rc", "=", rc, "rcpuerto")
                Dim numerorc As Object() = buscarendb("telefonotecnico", "=", numero, "phone", "rc1, rc2, rc3, rc4, rc5, rc6, rc7")
                numhab = IsInArray(numerorc, rc)
                Dim title As String = "Control de técnicos en pruebas para RC"
                Dim mailadr As String = buscarendb("rc", "=", rc, "rcpuerto", "email")(0)
                Dim aaa As New Class1
                Dim nombre As String
                Try
                    nombre = buscarendb("telefonotecnico", "=", numero, "phone", "nombre")(0)
                Catch ex As Exception
                    nombre = "desconocido"
                End Try
                If mogf = "NORMAL" Then
                    evt = "080902"
                    porXhoras = Nothing
                ElseIf mogf = "LLAMAR"
                    evt = "080903"
                    porXhoras = Nothing
                ElseIf Mid$(cuerpo, 15, 2) = ":2"
                    evt = "080900"
                    porXhoras = " por 2 horas"
                ElseIf Mid$(cuerpo, 15, 2) = ":4"
                    evt = "080901"
                    porXhoras = " por 4 horas"
                ElseIf Mid(cuerpo, 15, 2) = ":1"
                    porXhoras = " por 1 hora"
                    evt = "080809"
                Else
                    evt = "080809"
                    porXhoras = " por 1 hora"
                End If
                If numhab = True Then
                    Dim enviado As String
                    If mogf = "PRUEBA" Then
                        enviado = sendsms("La cuenta " & rc & "/" & abonado & " será puesta en prueba.", numero, 7)
                        aaa.email_send("mail.monitoreomayorista.com", False, "sms@monitoreomayorista.com", "Riva7906", 25, "sms@monitoreomayorista.com", title, "--- NO RESPONDER ESTE EMAIL, CUENTA NO SUPERVISADA ---" & vbCrLf & vbCrLf & "Hemos recibido una órden de poner en prueba a una de sus cuentas." & vbCrLf & "Teléfono emisor: " & numero & vbCrLf & "Nombre: " & nombre & vbCrLf & "Cuenta: " & abonado & vbCrLf & "Mensaje: " & cuerpo & vbCrLf & "Estado: " & enviado & vbCrLf & vbCrLf & "RAM SRL" & vbCrLf & "0810-362-0362" & vbCrLf & vbCrLf & "--- MAIL ENVIADO AUTOMÁTICAMENTE ---" & vbCrLf, mailadr)
                    ElseIf mogf = "NORMAL" Then
                        aaa.email_send("mail.monitoreomayorista.com", False, "sms@monitoreomayorista.com", "Riva7906", 25, "sms@monitoreomayorista.com", title, "--- NO RESPONDER ESTE EMAIL, CUENTA NO SUPERVISADA ---" & vbCrLf & vbCrLf & "Hemos recibido una órden de sacar de prueba a una de sus cuentas." & vbCrLf & "Teléfono emisor: " & numero & vbCrLf & "Nombre: " & nombre & vbCrLf & "Cuenta: " & abonado & vbCrLf & "Mensaje: " & cuerpo & vbCrLf & "Estado: " & sendsms("La cuenta " & rc & "/" & abonado & "será sacada de prueba.", numero, 7) & vbCrLf & vbCrLf & "RAM SRL" & vbCrLf & "0810-362-0362" & vbCrLf & vbCrLf & "--- MAIL ENVIADO AUTOMÁTICAMENTE ---" & vbCrLf, mailadr)
                    End If
                    puerto = puertoip(1)
                    serverip = CStr(puertoip(2))
                    d1 = Left$(abonado, 1)
                    d2 = Mid$(abonado, 2, 1)
                    d3 = Mid$(abonado, 3, 1)
                    d4 = Right$(abonado, 1)
                    cuentahex = "0" & d1 & "0" & d2 & "0" & d3 & "0" & d4
                    paqueteudp = "40d769920490" & cuentahex & "010801" & evt & "0a010a0a0a0e3e040081"
                    Console.WriteLine("[" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "]: " & mogf & " " & rc & "/" & abonado & porXhoras & " --- número de origen " & numero & " - habilitado: " & numhab.ToString & ".")
                    escribirenlog(mogf & " " & rc & "/" & abonado & porXhoras & " --- número de origen " & numero & " - habilitado: " & numhab.ToString & ".")
                    mandar(paqueteudp, puerto, serverip)
                Else
                    Console.WriteLine("[" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "]: " & mogf & " " & rc & "/" & abonado & porXhoras & " --- número de origen " & numero & " - habilitado: " & numhab.ToString & ".")
                    escribirenlog(mogf & " " & rc & "/" & abonado & porXhoras & " --- número de origen " & numero & " - habilitado: " & numhab.ToString & ".")
                    title = "ALERTA URGENTE NO AUTORIZADO - Control de técnicos en pruebas para RC"
                    aaa.email_send("mail.monitoreomayorista.com", False, "sms@monitoreomayorista.com", "Riva7906", 25, "sms@monitoreomayorista.com", title, "--- NO RESPONDER ESTE EMAIL, CUENTA NO SUPERVISADA ---" & vbCrLf & vbCrLf & "Hemos recibido una órden de poner en prueba a una de sus cuentas." & vbCrLf & "Teléfono emisor: " & numero & vbCrLf & "Nombre: " & nombre & vbCrLf & "Cuenta: " & abonado & vbCrLf & "Mensaje: " & cuerpo & vbCrLf & vbCrLf & "RAM SRL" & vbCrLf & "0810-362-0362" & vbCrLf & vbCrLf & "--- MAIL ENVIADO AUTOMÁTICAMENTE ---" & vbCrLf, mailadr)
                End If
            ElseIf Left$(mogf, 5) = "GL300"
                Console.WriteLine(Now.ToString("[dd/MM/yyyy HH:mm:ss]: ") & "SOS RECIBIDO DE " & numero & ".")
                escribirenlog("SOS RECIBIDO DE " & numero & ".")
                Dim ref As String() = getphonename(numero)
                Console.WriteLine(Now.ToString("[dd/MM/yyyy HH:mm:ss]: ") & "SOS ENCONTRADO DE " & ref(2) & " --- ENVIANDO A BYKOM")
                cuentahex = pone_ceros(ref(2))
                escribirenlog("SOS ENCONTRADO DE " & ref(1) & "-" & cuentahex & " --- ENVIANDO A BYKOM")
                cuentahex = "0" & cuentahex.Chars(0) & "0" & cuentahex.Chars(1) & "0" & cuentahex.Chars(2) & "0" & cuentahex.Chars(3)
                evt = "080904"
                paqueteudp = "40d769920490" & cuentahex & "010801" & evt & "0a010a0a0a0e3e040081"
                mandar(paqueteudp, 8033, "ram.dyndns.ws")
            ElseIf numero IsNot "desconocido" And numero.Length > 5
                Console.WriteLine("[" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "]: SMS DESCONOCIDO RECIBIDO DE " & numero & ".")
                escribirenlog("SMS RECIBIDO DE " & numero & ".")
                Dim aaa As New Class1
                cuerpo = cuerpo.Replace(vbLf, Nothing)
                Dim ref As String() = getphonename(numero)
                sendsms("Su mensaje no fue leído. Comuníquese al 0810-362-0362. Gracias.", numero, 7)
                Console.WriteLine(Now.ToString("[dd/MM/yyyy HH:mm:ss]: ") & "SMS ENCONTRADO DE " & ref(0) & " " & ref(1) & "-" & ref(2) & "... ENVIANDO RESPUESTA")
                escribirenlog("SMS ENCONTRADO DE " & ref(0) & " " & ref(1) & "-" & ref(2) & "... ENVIANDO RESPUESTA")
                aaa.email_send("mail.monitoreomayorista.com", False, "sms@monitoreomayorista.com", "Riva7906", 25, "sms@monitoreomayorista.com", "SMS RECIBIDO", "--- NO RESPONDER A ESTE EMAIL ---" & vbCrLf & vbCrLf & cuerpo & vbCrLf & vbCrLf & "Número de origen: " & numero & vbCrLf & "Nombre: " & ref(0) & vbCrLf & "Abonado: " & ref(1) & "-" & pone_ceros(ref(2)) & vbCrLf & vbCrLf & "--- NO RESPONDER A ESTE EMAIL ---" & vbCrLf, "smsfiltrado@monitoreomayorista.com")
            Else
                Console.WriteLine(Now.ToString("[dd/MM/yyyy HH:mm:ss]: ") & "MAIL DESCARTADO: """ & tema & """ de " & sender & ".")
                escribirenlog("MAIL DESCARTADO: """ & tema & """ de " & sender & ".")
            End If
        End If
    End Sub
    Private Sub mandar(paquete As String, puerto As Integer, ip As String)
        Dim strPath As String = GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().CodeBase)
        Dim adr As IPAddress = Dns.GetHostAddresses(ip).First
        send(paquete, adr.ToString, puerto)
    End Sub
    Sub send(datos As String, ip As String, puerto As Integer)
        Dim uc As New UdpClient
        uc = New UdpClient
        uc.Connect(ip, puerto)
        Dim senddata As Byte()
        senddata = hex2byte(datos)
        uc.Send(senddata, senddata.Length)
    End Sub
    Function getphonename(phone As String) As String()
        Dim conexionsql As New MySqlConnection
        Dim adapter As New MySqlDataAdapter
        Dim reg As New DataSet
        Dim consulta As String
        Dim idd, idd2, idrc As Integer
        Dim response(2) As String
        Try
            conexionsql = New MySqlConnection
            conexionsql.ConnectionString = "server=192.168.5.100;user id=bykom;password=bykom;"
            conexionsql.Open()

            consulta = "SELECT ORDER_RL FROM bykom.tlrlpersonas WHERE LOWER(CONVERT(TELEFONO USING utf8mb4)) LIKE '%" & Right(phone, 7) & "%'"
            adapter = New MySqlDataAdapter(consulta, conexionsql)
            reg = New DataSet
            adapter.Fill(reg, "bykom.tlrlpersonas")
            If reg.Tables(0).Rows.Count > 1 Then
                consulta = "SELECT ORDER_RL FROM bykom.tlrlpersonas WHERE LOWER(CONVERT(TELEFONO USING utf8mb4)) LIKE '%" & Right(phone, 8) & "%'"
                adapter = New MySqlDataAdapter(consulta, conexionsql)
                reg = New DataSet
                adapter.Fill(reg, "bykom.tlrlpersonas")
            End If
            idd = CInt(reg.Tables(0).Rows(0).Item(0))

            consulta = "SELECT NOMBRE_MIX FROM bykom.tlmapersonas WHERE ORDER_ID=" & CStr(idd)
            adapter = New MySqlDataAdapter(consulta, conexionsql)
            reg = New DataSet
            adapter.Fill(reg, "bykom.tlmapersonas")
            response(0) = CStr(reg.Tables(0).Rows(0).Item(0))

            consulta = "SELECT ORDER_RL FROM bykom.abrlusuarios WHERE CODIGO_ID=" & CStr(idd)
            adapter = New MySqlDataAdapter(consulta, conexionsql)
            reg = New DataSet
            adapter.Fill(reg, "bykom.abrlusuarios")
            If reg.Tables(0).Rows.Count = 0 Then
                consulta = "SELECT ORDER_RL FROM bykom.abrltelefonos WHERE CODIGO_ID=" & CStr(idd)
                adapter = New MySqlDataAdapter(consulta, conexionsql)
                reg = New DataSet
                adapter.Fill(reg, "bykom.abrlusuarios")
            End If
            idd2 = CInt(reg.Tables(0).Rows(0).Item(0))

            consulta = "SELECT ID_RC, ID_CL FROM bykom.abmacodigos WHERE ORDER_ID=" & CStr(idd2)
            adapter = New MySqlDataAdapter(consulta, conexionsql)
            reg = New DataSet
            adapter.Fill(reg, "bykom.abmacodigos")
            response(2) = CStr(reg.Tables(0).Rows(0).Item(1))
            idrc = CInt(reg.Tables(0).Rows(0).Item(0))

            consulta = "SELECT CODIGOALFA FROM bykom.rcmacodigos WHERE ORDER_ID=" & CStr(idrc)
            adapter = New MySqlDataAdapter(consulta, conexionsql)
            reg = New DataSet
            adapter.Fill(reg, "bykom.rcmacodigos")
            response(1) = CStr(reg.Tables(0).Rows(0).Item(0))

            conexionsql.Close()
        Catch ex As Exception

            response(0) = "DESCONOCIDO"
        response(1) = "DESCONOCIDO"
        response(2) = "DESCONOCIDO"
        End Try
        Return response
    End Function
    Function pone_ceros(texto As String) As String
        For i As Integer = 1 To (4 - texto.Length)
            texto = "0" & texto
        Next
        Return texto
    End Function
    Function hex2byte(string_ As String) As Byte()
        string_ = string_.Replace(" "c, "")
        Dim nBytes = string_.Length \ 2
        Dim a(nBytes - 1) As Byte
        For i As Integer = 0 To nBytes - 1
            Try
                a(i) = Convert.ToByte(string_.Substring(i * 2, 2), 16)
            Catch ex As Exception
            End Try
        Next
        Return a
    End Function
    Function buscarendb(buscarpor As String, operador As String, where As String, tabla As String, Optional columnas As String = "*") As Object()
        Dim conexion As New OleDbConnection
        Dim adaptador As New OleDbDataAdapter
        Dim registro As New DataSet
        Try
            conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=rc.accdb"
            conexion.Open()
            Dim consulta As String = "SELECT " & columnas & " FROM " & tabla & " WHERE (" & buscarpor & " " & operador & " '" & where & "')"
            adaptador = New OleDbDataAdapter(consulta, conexion)
            registro = New DataSet
            adaptador.Fill(registro, "rcpuerto")
            conexion.Close()
            Dim returneo As Object() = registro.Tables("rcpuerto").Rows(0).ItemArray
            Return returneo
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Sub escribirenlog(linesinCrLf As String)
        File.AppendAllText("log" & Now.ToString("yyyyMMdd") & ".txt", "[" & Now.ToString("HH:mm:ss") & "]:" & linesinCrLf & vbCrLf)
    End Sub
    Function IsInArray(array_ As Object(), objeto As Object) As Boolean
        If array_ IsNot Nothing Then
            For Each objecto As Object In array_
                If objecto IsNot Nothing Then
                    If objecto.ToString = objeto Then Return True
                End If
            Next
            Return False
        Else
            Return False
        End If
    End Function
    Function sendsms(mensaje As String, numsend As String, channel As Integer) As String
        numsend = numsend.Replace(" ", "")
        numsend = numsend.Replace("+549", "")
        numsend = numsend.Replace("+54", "")

        Dim clientito As WebRequest = WebRequest.Create("http://192.168.5.228/cgi-bin/exec?cmd=api_queue_sms&username=lyric_api&password=adriancito&content=" & reemplazar(mensaje) & "&destination=" & numsend & "&api_version=0.08&channel=" & channel)
        clientito.UseDefaultCredentials = False
        Dim cred As New NetworkCredential("admin", "admin")
        clientito.Credentials = cred
        Dim respuesta As WebResponse = clientito.GetResponse
        Dim ds As Stream = respuesta.GetResponseStream
        Dim lector As New StreamReader(ds)
        Dim total As String = lector.ReadToEnd
        total = Mid$(total, 14, 5)
        If total = "true," Then
            Return "El SMS de confirmación fue enviado al técnico."
        Else
            Return "No se ha podido enviar el SMS de confirmación al técnico."
        End If
    End Function
    Function reemplazar(mensaje As String) As String
        Return mensaje.Replace(" ", "+")
    End Function
End Module
Public Class Class1
    Public Function email_send(servidorsmtp As String, ssl As Boolean, usuariosmtp As String, passsmtp As String, puertosmtp As String, mailfrom As String, subject As String, body As String, Optional mailtosimple As String = Nothing, Optional mailtomultiple As String() = Nothing) As Boolean
        Try
            Dim smtpserver As New SmtpClient
            Dim mail As New MailMessage()
            smtpserver.UseDefaultCredentials = False
            smtpserver.Credentials = New NetworkCredential(usuariosmtp, passsmtp)
            smtpserver.Port = puertosmtp
            smtpserver.EnableSsl = ssl
            smtpserver.Host = servidorsmtp
            mail = New MailMessage()
            mail.From = New MailAddress(mailfrom)
            If mailtosimple IsNot Nothing Then
                mail.To.Add(mailtosimple)
            Else
                For Each mailadress As String In mailtomultiple
                    mail.To.Add(mailadress)
                Next
            End If
            mail.Subject = subject
            mail.IsBodyHtml = False
            mail.Body = body
            smtpserver.Send(mail)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function email_recieve(servidorpop3 As String, userpop3 As String, passpop3 As String, puertopop3 As Integer, borrar As Boolean) As Object
        Dim popconn As POP3
        Dim mailmess As EmailMessage
        Dim intMessCnt As Integer
        Dim mailsrecibidos(2, -1) As Object
        popconn = New POP3
        mailmess = New EmailMessage
        If popconn.POPConnect(servidorpop3, userpop3, passpop3, puertopop3) Then
            intMessCnt = popconn.GetMailStat()
            If intMessCnt > 0 Then
                For i As Integer = 1 To intMessCnt
                    ReDim Preserve mailsrecibidos(2, UBound(mailsrecibidos, 2) + 1)
                    Dim strMailContent As String = popconn.GetMailMessage(i)
                    mailsrecibidos(0, UBound(mailsrecibidos, 2)) = mailmess.ParseEmail(strMailContent, "From:")
                    mailsrecibidos(1, UBound(mailsrecibidos, 2)) = mailmess.ParseEmail(strMailContent, "Subject:")
                    mailsrecibidos(2, UBound(mailsrecibidos, 2)) = mailmess.ParseBody()
                    If borrar Then
                        Try
                            popconn.SendData("DELE " & i)
                            If popconn.WaitFor("+OK") = False Then
                                End
                            End If
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                Next
                popconn.CloseConn()
                Return mailsrecibidos
            Else
                popconn.CloseConn()
                Return Nothing
            End If
        Else
            popconn.CloseConn()
            Return Nothing
        End If
    End Function
End Class
Public Class POP3

    'all the vars for use in the classs
    Dim TCP As TcpClient
    Dim POP3Stream As Stream
    Dim inStream As StreamReader
    Dim strDataIn, strNumMains(2) As String
    Dim intNoEmails As Integer

    'Class to connect to the passed mail server on port 110
    Function POPConnect(ByVal strServer As String, ByVal strUserName As String, ByVal strPassword As String, port As Integer) As Boolean

        'connect to the pop3 server over port 110
        Try
            TCP = New TcpClient
            TCP.Connect(strServer, port)

            'create stream into the ip
            POP3Stream = TCP.GetStream
            inStream = New System.IO.StreamReader(POP3Stream)

            'Make sure we get the ok back from the server
            If WaitFor("+OK") = False Then Return False

            'send the email down 
            SendData("USER " & strUserName)
            If WaitFor("+OK") = False Then Return False

            SendData("PASS " & strPassword)
            If WaitFor("+OK") = False Then Return False
            Return True
        Catch ex As Exception

            Return False
        End Try
    End Function

    'Function to get the number of mail messages waiting on the server
    Function GetMailStat() As Integer

        'send the stat command and make sure it returns as expected
        ' Try
        SendData("STAT")
        If WaitFor("+OK") = False Then
            Return 0
        Else
            'split up the response. It should be +OK Num of emails size of emails
            strNumMains = Split(strDataIn, " ")
            Return strNumMains(1)
            intNoEmails = strNumMains(1)
        End If
        'Catch ex As Exception
        Return 0
        intNoEmails = 0
        ' End Try
    End Function
    'function to take in what we expect and compare to what we actually get back
    Public Function WaitFor(ByVal strTarget As String) As Boolean
        strDataIn = inStream.ReadLine
        If strDataIn.Contains(strTarget) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    'This function will get the email message of the pop3 server, based on the message number passed
    Public Function GetMailMessage(ByVal intNum As Integer) As String

        Dim strTemp As String
        Dim strEmailMess As String = ""
        Try
            'send the command to the server to return that email back. Command is RETR and the email no ie RETR 1
            SendData("RETR " & Str(intNum))
            'make sure we get a proper response back
            If WaitFor("+OK") = False Then

                Return "No Email was Retrived"
            End If

            'Should be ok at this point to read in the tcp stream. We read it in until the end of the email
            'whitch is signified by a line containing only a fullpoint(chr46)
            strTemp = inStream.ReadLine

            While (strTemp <> ".")
                strEmailMess = strEmailMess & strTemp & vbCrLf
                strTemp = inStream.ReadLine
            End While

            Return strEmailMess

        Catch ex As Exception
            'just return an error message if we fell over
            Return "No Email was Retrived"
        End Try

    End Function

    'function that will mark an email for deletion. Delete does not occur until the QUIT is issued to the server
    Function MarkForDelete(ByVal intMailItem As Integer) As Boolean
        Return True
    End Function

    'Function that will quit the connection to the server (deleting marked mail) and close open readers etc
    Sub CloseConn()

        Try
            SendData("QUIT")
            inStream.Close()
            POP3Stream.Close()
        Catch ex As Exception
            escribirenlog(ex.ToString)
        End Try

    End Sub

    'function to send data down the tcp pipe
    Public Sub SendData(ByVal strCommand As String)

        Dim outBuff As Byte()

        outBuff = ConvertStringToByteArray(strCommand & vbCrLf)
        POP3Stream.Write(outBuff, 0, strCommand.Length + 2)

    End Sub

    Public Shared Function ConvertStringToByteArray(ByVal stringToConvert As String) As Byte()
        Return (New System.Text.ASCIIEncoding).GetBytes(stringToConvert)
    End Function

End Class
Public Class EmailMessage

    Private m_MessageSource As String

    'function that will call the main proc with what to bring back for everything but the body text
    Public Function ParseEmail(ByVal strMessage As String, ByVal strType As String) As String

        m_MessageSource = strMessage

        'call the parse routine with the pass filed we want
        ParseEmail = ParseHeader(strType)

    End Function

    'Function to parse each of the header parts out of the email
    Private Function ParseHeader(ByVal strHeader As String) As String

        Dim intLenToStart As Integer
        Dim intLenToLineEnd As Integer
        Dim strTmp As String

        intLenToStart = (InStr(m_MessageSource, strHeader) - 1)
        intLenToLineEnd = InStr(Mid(m_MessageSource, intLenToStart), vbCrLf)
        strTmp = m_MessageSource.Substring(intLenToStart, intLenToLineEnd)

        ParseHeader = Replace(strTmp, vbCrLf, "")

    End Function

    'Funtion to parse out the email body 
    Public Function ParseBody() As String

        'To get the body, everything after the first null line of the message is it (rfc822)
        Dim strTmp As String

        'set the temp var to the message body by getting everything after the null line
        strTmp = m_MessageSource.Substring(m_MessageSource.IndexOf(vbCrLf + vbCrLf))

        'get the encoding of the message out, that way we know if we have to run it through the base64 decode
        'routine or not
        If InStr(m_MessageSource, "Content-Transfer-Encoding: base64") Then
            'call the decode routine
            strTmp = DecodeBase64(strTmp)
        End If

        'if the jobs got html content, remove that from the body
        If InStr(strTmp, "------_=_NextPart_") Then
            strTmp = strTmp.Substring(1, strTmp.IndexOf(vbCrLf & vbCrLf & vbCrLf & "------_=_NextPart_"))
        End If


        'Strip out the odd hex that apears at the start and the end
        strTmp = Replace(strTmp, Chr(10) & Chr(9), "")
        strTmp = Replace(strTmp, Chr(13), "")

        ParseBody = Trim(strTmp)

    End Function

    'Function that will decode base64 encoded email body
    Private Function DecodeBase64(ByVal strBody As String)

        Try
            Dim encoding As New System.Text.UTF8Encoding
            Dim Buffer As Byte() = Convert.FromBase64String(strBody)
            DecodeBase64 = encoding.GetString(Buffer)
            encoding = Nothing
            Buffer = Nothing
        Catch ex As Exception
            MsgBox("A problem occured while decoding a base64 email", MsgBoxStyle.Critical)
            DecodeBase64 = "ERROR"
            Exit Function
        End Try

    End Function


End Class
