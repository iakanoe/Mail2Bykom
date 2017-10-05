Imports System.IO
Imports System.Net.Sockets
Imports System.Net
Public Class Form1
    Dim TCP As New TcpClient
    Dim POP3Stream As Stream
    Dim inStream As StreamReader
    Dim strdatain As String
    Dim mailsrecibidos(2, -1) As Object
    Dim mailmess As EmailMessage
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            TCP.Connect("mail.monitoreomayorista.com", 110)
            POP3Stream = TCP.GetStream
            inStream = New StreamReader(POP3Stream)
            If waitforok() = False Then End
            Dim outBuff As Object
            outBuff = (New Text.ASCIIEncoding).GetBytes("USER sms@monitoreomayorista.com" & vbCrLf)
            POP3Stream.Write(outBuff, 0, ("USER sms@monitoreomayorista.com").Length + 2)
            If waitforok() = False Then End
            outBuff = New Object
            outBuff = (New Text.ASCIIEncoding).GetBytes("PASS Riva7906" & vbCrLf)
            POP3Stream.Write(outBuff, 0, ("PASS Riva7906").Length + 2)
            If waitforok() = False Then End
            outBuff = New Object
            outBuff = (New Text.ASCIIEncoding).GetBytes("STAT" & vbCrLf)
            POP3Stream.Write(outBuff, 0, ("STAT").Length + 2)
            If waitforok() = False Then End
            Dim stringss(1) As String
            stringss = Split(strdatain, " ")
            mailmess = New EmailMessage
            For i As Integer = 1 To stringss(1)
                ReDim Preserve mailsrecibidos(2, UBound(mailsrecibidos, 2) + 1)
                Dim strMailContent As String = getmailmessage(i)
                mailsrecibidos(0, UBound(mailsrecibidos, 2)) = mailmess.ParseEmail(strMailContent, "From:")
                mailsrecibidos(1, UBound(mailsrecibidos, 2)) = mailmess.ParseEmail(strMailContent, "Subject:")
                mailsrecibidos(2, UBound(mailsrecibidos, 2)) = mailmess.ParseBody()
            Next
            Dim conteomails As Integer = 0
            For i As Integer = 1 To UBound(mailsrecibidos, 2)
                If mailsrecibidos(2, i).ToString.Contains("iaka") Then
                    outBuff = New Object
                    outBuff = (New Text.ASCIIEncoding).GetBytes("DELE " & i & vbCrLf)
                    POP3Stream.Write(outBuff, 0, ("DELE " & i).Length + 2)
                    If waitforok() = False Then
                        MsgBox("No se pudo borrar el mensaje " & i)
                        End
                    End If
                    conteomails = conteomails + 1
                End If
            Next
            MsgBox("Finalizado. Borrados " & conteomails & " mails.")
            End
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function waitforok() As Boolean
        strdatain = inStream.ReadLine
        If strdatain.Contains("OK") Then
            Return True
        Else
            Return False
        End If
        Return False
    End Function
    Function getmailmessage(i As Integer) As String
        Dim strTemp As String
        Dim strEmailMess As String = ""
        Dim outbuff As Object
        outbuff = New Object
        outbuff = (New Text.ASCIIEncoding).GetBytes("RETR " & i & vbCrLf)
        POP3Stream.Write(outbuff, 0, ("RETR " & i).Length + 2)
        If waitforok() = False Then End
        strTemp = inStream.ReadLine
        While (strTemp <> ".")
            strEmailMess = strEmailMess & strTemp & vbCrLf
            strTemp = inStream.ReadLine
        End While
        Return strEmailMess
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