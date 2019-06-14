'**************************************************************************************************************************************************
'Code by Martin Viljoen 2019/06/04
'Post order from SQL database to SAGE 300
'**************************************************************************************************************************************************
Imports System
Imports System.Web
Imports System.Net
Imports System.Xml
Imports System.IO
Imports System.Text
Public Class frmMain
    'Web Server settings
    Dim WS_URL As String
    Dim Sage300User As String
    Dim Sage300Pass As String
    'General declares
    Dim strWSResponse As String
    Dim PostData As String

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        frmMySettings.Show()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WS_URL = My.Settings.WS_URL
        Sage300User = My.Settings.SAGE_USER
        Sage300Pass = My.Settings.SAGE_PASS
    End Sub

    Private Sub btn_PostData_Click(sender As Object, e As EventArgs) Handles btn_PostData.Click
        'USAGE: PostOrder(Sage_username,Sage_Password,JSON_Formatted_String)

        'Get Settings
        'Endpoint URL
        'Example: http://localhost/Sage300WebApi/v1.0/-/SAMINC/OE/OEOrders
        WS_URL = My.Settings.WS_URL 'Sage 300 End point URL
        'Sage300 user must have WEB-API access in security group setup
        Sage300User = My.Settings.SAGE_USER
        Sage300Pass = My.Settings.SAGE_PASS

        'Send Post request
        Label1.Text = "Status: Sending"
        PostOrder(Sage300User, Sage300Pass, txtPostData.Text)
        Label1.Text = "Status: ---"
    End Sub


    Function PostOrder(ByVal SageUserName As String, ByVal SagePassword As String, ByVal JSONString As String)
        On Error GoTo ErrHandler
        'Setup and Format HTML Basic Form Auth 
        Dim userCredentials As NetworkCredential = Nothing
        userCredentials = New NetworkCredential(SageUserName, SagePassword, "")
        'Determine the data length that we are about to post
        Dim DatainByteFormat As Byte() = Encoding.UTF8.GetBytes(JSONString)
        'Setup the web request and attach header meta information.
        Dim Sage300PostReq As Net.HttpWebRequest = Net.WebRequest.Create(WS_URL)
        'Setup Header Request
        Sage300PostReq.Credentials = userCredentials
        Sage300PostReq.ContentType = "application/json"
        Sage300PostReq.Method = "POST"
        Sage300PostReq.ContentLength = DatainByteFormat.Length
        'Build Data to be posted and attach to header
        Dim PostDatawriter As New StreamWriter(Sage300PostReq.GetRequestStream())
        PostDatawriter.Write(JSONString)
        PostDatawriter.Close()

        'Make the web request and get the response
        Dim response As Net.WebResponse = Sage300PostReq.GetResponse
        Dim stream As System.IO.Stream = response.GetResponseStream
        'Prepare buffer for reading from stream
        Dim buffer As Byte() = New Byte(1000) {}

        'Data read from stream is gathered here
        Dim data As New List(Of Byte)

        'Start reading stream
        Dim bytesRead = stream.Read(buffer, 0, buffer.Length)

        Do Until bytesRead = 0
            For i = 0 To bytesRead - 1
                data.Add(buffer(i))
            Next

            bytesRead = stream.Read(buffer, 0, buffer.Length)
        Loop


        'Get the JSON data
        'Debug.WriteLine(System.Text.Encoding.UTF8.GetString(data.ToArray))
        strWSResponse = System.Text.Encoding.UTF8.GetString(data.ToArray)
        'MsgBox(strWSResponse)


        response.Close()
        stream.Close()
        WriteToLog(strWSResponse)
        WriteJson(strWSResponse)
        Exit Function
ErrHandler:
        'MsgBox(Err.Description)
        WriteToLog(Err.Description)

    End Function

    Function WriteToLog(ByVal LogMessage As String)
        Dim today As DateTime
        today = Now
        Dim day = today.Day
        Dim month = today.Month
        Dim year = today.Year
        Dim hour = today.Hour
        Dim min = today.Minute
        Dim second = today.Second

        Dim MyYear As String
        Dim MyMonth As String
        Dim Myday As String
        Dim MyHour As String
        Dim MyMin As String
        Dim MySec As String

        MyYear = today.Year
        'Reformat date (Month and day only)
        If today.Month < 10 Then MyMonth = "0" & today.Month Else MyMonth = today.Month
        If today.Day < 10 Then Myday = "0" & today.Day Else Myday = today.Day
        'Reformat Time (HOUR,MIN,SECOND)
        If today.Hour < 10 Then MyHour = "0" & today.Hour Else MyHour = today.Hour
        If today.Minute < 10 Then MyMin = "0" & today.Minute Else MyMin = today.Minute
        If today.Second < 10 Then MySec = "0" & today.Second Else MySec = today.Second

        Dim file As System.IO.StreamWriter
        Dim fileName As String
        Dim DateTimePrefix As String

        DateTimePrefix = MyYear & "-" & MyMonth & "-" & Myday & " " & MyHour & ":" & MyMin & ":" & MySec

        fileName = "POST_ORDER_LOG " & MyYear & "-" & MyMonth & "-" & Myday & "_" & MyHour & "_" & MyMin & "_" & MySec & ".txt"
        file = My.Computer.FileSystem.OpenTextFileWriter("log\" & fileName, True)
        file.WriteLine(DateTimePrefix & " " & LogMessage)
        file.Close()
    End Function

    Function WriteJson(ByVal LogMessage As String)
        Dim today As DateTime
        today = Now
        Dim day = today.Day
        Dim month = today.Month
        Dim year = today.Year
        Dim hour = today.Hour
        Dim min = today.Minute
        Dim second = today.Second

        Dim MyYear As String
        Dim MyMonth As String
        Dim Myday As String
        Dim MyHour As String
        Dim MyMin As String
        Dim MySec As String

        MyYear = today.Year
        'Reformat date (Month and day only)
        If today.Month < 10 Then MyMonth = "0" & today.Month Else MyMonth = today.Month
        If today.Day < 10 Then Myday = "0" & today.Day Else Myday = today.Day
        'Reformat Time (HOUR,MIN,SECOND)
        If today.Hour < 10 Then MyHour = "0" & today.Hour Else MyHour = today.Hour
        If today.Minute < 10 Then MyMin = "0" & today.Minute Else MyMin = today.Minute
        If today.Second < 10 Then MySec = "0" & today.Second Else MySec = today.Second

        Dim file As System.IO.StreamWriter
        Dim fileName As String
        Dim DateTimePrefix As String

        DateTimePrefix = MyYear & "-" & MyMonth & "-" & Myday & " " & MyHour & ":" & MyMin & ":" & MySec

        fileName = "POST_ORDER_JSON " & MyYear & "-" & MyMonth & "-" & Myday & "_" & MyHour & "_" & MyMin & "_" & MySec & ".json"
        file = My.Computer.FileSystem.OpenTextFileWriter("log\" & fileName, True)
        file.WriteLine(LogMessage)
        file.Close()
    End Function


End Class
