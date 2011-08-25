Option Explicit On
Option Strict On

Public Class frmIIS

   Const secondsInDay As Integer = 86400

   Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       rdoHours.Checked = True

       ' This will return the hours, minutes, seconds
       CalculatePeriod()
   End Sub

   Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
       Dim ofdBrowse As New OpenFileDialog

       ofdBrowse.Title = "Browse for an IIS log file..."
       ofdBrowse.InitialDirectory = "C:\"
       ofdBrowse.Filter = "All files (*.*)|*.*|IIS log files (*.log)|*.log"
       ofdBrowse.FilterIndex = 2
       ofdBrowse.RestoreDirectory = True

       If ofdBrowse.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

           Try
               txtLogfile.Text = ofdBrowse.FileName
           Catch Ex As Exception
               MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
               'Finally

           End Try
       End If
   End Sub

   Private Sub LoadInFile()

       ' C:\Program Files\Log Parser 2.2\ex060707.log

       Dim fStream As IO.FileStream
       Dim sReader As IO.StreamReader
       Dim result As String = String.Empty

       fStream = New IO.FileStream(txtLogfile.Text, IO.FileMode.Open, _
               IO.FileAccess.Read, IO.FileShare.ReadWrite)

       sReader = New IO.StreamReader(fStream)

       ' read the entire file into a string if it has data in it
       If Not fStream.Length = 0 Then
           result = sReader.ReadToEnd()
       End If

       sReader.Close()
       fStream.Close()
       sReader = Nothing
       fStream = Nothing

       Dim arrLogLines As String() = Nothing

       ' Read every single line in the file into an index of the array
       arrLogLines = result.Split(Convert.ToChar("" & Microsoft.VisualBasic.Chr(10) & ""))

       '***IIS Log example
       ' log format
       ' #Software: Microsoft Internet Information Services 5.0
       ' #Version: 1.0
       ' #Date: 2007-04-18 11:56:28
       ' #Fields: date time c-ip cs-username s-ip s-port cs-method cs-uri-stem
       '          cs-uri-query sc-status sc-bytes cs-bytes time-taken cs(User-Agent) cs(Referer)
       '***

       Dim dt As New DataTable("log")
       ' Remove #Fields: (0)date, (1)time, (2)c-ip
       Dim revisedTitles As String = arrLogLines(3).Replace("#Fields: ", "")

       ' Place each IIS log title in its own index
       Dim arrTitles As String() = revisedTitles.Split(Convert.ToChar(" "))

       Dim i As Integer = 0
       Dim j As Integer = 0

       ' Read #Fields:
       ' Loop until the length which would be the last IIS log title
       While i < arrTitles.Length
           dt.Columns.Add(arrTitles(i))
           Debug.WriteLine(arrTitles(i))
           System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
       End While

       ' Read each line of the log
       j = arrLogLines.Length - 1

       ' j is 3 because it starts from the bottom of the file and goes up until it reaches
       ' the files comments section
       While j > 3

           If Not arrLogLines(j) = String.Empty Then
               ' Just in case an exceptional case occurs with an additional log being appended
               Try
                   dt.Rows.Add(arrLogLines(j).Split(Convert.ToChar(" ")))
                   System.Math.Max(System.Threading.Interlocked.Decrement(j), i + 1)
               Catch ex As Exception
                   MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "Error on parsing log line. ")
               End Try
           End If
       End While

       ' Some useful information for parsing
       ' 1 day = 86400 seconds
       ' 1 day = 1440 minutes
       ' 1 day = 24 hours

   End Sub

   Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click

       Try
           ' Read the IIS logfile specified in the textbox
           ReadLog(txtLogfile.Text)


       Catch ex As Exception

       End Try
   End Sub

   ' This method will return an integer for time based on the radio button selected by the user
   Private Function CalculatePeriod() As Integer

       ' the number of seconds in 1 day

       Dim timeValue As Integer = 0
       Dim summarySize As Integer = 0

       ' radio button control for time span
       If rdoHours.Checked Then
           timeValue = 3600
       ElseIf rdoMinutes.Checked Then
           timeValue = 60
       ElseIf rdoSeconds.Checked Then
           timeValue = 1
       Else
           ' shouldn't get here
       End If

       ' check that the timeValue is not 0, adn if not then calculate the summary(size)
       If Not timeValue = 0 Then
           summarySize = CInt(secondsInDay / timeValue)
           ' start from 1
           summarySize = summarySize - 1
       End If

       Return summarySize
   End Function

   Private Sub ReadLog(ByVal fileName As String)
       Dim streamReader As IO.StreamReader
       Dim logLine As String = Nothing
       Dim lineCount As Integer = 0
       Dim count As Integer = 0



       ' need to reset all indexes to 0

       Try
           streamReader = IO.File.OpenText(fileName)

           ' read each line in the log
           Do While True

               logLine = streamReader.ReadLine
               If logLine Is Nothing Then Exit Do
               logLine = logLine.Trim

               lineCount += 1

               ' check if there are empty lines, or comments and skip them
               If Not logLine = String.Empty AndAlso Not logLine.StartsWith("#") Then


               End If
           Loop

       Catch ex As Exception

       End Try
   End Sub

   Private Sub ParseLine()

       Dim httpResponse As Integer = 0
       Dim timeOffset As Integer = 0
       Dim timeSpan As Integer = 0
       Dim serverBytes As Integer = 0
       Dim clientBytes As Integer = 0
       Dim responseTimeMs As Integer = 0

   End Sub

   Private Sub IntializeData()

       ' represents an offset line in log
       Dim timeStamp As Integer = 0
       Dim httpResponse As Integer = 0
       Dim timespan As Integer = 0
       Dim serverBytes As Integer = 0
       Dim clientBytes As Integer = 0
       Dim responseTime As Integer = 0

       ' create an array for each column that will be output
       ' Setup the size of each array to the lowest resolution - seconds in a(day)
       Dim arraySize As Integer
       arraySize = secondsInDay - 1

       Dim concurrentUsers(arraySize) As Double
       Dim httpSuccess(arraySize) As Double
       Dim httpFailure(arraySize) As Double
       Dim c_bytes(arraySize) As Double
       Dim s_bytes(arraySize) As Double


       ' Setup offsets
       timeStamp = 1
       httpResponse = 9
       serverBytes = 10
       clientBytes = 11 ' apparently prod doesn't log client bytes, so this is server bytes
       timespan = 11
       responseTime = 12

       ' Pass in the value selected by the user -  seconds, minutes, or hours()
       Dim sumConcurrentUsers(CalculatePeriod()) As Double
       Dim sumSuccess(CalculatePeriod()) As Double
       Dim sumFailure(CalculatePeriod()) As Double
       Dim sumSbytes(CalculatePeriod()) As Double
       Dim sumCbytes(CalculatePeriod()) As Double
       Dim sumCdatarate(CalculatePeriod()) As Double
       Dim sumSdatarate(CalculatePeriod()) As Double
       Dim sumResponseTime(CalculatePeriod()) As Double
       Dim maxAverageDataRate As Double = 0.0
       Dim maxUsers As Double = 0.0
       Dim maxErrors As Double = 0.0

   End Sub

   Private Function ParseLineItem(ByVal item As String, ByVal itemNumber As Integer) As String

       Dim offset1 As Integer = 0
       Dim counter As Integer = 0
       Dim offset2 As Double = 0.0
       Dim result As String = String.Empty

       If itemNumber > 0 Then

           For counter = 1 To itemNumber
               offset1 = item.IndexOf(" ", offset1 + 1)
               If offset1 < 0 Then
                   Throw New Exception("GetLineItem invalid itemno")
               End If
           Next

       End If

       offset2 = item.IndexOf(" ", offset1 + 1)

       '''result = item.Substring(offset1, offset2 - offset1)

       Return result

   End Function

   Private Function ParseLineItemOffset(ByVal line As String, ByVal itemno As Integer) As Integer

       Try
           'Return TimeSpan.Parse(ParseLineItem(line, itemno)).TotalSeconds
       Catch ex As Exception
           MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "ParseLineItemOffset")
           Throw ex
       End Try

   End Function

   Private Function ParseLineItemDouble(ByVal line As String, ByVal itemno As Integer) As Integer

       Try
           'Return Double.Parse(ParseLineItem(line, itemno))
       Catch ex As Exception
           MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "ParseLineItemOffset")
           Throw ex
       End Try

   End Function

   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
       LoadInFile()
   End Sub
End Class