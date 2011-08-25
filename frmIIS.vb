{\rtf1\ansi\ansicpg1252\cocoartf1138
{\fonttbl\f0\froman\fcharset0 Times-Roman;}
{\colortbl;\red255\green255\blue255;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\deftab720
\pard\pardeftab720

\f0\fs24 \cf0 Option Explicit On\
Option Strict On\
\
Public Class frmIIS\
\
\'a0 \'a0Const secondsInDay As Integer = 86400\
\
\'a0 \'a0Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load\
\'a0 \'a0 \'a0 \'a0rdoHours.Checked = True\
\
\'a0 \'a0 \'a0 \'a0' This will return the hours, minutes, seconds\
\'a0 \'a0 \'a0 \'a0CalculatePeriod()\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click\
\'a0 \'a0 \'a0 \'a0Dim ofdBrowse As New OpenFileDialog\
\
\'a0 \'a0 \'a0 \'a0ofdBrowse.Title = "Browse for an IIS log file..."\
\'a0 \'a0 \'a0 \'a0ofdBrowse.InitialDirectory = "C:\\"\
\'a0 \'a0 \'a0 \'a0ofdBrowse.Filter = "All files (*.*)|*.*|IIS log files (*.log)|*.log"\
\'a0 \'a0 \'a0 \'a0ofdBrowse.FilterIndex = 2\
\'a0 \'a0 \'a0 \'a0ofdBrowse.RestoreDirectory = True\
\
\'a0 \'a0 \'a0 \'a0If ofdBrowse.ShowDialog() = System.Windows.Forms.DialogResult.OK Then\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0txtLogfile.Text = ofdBrowse.FileName\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Catch Ex As Exception\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0'Finally\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0End Try\
\'a0 \'a0 \'a0 \'a0End If\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Sub LoadInFile()\
\
\'a0 \'a0 \'a0 \'a0' C:\\Program Files\\Log Parser 2.2\\ex060707.log\
\
\'a0 \'a0 \'a0 \'a0Dim fStream As IO.FileStream\
\'a0 \'a0 \'a0 \'a0Dim sReader As IO.StreamReader\
\'a0 \'a0 \'a0 \'a0Dim result As String = String.Empty\
\
\'a0 \'a0 \'a0 \'a0fStream = New IO.FileStream(txtLogfile.Text, IO.FileMode.Open, _\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0IO.FileAccess.Read, IO.FileShare.ReadWrite)\
\
\'a0 \'a0 \'a0 \'a0sReader = New IO.StreamReader(fStream)\
\
\'a0 \'a0 \'a0 \'a0' read the entire file into a string if it has data in it\
\'a0 \'a0 \'a0 \'a0If Not fStream.Length = 0 Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0result = sReader.ReadToEnd()\
\'a0 \'a0 \'a0 \'a0End If\
\
\'a0 \'a0 \'a0 \'a0sReader.Close()\
\'a0 \'a0 \'a0 \'a0fStream.Close()\
\'a0 \'a0 \'a0 \'a0sReader = Nothing\
\'a0 \'a0 \'a0 \'a0fStream = Nothing\
\
\'a0 \'a0 \'a0 \'a0Dim arrLogLines As String() = Nothing\
\
\'a0 \'a0 \'a0 \'a0' Read every single line in the file into an index of the array\
\'a0 \'a0 \'a0 \'a0arrLogLines = result.Split(Convert.ToChar("" & Microsoft.VisualBasic.Chr(10) & ""))\
\
\'a0 \'a0 \'a0 \'a0'***IIS Log example\
\'a0 \'a0 \'a0 \'a0' log format\
\'a0 \'a0 \'a0 \'a0' #Software: Microsoft Internet Information Services 5.0\
\'a0 \'a0 \'a0 \'a0' #Version: 1.0\
\'a0 \'a0 \'a0 \'a0' #Date: 2007-04-18 11:56:28\
\'a0 \'a0 \'a0 \'a0' #Fields: date time c-ip cs-username s-ip s-port cs-method cs-uri-stem\
\'a0 \'a0 \'a0 \'a0' \'a0 \'a0 \'a0 \'a0 \'a0cs-uri-query sc-status sc-bytes cs-bytes time-taken cs(User-Agent) cs(Referer)\
\'a0 \'a0 \'a0 \'a0'***\
\
\'a0 \'a0 \'a0 \'a0Dim dt As New DataTable("log")\
\'a0 \'a0 \'a0 \'a0' Remove #Fields: (0)date, (1)time, (2)c-ip\
\'a0 \'a0 \'a0 \'a0Dim revisedTitles As String = arrLogLines(3).Replace("#Fields: ", "")\
\
\'a0 \'a0 \'a0 \'a0' Place each IIS log title in its own index\
\'a0 \'a0 \'a0 \'a0Dim arrTitles As String() = revisedTitles.Split(Convert.ToChar(" "))\
\
\'a0 \'a0 \'a0 \'a0Dim i As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim j As Integer = 0\
\
\'a0 \'a0 \'a0 \'a0' Read #Fields:\
\'a0 \'a0 \'a0 \'a0' Loop until the length which would be the last IIS log title\
\'a0 \'a0 \'a0 \'a0While i < arrTitles.Length\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0dt.Columns.Add(arrTitles(i))\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Debug.WriteLine(arrTitles(i))\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)\
\'a0 \'a0 \'a0 \'a0End While\
\
\'a0 \'a0 \'a0 \'a0' Read each line of the log\
\'a0 \'a0 \'a0 \'a0j = arrLogLines.Length - 1\
\
\'a0 \'a0 \'a0 \'a0' j is 3 because it starts from the bottom of the file and goes up until it reaches\
\'a0 \'a0 \'a0 \'a0' the files comments section\
\'a0 \'a0 \'a0 \'a0While j > 3\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0If Not arrLogLines(j) = String.Empty Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0' Just in case an exceptional case occurs with an additional log being appended\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0dt.Rows.Add(arrLogLines(j).Split(Convert.ToChar(" ")))\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0System.Math.Max(System.Threading.Interlocked.Decrement(j), i + 1)\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0Catch ex As Exception\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "Error on parsing log line. ")\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0End Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0End If\
\'a0 \'a0 \'a0 \'a0End While\
\
\'a0 \'a0 \'a0 \'a0' Some useful information for parsing\
\'a0 \'a0 \'a0 \'a0' 1 day = 86400 seconds\
\'a0 \'a0 \'a0 \'a0' 1 day = 1440 minutes\
\'a0 \'a0 \'a0 \'a0' 1 day = 24 hours\
\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click\
\
\'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0' Read the IIS logfile specified in the textbox\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0ReadLog(txtLogfile.Text)\
\
\
\'a0 \'a0 \'a0 \'a0Catch ex As Exception\
\
\'a0 \'a0 \'a0 \'a0End Try\
\'a0 \'a0End Sub\
\
\'a0 \'a0' This method will return an integer for time based on the radio button selected by the user\
\'a0 \'a0Private Function CalculatePeriod() As Integer\
\
\'a0 \'a0 \'a0 \'a0' the number of seconds in 1 day\
\
\'a0 \'a0 \'a0 \'a0Dim timeValue As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim summarySize As Integer = 0\
\
\'a0 \'a0 \'a0 \'a0' radio button control for time span\
\'a0 \'a0 \'a0 \'a0If rdoHours.Checked Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0timeValue = 3600\
\'a0 \'a0 \'a0 \'a0ElseIf rdoMinutes.Checked Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0timeValue = 60\
\'a0 \'a0 \'a0 \'a0ElseIf rdoSeconds.Checked Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0timeValue = 1\
\'a0 \'a0 \'a0 \'a0Else\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0' shouldn't get here\
\'a0 \'a0 \'a0 \'a0End If\
\
\'a0 \'a0 \'a0 \'a0' check that the timeValue is not 0, adn if not then calculate the summary(size)\
\'a0 \'a0 \'a0 \'a0If Not timeValue = 0 Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0summarySize = CInt(secondsInDay / timeValue)\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0' start from 1\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0summarySize = summarySize - 1\
\'a0 \'a0 \'a0 \'a0End If\
\
\'a0 \'a0 \'a0 \'a0Return summarySize\
\'a0 \'a0End Function\
\
\'a0 \'a0Private Sub ReadLog(ByVal fileName As String)\
\'a0 \'a0 \'a0 \'a0Dim streamReader As IO.StreamReader\
\'a0 \'a0 \'a0 \'a0Dim logLine As String = Nothing\
\'a0 \'a0 \'a0 \'a0Dim lineCount As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim count As Integer = 0\
\
\
\
\'a0 \'a0 \'a0 \'a0' need to reset all indexes to 0\
\
\'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0streamReader = IO.File.OpenText(fileName)\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0' read each line in the log\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Do While True\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0logLine = streamReader.ReadLine\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0If logLine Is Nothing Then Exit Do\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0logLine = logLine.Trim\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0lineCount += 1\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0' check if there are empty lines, or comments and skip them\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0If Not logLine = String.Empty AndAlso Not logLine.StartsWith("#") Then\
\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0End If\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Loop\
\
\'a0 \'a0 \'a0 \'a0Catch ex As Exception\
\
\'a0 \'a0 \'a0 \'a0End Try\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Sub ParseLine()\
\
\'a0 \'a0 \'a0 \'a0Dim httpResponse As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim timeOffset As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim timeSpan As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim serverBytes As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim clientBytes As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim responseTimeMs As Integer = 0\
\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Sub IntializeData()\
\
\'a0 \'a0 \'a0 \'a0' represents an offset line in log\
\'a0 \'a0 \'a0 \'a0Dim timeStamp As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim httpResponse As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim timespan As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim serverBytes As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim clientBytes As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim responseTime As Integer = 0\
\
\'a0 \'a0 \'a0 \'a0' create an array for each column that will be output\
\'a0 \'a0 \'a0 \'a0' Setup the size of each array to the lowest resolution - seconds in a(day)\
\'a0 \'a0 \'a0 \'a0Dim arraySize As Integer\
\'a0 \'a0 \'a0 \'a0arraySize = secondsInDay - 1\
\
\'a0 \'a0 \'a0 \'a0Dim concurrentUsers(arraySize) As Double\
\'a0 \'a0 \'a0 \'a0Dim httpSuccess(arraySize) As Double\
\'a0 \'a0 \'a0 \'a0Dim httpFailure(arraySize) As Double\
\'a0 \'a0 \'a0 \'a0Dim c_bytes(arraySize) As Double\
\'a0 \'a0 \'a0 \'a0Dim s_bytes(arraySize) As Double\
\
\
\'a0 \'a0 \'a0 \'a0' Setup offsets\
\'a0 \'a0 \'a0 \'a0timeStamp = 1\
\'a0 \'a0 \'a0 \'a0httpResponse = 9\
\'a0 \'a0 \'a0 \'a0serverBytes = 10\
\'a0 \'a0 \'a0 \'a0clientBytes = 11 ' apparently prod doesn't log client bytes, so this is server bytes\
\'a0 \'a0 \'a0 \'a0timespan = 11\
\'a0 \'a0 \'a0 \'a0responseTime = 12\
\
\'a0 \'a0 \'a0 \'a0' Pass in the value selected by the user - \'a0seconds, minutes, or hours()\
\'a0 \'a0 \'a0 \'a0Dim sumConcurrentUsers(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumSuccess(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumFailure(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumSbytes(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumCbytes(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumCdatarate(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumSdatarate(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim sumResponseTime(CalculatePeriod()) As Double\
\'a0 \'a0 \'a0 \'a0Dim maxAverageDataRate As Double = 0.0\
\'a0 \'a0 \'a0 \'a0Dim maxUsers As Double = 0.0\
\'a0 \'a0 \'a0 \'a0Dim maxErrors As Double = 0.0\
\
\'a0 \'a0End Sub\
\
\'a0 \'a0Private Function ParseLineItem(ByVal item As String, ByVal itemNumber As Integer) As String\
\
\'a0 \'a0 \'a0 \'a0Dim offset1 As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim counter As Integer = 0\
\'a0 \'a0 \'a0 \'a0Dim offset2 As Double = 0.0\
\'a0 \'a0 \'a0 \'a0Dim result As String = String.Empty\
\
\'a0 \'a0 \'a0 \'a0If itemNumber > 0 Then\
\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0For counter = 1 To itemNumber\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0offset1 = item.IndexOf(" ", offset1 + 1)\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0If offset1 < 0 Then\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0Throw New Exception("GetLineItem invalid itemno")\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0End If\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Next\
\
\'a0 \'a0 \'a0 \'a0End If\
\
\'a0 \'a0 \'a0 \'a0offset2 = item.IndexOf(" ", offset1 + 1)\
\
\'a0 \'a0 \'a0 \'a0'''result = item.Substring(offset1, offset2 - offset1)\
\
\'a0 \'a0 \'a0 \'a0Return result\
\
\'a0 \'a0End Function\
\
\'a0 \'a0Private Function ParseLineItemOffset(ByVal line As String, ByVal itemno As Integer) As Integer\
\
\'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0'Return TimeSpan.Parse(ParseLineItem(line, itemno)).TotalSeconds\
\'a0 \'a0 \'a0 \'a0Catch ex As Exception\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "ParseLineItemOffset")\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Throw ex\
\'a0 \'a0 \'a0 \'a0End Try\
\
\'a0 \'a0End Function\
\
\'a0 \'a0Private Function ParseLineItemDouble(ByVal line As String, ByVal itemno As Integer) As Integer\
\
\'a0 \'a0 \'a0 \'a0Try\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0'Return Double.Parse(ParseLineItem(line, itemno))\
\'a0 \'a0 \'a0 \'a0Catch ex As Exception\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "ParseLineItemOffset")\
\'a0 \'a0 \'a0 \'a0 \'a0 \'a0Throw ex\
\'a0 \'a0 \'a0 \'a0End Try\
\
\'a0 \'a0End Function\
\
\'a0 \'a0Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click\
\'a0 \'a0 \'a0 \'a0LoadInFile()\
\'a0 \'a0End Sub\
End Class}