Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
Imports MySql.Data.MySqlClient

Public Class TelegramNotifier

    ' =========================================================
    ' TELEGRAM CONFIGURATION
    ' =========================================================
    Private Shared ReadOnly TelegramBotToken As String = "8849670353:AAG6aUf_AwokE1vs8hxW5lTDCFQK2PMMWY0"

    ''' <summary>
    ''' Checks TiDB Cloud database and dispatches Telegram reminder alerts for active tickets until Closed.
    ''' </summary>
    Public Shared Sub SendAssignmentAlertsTelegram(connString As String)
        Console.WriteLine("--> Checking for active open/assigned tickets to send Telegram reminders...")

        Dim pendingTickets As New List(Of (TicketID As String, TicketNumber As String, Subject As String, Assignee As String, Status As String, AgentChatId As String))()

        Try
            ' Explicitly using MySqlConnection for TiDB Cloud
            Using conn As New MySqlConnection(connString)
                conn.Open()

                ' Query active tickets where Status is NOT Closed or Resolved
                Dim selectSql As String = "
                    SELECT 
                        t.TicketID, 
                        t.TicketNumber, 
                        t.Subject, 
                        t.Assignee, 
                        t.Status, 
                        a.TelegramChatId
                    FROM Zoho_Tickets_Staging t
                    INNER JOIN Telegram_Agents a ON t.Assignee = a.AgentName
                    WHERE t.Status NOT IN ('Closed', 'Resolved')
                      AND t.Assignee IS NOT NULL 
                      AND t.Assignee <> '' 
                      AND t.Assignee <> 'Unassigned';"

                Using cmd As New MySqlCommand(selectSql, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            pendingTickets.Add((
                                reader("TicketID").ToString(),
                                reader("TicketNumber").ToString(),
                                reader("Subject").ToString(),
                                reader("Assignee").ToString(),
                                reader("Status").ToString(),
                                reader("TelegramChatId").ToString()
                            ))
                        End While
                    End Using
                End Using

                If pendingTickets.Count = 0 Then
                    Console.WriteLine("   [!] No active pending tickets found requiring alerts.")
                    Return
                End If

                Console.WriteLine($"   [+] Dispatching {pendingTickets.Count} active ticket reminder(s)...")

                ' Send Individual Telegram Notifications via JSON POST
                Using client As New HttpClient()
                    Dim telegramApiUrl As String = $"https://api.telegram.org/bot{TelegramBotToken}/sendMessage"

                    For Each t In pendingTickets
                        ' Direct Zoho Desk Ticket URL
                        Dim deskUrl As String = $"https://desk.zoho.in/agent/mycloud/zap/tickets/details/{t.TicketID}"

                        ' Clean HTML formatted message with direct link
                        Dim formattedMessage As String = $"<b>🚨 Ticket Reminder / Assignment Alert!</b>" & vbCrLf & vbCrLf &
                                                        $"<b>Ticket #:</b> <a href=""{deskUrl}"">{t.TicketNumber}</a>" & vbCrLf &
                                                        $"<b>Ticket ID:</b> {t.TicketID}" & vbCrLf &
                                                        $"<b>Subject:</b> {t.Subject}" & vbCrLf &
                                                        $"<b>Status:</b> {t.Status}" & vbCrLf &
                                                        $"<b>Assigned To:</b> {t.Assignee}" & vbCrLf & vbCrLf &
                                                        $"🔗 <a href=""{deskUrl}"">Open Ticket in Zoho Desk</a>"

                        ' Create clean JSON payload
                        Dim payloadObj = New With {
                            .chat_id = t.AgentChatId,
                            .text = formattedMessage,
                            .parse_mode = "HTML"
                        }

                        Dim jsonPayload As String = JsonSerializer.Serialize(payloadObj)
                        Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                        Try
                            Dim response = client.PostAsync(telegramApiUrl, content).GetAwaiter().GetResult()
                            Dim responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                            If response.IsSuccessStatusCode Then
                                Console.WriteLine($"   [✓] Telegram alert sent to {t.Assignee} (Chat ID: {t.AgentChatId}) for Ticket #{t.TicketNumber}")

                                ' Update flag for tracking sync timestamp
                                Dim updateSql As String = "UPDATE Zoho_Tickets_Staging SET AssignmentNotified = 1 WHERE TicketID = @TicketID;"
                                Using updateCmd As New MySqlCommand(updateSql, conn)
                                    updateCmd.Parameters.AddWithValue("@TicketID", t.TicketID)
                                    updateCmd.ExecuteNonQuery()
                                End Using
                            Else
                                Console.WriteLine($"   [X] Failed for {t.Assignee} (Ticket #{t.TicketNumber}). Response: {responseBody}")
                            End If
                        Catch exHttp As Exception
                            Console.WriteLine($"   [X] HTTP Exception for {t.Assignee}: {exHttp.Message}")
                        End Try
                    Next
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine($"   [X] Telegram Dispatcher Error: {ex.Message}")
        End Try
    End Sub

End Class