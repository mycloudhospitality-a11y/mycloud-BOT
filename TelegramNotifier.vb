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
    ''' Checks TiDB Cloud database for unnotified ticket assignments and dispatches individual Telegram alerts to mapped agents.
    ''' </summary>
    Public Shared Sub SendAssignmentAlertsTelegram(connString As String)
        Console.WriteLine("--> Checking for newly assigned tickets to notify agents via Telegram...")

        Dim pendingTickets As New List(Of (TicketID As String, TicketNumber As String, Subject As String, Assignee As String, Status As String, AgentChatId As String))()

        Try
            Using conn As New MySqlConnection(connString)
                conn.Open()

                ' JOIN Zoho_Tickets_Staging with Telegram_Agents to get individual Chat IDs
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
                    WHERE (t.AssignmentNotified IS NULL OR t.AssignmentNotified = 0)
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
                    Console.WriteLine("   [!] No new unnotified ticket assignments found for registered agents.")
                    Return
                End If

                Console.WriteLine($"   [+] Found {pendingTickets.Count} pending individual notification(s) to dispatch.")

                ' Send Individual Telegram Notifications via JSON POST
                Using client As New HttpClient()
                    Dim telegramApiUrl As String = $"https://api.telegram.org/bot{TelegramBotToken}/sendMessage"

                    For Each t In pendingTickets
                        ' Clean HTML formatted message with standard line breaks (\n)
                        Dim formattedMessage As String = $"<b>🚨 New Ticket Assigned to You!</b>" & vbCrLf & vbCrLf &
                                                        $"<b>Ticket #:</b> {t.TicketNumber}" & vbCrLf &
                                                        $"<b>Ticket ID:</b> {t.TicketID}" & vbCrLf &
                                                        $"<b>Subject:</b> {t.Subject}" & vbCrLf &
                                                        $"<b>Status:</b> {t.Status}" & vbCrLf &
                                                        $"<b>Assigned To:</b> {t.Assignee}"

                        ' Create clean JSON payload to avoid URL encoding issues with HTML tags
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

                                ' Mark as notified in SQL
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