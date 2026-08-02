Imports MySql.Data.MySqlClient
Imports System.Net.Http
Imports System.Text
Imports System.Text.Json

Public Module TelegramNotifier

    ' Bot Token for @Mycloud_pmsbot
    Private ReadOnly TelegramBotToken As String = "8849670353:AAG6aUf_AwokE1vs8hxW5lTDCFQK2PMMWY0"
    Private ReadOnly DefaultChatId As String = "-1002288339192"

    Public Sub SendAssignmentAlertsTelegram(connectionString As String)
        Console.WriteLine(vbLf & "--> Checking for newly assigned tickets to notify agents via Telegram...")

        ' SQL Query: Find un-notified active tickets directly from Zoho_Tickets_Staging
        Dim query As String = "
            SELECT 
                TicketId, 
                TicketNumber, 
                Subject, 
                Assignee,
                Status
            FROM Zoho_Tickets_Staging
            WHERE IFNULL(AssignmentNotified, 0) = 0 
              AND Assignee IS NOT NULL 
              AND Assignee <> '' 
              AND Assignee <> 'Unassigned'
              AND Status IN ('Open', 'Hold', 'Pending', 'Escalated')"

        Dim pendingNotifications As New List(Of TicketAlertDto)()

        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            ' 1. Fetch pending alerts
            Using cmd As New MySqlCommand(query, conn)
                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        pendingNotifications.Add(New TicketAlertDto With {
                            .TicketId = reader("TicketId").ToString(),
                            .TicketNumber = reader("TicketNumber").ToString(),
                            .Subject = reader("Subject").ToString(),
                            .AssigneeName = reader("Assignee").ToString(),
                            .ChatId = DefaultChatId
                        })
                    End While
                End Using
            End Using

            If pendingNotifications.Count = 0 Then
                Console.WriteLine("   [√] No new ticket assignments to notify.")
                Exit Sub
            End If

            ' 2. Dispatch messages via Telegram API
            Using client As New HttpClient()
                For Each item In pendingNotifications
                    Dim message As String = $"🚨 *New Ticket Assigned!*" & vbCrLf &
                                           $"*Ticket #:* {item.TicketNumber}" & vbCrLf &
                                           $"*Ticket ID:* {item.TicketId}" & vbCrLf &
                                           $"*Subject:* {item.Subject}" & vbCrLf &
                                           $"*Assigned To:* {item.AssigneeName}" & vbCrLf &
                                           $"[View Ticket in Zoho Desk](https://desk.zoho.com/support/tickets/{item.TicketId})"

                    Dim isSent As Boolean = SendTelegramMessage(client, item.ChatId, message)

                    ' 3. Mark as notified in SQL so it won't send duplicates
                    If isSent Then
                        Dim updateQuery As String = "UPDATE Zoho_Tickets_Staging SET AssignmentNotified = 1 WHERE TicketID = @TicketId"
                        Using updateCmd As New MySqlCommand(updateQuery, conn)
                            updateCmd.Parameters.AddWithValue("@TicketId", item.TicketId)
                            updateCmd.ExecuteNonQuery()
                        End Using
                        Console.WriteLine($"   [√] Telegram alert sent for Ticket #{item.TicketNumber} -> {item.AssigneeName}")
                    End If
                Next
            End Using
        End Using
    End Sub

    Private Function SendTelegramMessage(client As HttpClient, chatId As String, text As String) As Boolean
        Try
            Dim url As String = $"https://api.telegram.org/bot{TelegramBotToken}/sendMessage"
            Dim payload = New With {
                .chat_id = chatId,
                .text = text,
                .parse_mode = "Markdown",
                .disable_web_page_preview = False
            }

            Dim jsonPayload As String = JsonSerializer.Serialize(payload)
            Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

            Dim response As HttpResponseMessage = client.PostAsync(url, content).GetAwaiter().GetResult()
            Return response.IsSuccessStatusCode
        Catch ex As Exception
            Console.WriteLine($"   [!] Failed to send Telegram alert: {ex.Message}")
            Return False
        End Try
    End Function

    Private Class TicketAlertDto
        Public Property TicketId As String
        Public Property TicketNumber As String
        Public Property Subject As String
        Public Property AssigneeName As String
        Public Property ChatId As String
    End Class

End Module