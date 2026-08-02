Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text.Json
Imports System.Threading.Tasks
Imports MySql.Data.MySqlClient
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Module Module1

    ' =========================================================
    ' CONFIGURATION — ZOHO DESK & TIDB CLOUD CREDENTIALS
    ' =========================================================
    Private ClientId As String = "1000.U3SD5Z72T619AS9SD2SI79G9FD0DAY"
    Private ClientSecret As String = "9b3ad8bd4eeb60541e527d712424450eb46e2cc476"
    Private RefreshToken As String = "1000.778bd16ec6ba2c53d3d9558d27115dcf.c530214c022f43efe43f4e2f521d10c7"
    Private ZohoOrgId As String = "631978161"

    ' Zoho India (.in) Regional Endpoints
    Private ReadOnly TokenUrl As String = "https://accounts.zoho.com/oauth/v2/token"
    Private ReadOnly BaseDeskUrl As String = "https://desk.zoho.com/api/v1/tickets"

    ' TiDB Cloud Connection String (MySQL Compatible)
    Private SqlConnectionString As String = "server=gateway01.ap-southeast-1.prod.aws.tidbcloud.com;port=4000;database=sys;uid=3fvtL6XakG6M5TM.root;pwd=YOUR_GENERATED_PASSWORD;SslMode=Preferred;"

    ' --- TARGET EXCEL OUTPUT PATH ---
    Private ReadOnly ExcelOutputPath As String = "E:\ZohoSqlBot\Zoho_Desk_CEO_Dashboard.xlsx"

    ' --- DATE RANGE FILTER GLOBALS ---
    Private FilterEndDate As DateTime = DateTime.Today.AddDays(1).AddSeconds(-1)
    Private FilterStartDate As DateTime = DateTime.Today.AddDays(-10)

    Sub Main()
        Console.WriteLine("=========================================================")
        Console.WriteLine("   ZOHO DESK — EXECUTIVE CEO DASHBOARD GENERATOR (.in)   ")
        Console.WriteLine("=========================================================")
        Console.WriteLine()

        SetupTenDayDateFilter()

        Try
            Console.WriteLine("1. Getting Fresh Access Token from Zoho OAuth (.in)...")
            Dim freshAccessToken As String = GetFreshAccessTokenAsync().GetAwaiter().GetResult()
            Console.WriteLine("   [✓] Access token retrieved successfully.")
            Console.WriteLine()

            Console.WriteLine($"2. Fetching tickets for date range: {FilterStartDate:yyyy-MM-dd HH:mm:ss} to {FilterEndDate:yyyy-MM-dd HH:mm:ss}...")
            Dim totalSynced As Integer = FetchAndSyncTicketsByDateRange(freshAccessToken, ZohoOrgId, FilterStartDate, FilterEndDate)
            Console.WriteLine($"   --> SUCCESS: Total {totalSynced} tickets synced to TiDB Cloud!")
            Console.WriteLine()

            ' -------------------------------------------------------------
            ' TELEGRAM NOTIFICATION DISPATCHER (@Mycloud_pmsbot)
            ' -------------------------------------------------------------
            Console.WriteLine("3. Dispatching pending assignment notifications to Telegram...")
            TelegramNotifier.SendAssignmentAlertsTelegram(SqlConnectionString)
            Console.WriteLine()

            Console.WriteLine("4. Generating Professional CEO Excel Dashboard with Formulas...")
            GenerateCeoDashboardFromSql()

            Console.WriteLine()
            Console.WriteLine($"=========================================================")
            Console.WriteLine($"SUCCESS! Executive Dashboard generated at:")
            Console.WriteLine($"{ExcelOutputPath}")
            Console.WriteLine($"=========================================================")

        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine($"CRITICAL ERROR: {ex.Message}")
            Console.WriteLine(ex.StackTrace)
        End Try

        Console.WriteLine()
        Console.WriteLine("Press any key to exit...")
        Console.ReadKey()
    End Sub

    ' --- SET UP 10-DAY DATE FILTER ---
    Private Sub SetupTenDayDateFilter()
        Console.WriteLine("--- AUTOMATIC DATE FILTER (LAST 10 DAYS) ---")
        FilterEndDate = DateTime.Today.AddDays(1).AddSeconds(-1)
        FilterStartDate = DateTime.Today.AddDays(-10)

        Console.WriteLine($"   [+] Start Date Range : {FilterStartDate:yyyy-MM-dd HH:mm:ss}")
        Console.WriteLine($"   [+] End Date Range   : {FilterEndDate:yyyy-MM-dd HH:mm:ss}")
        Console.WriteLine()
    End Sub

    ' =========================================================
    ' STEP 1: TOKEN REFRESH (.in INDIA REGION)
    ' =========================================================
    Private Async Function GetFreshAccessTokenAsync() As Task(Of String)
        Using client As New HttpClient()
            Dim values = New Dictionary(Of String, String) From {
                {"refresh_token", RefreshToken},
                {"client_id", ClientId},
                {"client_secret", ClientSecret},
                {"grant_type", "refresh_token"}
            }

            Dim response As HttpResponseMessage = Await client.PostAsync(TokenUrl, New FormUrlEncodedContent(values))
            Dim responseString As String = Await response.Content.ReadAsStringAsync()

            Using doc As JsonDocument = JsonDocument.Parse(responseString)
                Dim root As JsonElement = doc.RootElement
                Dim accessTokenElem As JsonElement = Nothing

                If root.TryGetProperty("access_token", accessTokenElem) Then
                    Return accessTokenElem.GetString()
                Else
                    Throw New Exception($"Token refresh failed: {responseString}")
                End If
            End Using
        End Using
    End Function

    ' =========================================================
    ' STEP 2: PAGINATED TICKET FETCH (STANDARD ENDPOINT WITH OFFSET)
    ' =========================================================
    Private Function FetchAndSyncTicketsByDateRange(accessToken As String, orgId As String, startDate As DateTime, endDate As DateTime) As Integer
        Dim totalSaved As Integer = 0
        Dim fromIndex As Integer = 1
        Dim limit As Integer = 100
        Dim hasMoreRecords As Boolean = True
        Dim lastFirstTicketId As String = ""

        Using client As New HttpClient()
            While hasMoreRecords
                Dim requestUrl As String = $"{BaseDeskUrl}?include=assignee,products&from={fromIndex}&limit={limit}"

                Using request As New HttpRequestMessage(HttpMethod.Get, requestUrl)
                    request.Headers.TryAddWithoutValidation("orgId", orgId)
                    request.Headers.TryAddWithoutValidation("Authorization", "Zoho-oauthtoken " & accessToken)

                    Dim response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                    Dim jsonText As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                    If Not response.IsSuccessStatusCode Then
                        Throw New Exception($"Zoho API Call Error ({response.StatusCode}): {jsonText}")
                    End If

                    Dim currentFirstTicketId As String = ""
                    Using doc As JsonDocument = JsonDocument.Parse(jsonText)
                        Dim root = doc.RootElement
                        If root.ValueKind = JsonValueKind.Array AndAlso root.GetArrayLength() > 0 Then
                            currentFirstTicketId = GetJsonPropString(root(0), "id", "")
                        ElseIf root.ValueKind = JsonValueKind.Object AndAlso root.TryGetProperty("data", Nothing) Then
                            Dim dataArr = root.GetProperty("data")
                            If dataArr.ValueKind = JsonValueKind.Array AndAlso dataArr.GetArrayLength() > 0 Then
                                currentFirstTicketId = GetJsonPropString(dataArr(0), "id", "")
                            End If
                        End If
                    End Using

                    If Not String.IsNullOrEmpty(currentFirstTicketId) AndAlso currentFirstTicketId = lastFirstTicketId Then
                        Console.WriteLine("   [!] Detected duplicate API page from Zoho. Stopping fetch.")
                        Exit While
                    End If
                    lastFirstTicketId = currentFirstTicketId

                    Dim itemsInBatch As Integer = 0
                    Dim batchSaved As Integer = ProcessAndSaveBatchToSql(jsonText, itemsInBatch)

                    If itemsInBatch = 0 Then
                        Console.WriteLine($"   Reached end of records at index {fromIndex}. Stopping fetch.")
                        hasMoreRecords = False
                    Else
                        totalSaved += batchSaved
                        Console.WriteLine($"   Fetched Range {fromIndex}-{fromIndex + itemsInBatch - 1} | Records received: {itemsInBatch} | Synced to Cloud: {batchSaved} | Cumulative Total: {totalSaved}")

                        If batchSaved = 0 Then
                            Console.WriteLine("   [!] Batch yielded 0 synced records within target date range. Stopping fetch.")
                            hasMoreRecords = False
                        ElseIf itemsInBatch < limit Then
                            Console.WriteLine("   Last batch reached (batch size < limit). Stopping fetch.")
                            hasMoreRecords = False
                        Else
                            fromIndex += limit
                        End If
                    End If
                End Using
            End While
        End Using

        Return totalSaved
    End Function

    ' =========================================================
    ' STEP 3: MYSQL STAGING UPSERT (DUPLICATE KEY UPDATE)
    ' =========================================================
    Private Function ProcessAndSaveBatchToSql(jsonText As String, ByRef itemsInBatch As Integer) As Integer
        Dim savedCount As Integer = 0
        itemsInBatch = 0

        Using doc As JsonDocument = JsonDocument.Parse(jsonText)
            Dim root As JsonElement = doc.RootElement
            Dim ticketsArray As JsonElement

            If root.ValueKind = JsonValueKind.Array Then
                ticketsArray = root
            ElseIf root.ValueKind = JsonValueKind.Object AndAlso root.TryGetProperty("data", ticketsArray) AndAlso ticketsArray.ValueKind = JsonValueKind.Array Then
                ' Extracted array from data wrapper
            Else
                itemsInBatch = 0
                Return 0
            End If

            itemsInBatch = ticketsArray.GetArrayLength()
            If itemsInBatch = 0 Then Return 0

            Using conn As New MySqlConnection(SqlConnectionString)
                conn.Open()
                Using transaction As MySqlTransaction = conn.BeginTransaction()
                    Dim query As String = "
                        INSERT INTO Zoho_Tickets_Staging (TicketID, TicketNumber, Subject, Status, Priority, Assignee, AssigneeId, Category, Product, CreatedTime, ClosedTime, ResolutionTimeHours) 
                        VALUES (@TicketID, @TicketNumber, @Subject, @Status, @Priority, @Assignee, @AssigneeId, @Category, @Product, @CreatedTime, @ClosedTime, @ResolutionTimeHours)
                        ON DUPLICATE KEY UPDATE 
                            Status = VALUES(Status), 
                            Priority = VALUES(Priority), 
                            Subject = VALUES(Subject),
                            Assignee = VALUES(Assignee),
                            AssigneeId = VALUES(AssigneeId),
                            Category = VALUES(Category),
                            Product = VALUES(Product),
                            ClosedTime = VALUES(ClosedTime),
                            ResolutionTimeHours = VALUES(ResolutionTimeHours),
                            LastUpdated = CURRENT_TIMESTAMP;"

                    For Each ticket As JsonElement In ticketsArray.EnumerateArray()
                        Dim ticketId As String = GetJsonPropString(ticket, "id", "")
                        Dim ticketNumber As String = GetJsonPropString(ticket, "ticketNumber", "")

                        If String.IsNullOrEmpty(ticketId) Then Continue For

                        Dim subject As String = GetJsonPropString(ticket, "subject", "No Subject")
                        Dim status As String = GetJsonPropString(ticket, "status", "Open")
                        Dim priority As String = GetJsonPropString(ticket, "priority", "Medium")
                        Dim category As String = GetJsonPropString(ticket, "category", "General")
                        Dim productName As String = GetJsonPropString(ticket, "productName", "General Product")

                        ' Extract Assignee Details (ID and Name)
                        Dim assigneeName As String = "Unassigned"
                        Dim assigneeId As String = ""
                        Dim assigneeElem As JsonElement = Nothing

                        If ticket.TryGetProperty("assignee", assigneeElem) AndAlso assigneeElem.ValueKind = JsonValueKind.Object Then
                            assigneeId = GetJsonPropString(assigneeElem, "id", "")
                            Dim fn As String = GetJsonPropString(assigneeElem, "firstName", "")
                            Dim ln As String = GetJsonPropString(assigneeElem, "lastName", "")
                            assigneeName = (fn & " " & ln).Trim()
                            If String.IsNullOrEmpty(assigneeName) Then assigneeName = GetJsonPropString(assigneeElem, "email", "Unassigned")
                        End If

                        ' Extract Dates
                        Dim createdTime As DateTime = DateTime.Now
                        Dim createdElem As JsonElement = Nothing
                        If ticket.TryGetProperty("createdTime", createdElem) AndAlso createdElem.ValueKind <> JsonValueKind.Null Then
                            DateTime.TryParse(createdElem.GetString(), createdTime)
                        End If

                        If createdTime < FilterStartDate OrElse createdTime > FilterEndDate Then
                            Continue For
                        End If

                        Dim closedTime As Object = DBNull.Value
                        Dim closedElem As JsonElement = Nothing
                        If ticket.TryGetProperty("closedTime", closedElem) AndAlso closedElem.ValueKind <> JsonValueKind.Null Then
                            Dim dt As DateTime
                            If DateTime.TryParse(closedElem.GetString(), dt) Then closedTime = dt
                        End If

                        ' Resolution calculation
                        Dim resolutionHours As Double = 0
                        If closedTime IsNot DBNull.Value Then
                            resolutionHours = Math.Round((CType(closedTime, DateTime) - createdTime).TotalHours, 2)
                        Else
                            resolutionHours = Math.Round((DateTime.Now - createdTime).TotalHours, 2)
                        End If

                        Using cmd As New MySqlCommand(query, conn, transaction)
                            cmd.Parameters.AddWithValue("@TicketID", ticketId)
                            cmd.Parameters.AddWithValue("@TicketNumber", ticketNumber)
                            cmd.Parameters.AddWithValue("@Subject", subject)
                            cmd.Parameters.AddWithValue("@Status", status)
                            cmd.Parameters.AddWithValue("@Priority", priority)
                            cmd.Parameters.AddWithValue("@Assignee", assigneeName)
                            cmd.Parameters.AddWithValue("@AssigneeId", If(String.IsNullOrEmpty(assigneeId), DBNull.Value, CObj(assigneeId)))
                            cmd.Parameters.AddWithValue("@Category", category)
                            cmd.Parameters.AddWithValue("@Product", productName)
                            cmd.Parameters.AddWithValue("@CreatedTime", createdTime)
                            cmd.Parameters.AddWithValue("@ClosedTime", closedTime)
                            cmd.Parameters.AddWithValue("@ResolutionTimeHours", resolutionHours)
                            cmd.ExecuteNonQuery()
                            savedCount += 1
                        End Using
                    Next
                    transaction.Commit()
                End Using
            End Using
        End Using
        Return savedCount
    End Function

    Private Function GetJsonPropString(elem As JsonElement, propName As String, defaultValue As String) As String
        Dim val As JsonElement = Nothing
        If elem.TryGetProperty(propName, val) AndAlso val.ValueKind <> JsonValueKind.Null Then
            Return val.ToString()
        End If
        Return defaultValue
    End Function

    ' =========================================================
    ' STEP 4: BUILD CEO EXECUTIVE EXCEL DASHBOARD (EPPlus)
    ' =========================================================
    Private Sub GenerateCeoDashboardFromSql()
        ExcelPackage.License.SetNonCommercialPersonal("ZohoSqlBot")

        Dim fileInfo As New FileInfo(ExcelOutputPath)
        If fileInfo.Directory IsNot Nothing AndAlso Not fileInfo.Directory.Exists Then
            fileInfo.Directory.Create()
        End If

        If fileInfo.Exists Then
            Try
                fileInfo.Delete()
            Catch ex As Exception
                Throw New Exception($"Please close '{ExcelOutputPath}' in Excel and run again.")
            End Try
        End If

        Using package As New ExcelPackage()
            ' ---------------------------------------------------------
            ' TAB 1: CEO EXECUTIVE SUMMARY
            ' ---------------------------------------------------------
            Dim ws1 = package.Workbook.Worksheets.Add("CEO Executive Summary")
            ws1.View.ShowGridLines = True

            ws1.Column(1).Width = 3
            ws1.Column(2).Width = 26
            ws1.Column(3).Width = 15

            ' Title Banner
            ws1.Cells("B2:M2").Merge = True
            ws1.Cells("B2").Value = "ZOHO DESK — EXECUTIVE HELPDESK PERFORMANCE REPORT"
            ws1.Cells("B2").Style.Font.Size = 15
            ws1.Cells("B2").Style.Font.Bold = True
            ws1.Cells("B2").Style.Font.Color.SetColor(System.Drawing.Color.White)
            ws1.Cells("B2").Style.Fill.PatternType = ExcelFillStyle.Solid
            ws1.Cells("B2").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(15, 32, 67))
            ws1.Cells("B2").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            Using conn As New MySqlConnection(SqlConnectionString)
                conn.Open()

                Dim distinctStatuses As New List(Of String)()
                Dim statusQuery As String = "SELECT DISTINCT Status FROM Zoho_Tickets_Staging WHERE CreatedTime BETWEEN @StartDate AND @EndDate ORDER BY Status;"
                Using cmdStatus As New MySqlCommand(statusQuery, conn)
                    cmdStatus.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmdStatus.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmdStatus.ExecuteReader()
                        While r.Read()
                            distinctStatuses.Add(r("Status").ToString())
                        End While
                    End Using
                End Using

                Dim totalTickets As Integer = 0
                Dim closedTickets As Integer = 0
                Dim avgResTime As Double = 0.0

                Dim kpiQuery As String = "
                    SELECT 
                        COUNT(*) AS Total,
                        SUM(CASE WHEN Status IN ('Closed','Resolved') THEN 1 ELSE 0 END) AS ClosedCount,
                        AVG(CASE WHEN Status IN ('Closed','Resolved') THEN ResolutionTimeHours END) AS AvgResolutionHours
                    FROM Zoho_Tickets_Staging
                    WHERE CreatedTime BETWEEN @StartDate AND @EndDate;"

                Using cmd As New MySqlCommand(kpiQuery, conn)
                    cmd.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmd.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then
                            totalTickets = If(IsDBNull(r("Total")), 0, Convert.ToInt32(r("Total")))
                            closedTickets = If(IsDBNull(r("ClosedCount")), 0, Convert.ToInt32(r("ClosedCount")))
                            avgResTime = If(IsDBNull(r("AvgResolutionHours")), 0.0, Math.Round(Convert.ToDouble(r("AvgResolutionHours")), 1))
                        End If
                    End Using
                End Using

                CreateKpiCard(ws1, "B4:C4", "B5:C5", "TOTAL VOLUME", totalTickets.ToString(), System.Drawing.Color.FromArgb(240, 244, 248), System.Drawing.Color.FromArgb(15, 32, 67))
                CreateKpiCard(ws1, "D4:E4", "D5:E5", "TOTAL CLOSED", closedTickets.ToString(), System.Drawing.Color.FromArgb(235, 247, 238), System.Drawing.Color.FromArgb(34, 112, 62))
                CreateKpiCard(ws1, "F4:H4", "F5:H5", "AVG CLOSING TIME (CLOSED TICKETS)", avgResTime.ToString() & " hrs", System.Drawing.Color.FromArgb(240, 244, 248), System.Drawing.Color.FromArgb(15, 32, 67))

                ' --- SECTION 1: AGENT WISE STATUS BREAKDOWN ---
                ws1.Cells("B7").Value = "AGENT PERFORMANCE & STATUS WISE BREAKUP"
                ws1.Cells("B7").Style.Font.Bold = True
                ws1.Cells("B7").Style.Font.Size = 11
                ws1.Cells("B7").Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(15, 32, 67))

                Dim colIdx As Integer = 2
                ws1.Cells(8, colIdx).Value = "Agent Name" : ws1.Column(colIdx).Width = 24 : colIdx += 1
                ws1.Cells(8, colIdx).Value = "Total Handled" : ws1.Column(colIdx).Width = 14 : colIdx += 1

                For Each st In distinctStatuses
                    ws1.Cells(8, colIdx).Value = st
                    ws1.Column(colIdx).Width = 14
                    colIdx += 1
                Next

                ws1.Cells(8, colIdx).Value = "Total Closed Hrs" : ws1.Column(colIdx).Width = 18 : colIdx += 1
                ws1.Cells(8, colIdx).Value = "Total Open Hrs" : ws1.Column(colIdx).Width = 18 : colIdx += 1
                ws1.Cells(8, colIdx).Value = "Avg Closing Time" : ws1.Column(colIdx).Width = 20 : colIdx += 1

                Dim totalAgentCols As Integer = colIdx - 1
                For c As Integer = 2 To totalAgentCols
                    Dim cell = ws1.Cells(8, c)
                    cell.Style.Font.Bold = True
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(31, 78, 120))
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next

                Dim statusPivotSql As String = ""
                For Each st In distinctStatuses
                    statusPivotSql &= $", SUM(CASE WHEN Status = '{st.Replace("'", "''")}' THEN 1 ELSE 0 END) AS `{st}`"
                Next

                Dim agentQuery As String = $"
                    SELECT Assignee, 
                           COUNT(*) AS TotalHandled
                           {statusPivotSql},
                           SUM(CASE WHEN Status IN ('Closed','Resolved') THEN ResolutionTimeHours ELSE 0 END) AS TotalClosedHours,
                           SUM(CASE WHEN Status NOT IN ('Closed','Resolved') THEN ResolutionTimeHours ELSE 0 END) AS TotalOpenHours
                    FROM Zoho_Tickets_Staging 
                    WHERE CreatedTime BETWEEN @StartDate AND @EndDate
                    GROUP BY Assignee 
                    ORDER BY TotalHandled DESC;"

                Dim rIdx As Integer = 9
                Using cmdAgent As New MySqlCommand(agentQuery, conn)
                    cmdAgent.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmdAgent.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmdAgent.ExecuteReader()
                        While r.Read()
                            Dim curCol As Integer = 2
                            ws1.Cells(rIdx, curCol).Value = r("Assignee").ToString() : curCol += 1
                            ws1.Cells(rIdx, curCol).Value = Convert.ToInt32(r("TotalHandled")) : curCol += 1

                            Dim closedColLetter As String = ""
                            For Each st In distinctStatuses
                                ws1.Cells(rIdx, curCol).Value = Convert.ToInt32(r(st))
                                If st.Equals("Closed", StringComparison.OrdinalIgnoreCase) OrElse st.Equals("Resolved", StringComparison.OrdinalIgnoreCase) Then
                                    closedColLetter = GetColumnAddress(curCol)
                                End If
                                curCol += 1
                            Next

                            Dim closedHrsColLetter As String = GetColumnAddress(curCol)
                            ws1.Cells(rIdx, curCol).Value = Math.Round(Convert.ToDouble(r("TotalClosedHours")), 1) : curCol += 1

                            ws1.Cells(rIdx, curCol).Value = Math.Round(Convert.ToDouble(r("TotalOpenHours")), 1) : curCol += 1

                            If Not String.IsNullOrEmpty(closedColLetter) Then
                                ws1.Cells(rIdx, curCol).Formula = $"IF({closedColLetter}{rIdx}>0, {closedHrsColLetter}{rIdx}/{closedColLetter}{rIdx}, 0)"
                            Else
                                ws1.Cells(rIdx, curCol).Value = 0
                            End If
                            ws1.Cells(rIdx, curCol).Style.Numberformat.Format = "0.0 ""hrs"""

                            ws1.Cells(rIdx, 3, rIdx, curCol).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                            rIdx += 1
                        End While
                    End Using
                End Using

                ' --- SECTION 2: CATEGORY BREAKDOWN ---
                Dim catStartRow As Integer = rIdx + 2
                ws1.Cells(catStartRow, 2).Value = "CATEGORY / DEPARTMENT BREAKDOWN ANALYSIS"
                ws1.Cells(catStartRow, 2).Style.Font.Bold = True
                ws1.Cells(catStartRow, 2).Style.Font.Size = 11
                ws1.Cells(catStartRow, 2).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(15, 32, 67))

                Dim catHeaderRow As Integer = catStartRow + 1
                Dim catColIdx As Integer = 2
                ws1.Cells(catHeaderRow, catColIdx).Value = "Category" : catColIdx += 1
                ws1.Cells(catHeaderRow, catColIdx).Value = "Total Volume" : catColIdx += 1

                For Each st In distinctStatuses
                    ws1.Cells(catHeaderRow, catColIdx).Value = st
                    catColIdx += 1
                Next
                ws1.Cells(catHeaderRow, catColIdx).Value = "% Share"

                For c As Integer = 2 To catColIdx
                    Dim cell = ws1.Cells(catHeaderRow, c)
                    cell.Style.Font.Bold = True
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(70, 95, 133))
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next

                Dim catQuery As String = $"
                    SELECT Category, 
                           COUNT(*) AS Volume
                           {statusPivotSql}
                    FROM Zoho_Tickets_Staging 
                    WHERE CreatedTime BETWEEN @StartDate AND @EndDate
                    GROUP BY Category 
                    ORDER BY Volume DESC;"

                Dim catDataRow As Integer = catHeaderRow + 1
                Using cmdCat As New MySqlCommand(catQuery, conn)
                    cmdCat.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmdCat.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmdCat.ExecuteReader()
                        While r.Read()
                            Dim curCol As Integer = 2
                            ws1.Cells(catDataRow, curCol).Value = r("Category").ToString() : curCol += 1
                            ws1.Cells(catDataRow, curCol).Value = Convert.ToInt32(r("Volume")) : curCol += 1

                            For Each st In distinctStatuses
                                ws1.Cells(catDataRow, curCol).Value = Convert.ToInt32(r(st))
                                curCol += 1
                            Next

                            ws1.Cells(catDataRow, curCol).Formula = $"C{catDataRow}/SUM(C${catHeaderRow + 1}:C${catDataRow + 50})"
                            ws1.Cells(catDataRow, curCol).Style.Numberformat.Format = "0.0%"

                            ws1.Cells(catDataRow, 3, catDataRow, curCol).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                            catDataRow += 1
                        End While
                    End Using
                End Using

                ' --- SECTION 3: PRODUCT BREAKDOWN ---
                Dim prodStartRow As Integer = catDataRow + 2
                ws1.Cells(prodStartRow, 2).Value = "PRODUCT-WISE BREAKDOWN ANALYSIS"
                ws1.Cells(prodStartRow, 2).Style.Font.Bold = True
                ws1.Cells(prodStartRow, 2).Style.Font.Size = 11
                ws1.Cells(prodStartRow, 2).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(15, 32, 67))

                Dim prodHeaderRow As Integer = prodStartRow + 1
                Dim prodColIdx As Integer = 2
                ws1.Cells(prodHeaderRow, prodColIdx).Value = "Product Name" : prodColIdx += 1
                ws1.Cells(prodHeaderRow, prodColIdx).Value = "Total Volume" : prodColIdx += 1

                For Each st In distinctStatuses
                    ws1.Cells(prodHeaderRow, prodColIdx).Value = st
                    prodColIdx += 1
                Next
                ws1.Cells(prodHeaderRow, prodColIdx).Value = "% Share"

                For c As Integer = 2 To prodColIdx
                    Dim cell = ws1.Cells(prodHeaderRow, c)
                    cell.Style.Font.Bold = True
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(41, 128, 185))
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next

                Dim prodQuery As String = $"
                    SELECT Product, 
                           COUNT(*) AS Volume
                           {statusPivotSql}
                    FROM Zoho_Tickets_Staging 
                    WHERE CreatedTime BETWEEN @StartDate AND @EndDate
                    GROUP BY Product 
                    ORDER BY Volume DESC;"

                Dim prodDataRow As Integer = prodHeaderRow + 1
                Using cmdProd As New MySqlCommand(prodQuery, conn)
                    cmdProd.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmdProd.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmdProd.ExecuteReader()
                        While r.Read()
                            Dim curCol As Integer = 2
                            ws1.Cells(prodDataRow, curCol).Value = r("Product").ToString() : curCol += 1
                            ws1.Cells(prodDataRow, curCol).Value = Convert.ToInt32(r("Volume")) : curCol += 1

                            For Each st In distinctStatuses
                                ws1.Cells(prodDataRow, curCol).Value = Convert.ToInt32(r(st))
                                curCol += 1
                            Next

                            ws1.Cells(prodDataRow, curCol).Formula = $"C{prodDataRow}/SUM(C${prodHeaderRow + 1}:C${prodDataRow + 50})"
                            ws1.Cells(prodDataRow, curCol).Style.Numberformat.Format = "0.0%"

                            ws1.Cells(prodDataRow, 3, prodDataRow, curCol).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                            prodDataRow += 1
                        End While
                    End Using
                End Using

            End Using

            ' ---------------------------------------------------------
            ' TAB 2: LIVE TICKET REGISTER
            ' ---------------------------------------------------------
            Dim ws2 = package.Workbook.Worksheets.Add("Live Ticket Register")
            ws2.View.ShowGridLines = True

            Dim headers As String() = {"Ticket ID", "Number", "Subject", "Status", "Priority", "Assignee", "Category", "Product", "Created Time", "Closed Time", "Resolution Time (Hrs)"}
            For i As Integer = 0 To headers.Length - 1
                Dim cell = ws2.Cells(1, i + 1)
                cell.Value = headers(i)
                cell.Style.Font.Bold = True
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(15, 32, 67))
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            Next

            Using conn As New MySqlConnection(SqlConnectionString)
                conn.Open()
                Dim detailQuery As String = "
                    SELECT TicketID, TicketNumber, Subject, Status, Priority, Assignee, Category, Product, CreatedTime, ClosedTime, ResolutionTimeHours 
                    FROM Zoho_Tickets_Staging 
                    WHERE CreatedTime BETWEEN @StartDate AND @EndDate 
                    ORDER BY CreatedTime DESC;"

                Using cmdDetail As New MySqlCommand(detailQuery, conn)
                    cmdDetail.Parameters.AddWithValue("@StartDate", FilterStartDate)
                    cmdDetail.Parameters.AddWithValue("@EndDate", FilterEndDate)
                    Using r = cmdDetail.ExecuteReader()
                        Dim rowNum As Integer = 2
                        While r.Read()
                            ws2.Cells(rowNum, 1).Value = r("TicketID").ToString()
                            ws2.Cells(rowNum, 2).Value = r("TicketNumber").ToString()
                            ws2.Cells(rowNum, 3).Value = r("Subject").ToString()
                            ws2.Cells(rowNum, 4).Value = r("Status").ToString()
                            ws2.Cells(rowNum, 5).Value = r("Priority").ToString()
                            ws2.Cells(rowNum, 6).Value = r("Assignee").ToString()
                            ws2.Cells(rowNum, 7).Value = r("Category").ToString()
                            ws2.Cells(rowNum, 8).Value = r("Product").ToString()

                            ws2.Cells(rowNum, 9).Value = Convert.ToDateTime(r("CreatedTime"))
                            ws2.Cells(rowNum, 9).Style.Numberformat.Format = "yyyy-MM-dd HH:mm"

                            If Not IsDBNull(r("ClosedTime")) Then
                                ws2.Cells(rowNum, 10).Value = Convert.ToDateTime(r("ClosedTime"))
                                ws2.Cells(rowNum, 10).Style.Numberformat.Format = "yyyy-MM-dd HH:mm"
                            Else
                                ws2.Cells(rowNum, 10).Value = "N/A"
                            End If

                            ws2.Cells(rowNum, 11).Value = Convert.ToDouble(r("ResolutionTimeHours"))
                            ws2.Cells(rowNum, 11).Style.Numberformat.Format = "0.0"

                            rowNum += 1
                        End While
                    End Using
                End Using
            End Using

            If ws2.Dimension IsNot Nothing Then
                ws2.Cells(ws2.Dimension.Address).AutoFitColumns()
            End If

            package.SaveAs(fileInfo)
        End Using
    End Sub

    ' =========================================================
    ' HELPER FUNCTIONS FOR EXCEL FORMATTING
    ' =========================================================
    Private Sub CreateKpiCard(ws As ExcelWorksheet, headerRange As String, bodyRange As String, title As String, value As String, bgColor As System.Drawing.Color, textColor As System.Drawing.Color)
        ws.Cells(headerRange).Merge = True
        ws.Cells(headerRange).Value = title
        ws.Cells(headerRange).Style.Font.Size = 9
        ws.Cells(headerRange).Style.Font.Bold = True
        ws.Cells(headerRange).Style.Font.Color.SetColor(textColor)
        ws.Cells(headerRange).Style.Fill.PatternType = ExcelFillStyle.Solid
        ws.Cells(headerRange).Style.Fill.BackgroundColor.SetColor(bgColor)
        ws.Cells(headerRange).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

        ws.Cells(bodyRange).Merge = True
        ws.Cells(bodyRange).Value = value
        ws.Cells(bodyRange).Style.Font.Size = 16
        ws.Cells(bodyRange).Style.Font.Bold = True
        ws.Cells(bodyRange).Style.Font.Color.SetColor(textColor)
        ws.Cells(bodyRange).Style.Fill.PatternType = ExcelFillStyle.Solid
        ws.Cells(bodyRange).Style.Fill.BackgroundColor.SetColor(bgColor)
        ws.Cells(bodyRange).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    End Sub

    Private Function GetColumnAddress(colIndex As Integer) As String
        Dim dividend As Integer = colIndex
        Dim columnName As String = String.Empty
        Dim modifier As Integer

        While dividend > 0
            modifier = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modifier) & columnName
            dividend = CInt((dividend - modifier) / 26)
        End While

        Return columnName
    End Function

End Module