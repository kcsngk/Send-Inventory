Imports SendInventory.com.netsuite.webservices
Imports OfficeOpenXml

Module SendInventory

	Private ns As New NetSuiteService

	Private LOOKUP_MFR As New Dictionary(Of Integer, String)
	Private LOOKUP_CLASS As New Dictionary(Of Integer, String)

	Private Sub MakeLookups()
		Dim rr As ReadResponse = ns.get(New RecordRef With {.internalId = "60", .type = RecordType.customList, .typeSpecified = True})
		For Each cv As CustomListCustomValue In CType(rr.record, CustomList).customValueList.customValue
			LOOKUP_MFR(cv.valueId) = cv.value
		Next

		rr = ns.get(New RecordRef With {.internalId = "241", .type = RecordType.customList, .typeSpecified = True})
		For Each cv As CustomListCustomValue In CType(rr.record, CustomList).customValueList.customValue
			LOOKUP_CLASS(cv.valueId) = cv.value
		Next
	End Sub

	Sub Main()

#Region " build dataset "
		Dim dsWorkbook As New DataSet("Inventory")

		Dim sheet_general As New DataTable("General")
		sheet_general.Columns.Add("Location", GetType(String))
		sheet_general.Columns.Add("Manufacturer", GetType(String))
		sheet_general.Columns.Add("Category", GetType(String))
		sheet_general.Columns.Add("C2PN", GetType(String))
		sheet_general.Columns.Add("MPN", GetType(String))
		sheet_general.Columns.Add("Description", GetType(String))
		sheet_general.Columns.Add("Available", GetType(Integer))
		sheet_general.Columns.Add("On Hand", GetType(Integer))
		sheet_general.Columns.Add("Committed", GetType(Integer))
		sheet_general.Columns.Add("Backordered", GetType(Integer))
		sheet_general.Columns.Add("On Order", GetType(Integer))
        sheet_general.Columns.Add("Avg Cost", GetType(Double))
        sheet_general.Columns.Add("Avg 2wk Selling", GetType(Double))
        'sheet_general.Columns.Add("D1", GetType(Double))
        'sheet_general.Columns.Add("M1", GetType(Double))
        'sheet_general.Columns.Add("R1", GetType(Double))
        'sheet_general.Columns.Add("R2", GetType(Double))

        sheet_general.Columns.Add("I1", GetType(Double))
        sheet_general.Columns.Add("C1", GetType(Double))
        sheet_general.Columns.Add("D1", GetType(Double))
        sheet_general.Columns.Add("R2", GetType(Double))
        sheet_general.Columns.Add("R1", GetType(Double))
        sheet_general.Columns.Add("B1", GetType(Double))
        sheet_general.Columns.Add("M1", GetType(Double))

        sheet_general.Columns.Add("MSRP", GetType(Double))
        sheet_general.Columns.Add("UPC", GetType(String))
		sheet_general.Columns.Add("LastActivity", GetType(Date))
		sheet_general.Columns.Add("Class", GetType(String))
        sheet_general.Columns.Add("DateCreated", GetType(Date))

        Dim sheet_vmi As DataTable = sheet_general.Clone()
		sheet_vmi.TableName = "VMI"
		Dim sheet_moulton As DataTable = sheet_general.Clone()
		sheet_moulton.TableName = "Moulton"
		Dim sheet_rma As DataTable = sheet_general.Clone()
		sheet_rma.TableName = "RMA"
		Dim sheet_vra As DataTable = sheet_general.Clone()
		sheet_vra.TableName = "VRA"
        'Dim sheet_noact As DataTable = sheet_general.Clone()
        'sheet_noact.TableName = "NoActivity"

        dsWorkbook.Tables.Add(sheet_general)
		dsWorkbook.Tables.Add(sheet_vmi)
		dsWorkbook.Tables.Add(sheet_moulton)
		dsWorkbook.Tables.Add(sheet_rma)
        dsWorkbook.Tables.Add(sheet_vra)
        'dsWorkbook.Tables.Add(sheet_noact)
#End Region

        Console.Title = "SendInventory"

		Console.WriteLine("starting")

        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        'ns.CookieContainer = New Net.CookieContainer
        ns.applicationInfo = New ApplicationInfo With {.applicationId = "823D3679-9AE1-4C6E-A315-6C9112A4614D"}
        ns.Timeout = 120000
        ns.preferences = New Preferences With {.runServerSuiteScriptAndTriggerWorkflows = False, .runServerSuiteScriptAndTriggerWorkflowsSpecified = True}

        Console.WriteLine("login")

        ns.passport = New Passport With {.account = "1027615", .email = "kng@c2wireless.com", .password = "C2Wireless"}
        'ns.passport = New Passport With {.account = "1027615", .email = "lphan@c2wireless.com", .password = "MinhTrang10"}
        ns.login(ns.passport)

        Console.WriteLine("building lookups")

        MakeLookups()

        Dim sellprices As Dictionary(Of String, Double) = GetSellPrices()
        'Dim sellprices As New Dictionary(Of String, Double)

        Dim search As New ItemSearchAdvanced
        search.savedSearchScriptId = "customsearch_dailyinv"

        Console.Write("running item search ... ")

        Dim sr As SearchResult = ns.search(search)

		Console.WriteLine(sr.totalRecords & " results ... " & sr.totalPages & " pages")

        Dim excludeids As New List(Of String)

#Region " fill tables "
        For page As Integer = 1 To sr.totalPages
			For Each row As ItemSearchRow In sr.searchRowList
				Dim location As String = Nothing

				Try
					location = row.inventoryLocationJoin.nameNoHierarchy(0).searchValue
				Catch ex As Exception

				End Try

				If location IsNot Nothing Then
					Dim sheet As DataTable = Nothing

					If location.ToLower.Contains("rma") Then
						sheet = sheet_rma
					ElseIf location.ToLower.Contains("vmi") Then
						sheet = sheet_vmi
					ElseIf location.ToLower.Contains("vendor") Then
						sheet = sheet_vra
					ElseIf location.ToLower.Contains("nuys") Or location.ToLower.Contains("dayton") Then
						sheet = sheet_moulton
					Else
						sheet = sheet_general
					End If

					If sheet IsNot Nothing Then
						Dim l As DataRow = sheet.NewRow

#Region " safely get all the crap "
                        Try
                            excludeids.Add(row.basic.internalId(0).searchValue.internalId)
                        Catch ex As Exception

                        End Try
                        Try
							l("Available") = row.basic.locationQuantityAvailable(0).searchValue
						Catch ex As Exception
							l("Available") = 0
						End Try
                        Try
                            l("Avg Cost") = row.basic.locationAverageCost(0).searchValue
                        Catch ex As Exception
                            l("Avg Cost") = 0.00
                        End Try
                        Try
                            Dim key As String = row.basic.inventoryLocation(0).searchValue.internalId & "-" & row.basic.internalId(0).searchValue.internalId

                            If sellprices.ContainsKey(key) Then
                                l("Avg 2wk Selling") = sellprices(key)
                            Else
                                l("Avg 2wk Selling") = 0.00
                            End If
                        Catch ex As Exception
                            l("Avg 2wk Selling") = 0.00
                        End Try
                        Try
							l("Backordered") = row.basic.locationQuantityBackOrdered(0).searchValue
						Catch ex As Exception
							l("Backordered") = 0
						End Try
						Try
							l("C2PN") = row.basic.displayName(0).searchValue
						Catch ex As Exception
							l("C2PN") = ""
						End Try
						Try
							l("Category") = row.parentJoin.itemId(0).searchValue
						Catch ex As Exception
							l("Category") = ""
						End Try
						Try
							l("Class") = LOOKUP_CLASS(GetCustom(row.basic.customFieldList, "custitem_classification"))
						Catch ex As Exception
							l("Class") = ""
						End Try
						Try
							l("Committed") = row.basic.locationQuantityCommitted(0).searchValue
						Catch ex As Exception
							l("Committed") = 0
						End Try
                        'Try
                        '	l("D1") = GetPrice(row.basic.otherPrices, "D1 Pricing")
                        'Catch ex As Exception

                        'End Try
                        Try
							l("Description") = row.basic.salesDescription(0).searchValue
						Catch ex As Exception
							l("Description") = ""
						End Try
                        Try
                            l("LastActivity") = row.basic.lastQuantityAvailableChange(0).searchValue
                        Catch ex As Exception
                            l("LastActivity") = ""
                        End Try
                        Try
                            l("DateCreated") = row.basic.created(0).searchValue
                        Catch ex As Exception
                            l("DateCreated") = ""
                        End Try
                        Try
                            l("Location") = location
                        Catch ex As Exception
                            l("Location") = ""
						End Try
                        'Try
                        '	l("M1") = GetPrice(row.basic.otherPrices, "M1 Pricing")
                        'Catch ex As Exception

                        'End Try
                        Try
							l("Manufacturer") = LOOKUP_MFR(GetCustom(row.basic.customFieldList, "custitem13_2"))
						Catch ex As Exception
							l("Manufacturer") = ""
						End Try
						Try
							l("MPN") = row.basic.mpn(0).searchValue
						Catch ex As Exception
							l("MPN") = ""
						End Try
						Try
							l("MSRP") = GetPrice(row.basic.otherPrices, "Online Price")
						Catch ex As Exception

						End Try
						Try
							l("On Hand") = row.basic.locationQuantityOnHand(0).searchValue
						Catch ex As Exception
							l("On Hand") = 0
						End Try
						Try
							l("On Order") = row.basic.locationQuantityOnOrder(0).searchValue
						Catch ex As Exception
							l("On Order") = 0
						End Try
                        'Try
                        '	l("R1") = GetPrice(row.basic.otherPrices, "R1 Pricing")
                        'Catch ex As Exception

                        'End Try
                        'Try
                        '	l("R2") = GetPrice(row.basic.otherPrices, "R2 Pricing")
                        'Catch ex As Exception

                        'End Try
                        Try
                            l("UPC") = row.basic.upcCode(0).searchValue
                        Catch ex As Exception
                            l("UPC") = ""
                        End Try

                        Try
                            l("I1") = GetPrice(row.basic.otherPrices, "1- I1/International Distributor")
                        Catch ex As Exception

                        End Try
                        Try
                            l("C1") = GetPrice(row.basic.otherPrices, "2- C1/Carrier/Master Distributor")
                        Catch ex As Exception

                        End Try
                        Try
                            l("D1") = GetPrice(row.basic.otherPrices, "3- D1/Sub-Distribution")
                        Catch ex As Exception

                        End Try
                        Try
                            l("R2") = GetPrice(row.basic.otherPrices, "4- R2/Large Volume Retailer")
                        Catch ex As Exception

                        End Try
                        Try
                            l("R1") = GetPrice(row.basic.otherPrices, "5- R1/Small Volume Retailer")
                        Catch ex As Exception

                        End Try
                        Try
                            l("B1") = GetPrice(row.basic.otherPrices, "6- B1/B2B")
                        Catch ex As Exception

                        End Try
                        Try
                            l("M1") = GetPrice(row.basic.otherPrices, "7- M1/Ecommerce")
                        Catch ex As Exception

                        End Try
#End Region

                        sheet.Rows.Add(l)
					End If
				End If
			Next

			If page < sr.totalPages Then
				Console.WriteLine("getting page " & page + 1 & " ...")

                sr = ns.searchMoreWithId(sr.searchId, page + 1)
            End If
		Next
#End Region

        GetNoActSkus(ns, sheet_general, excludeids)

        Dim filename As String = "inventory-" & Now.Year & "." & Now.Month.ToString.PadLeft(2, "0") & "." & Now.Day.ToString.PadLeft(2, "0") & ".xlsx"

		Console.WriteLine("building " & filename)

		Dim ms As New IO.MemoryStream()

		Using excelPackage As New ExcelPackage(ms)
            For Each tab As DataTable In dsWorkbook.Tables
                If tab IsNot Nothing AndAlso tab.Rows IsNot Nothing AndAlso tab.Rows.Count > 0 Then
                    Dim ws As ExcelWorksheet = excelPackage.Workbook.Worksheets.Add(tab.TableName)
                    Dim tabtable As Table.ExcelTable = ws.Tables.Add(ws.Cells("A1").LoadFromDataTable(tab, True), tab.TableName)
                    tabtable.ShowHeader = True
                    tabtable.TableStyle = Table.TableStyles.Medium2
                    ws.Cells("L:U").Style.Numberformat.Format = "0.00"
                    ws.Cells("W:W").Style.Numberformat.Format = "m/d/yyyy"
                    ws.Cells("Y:Y").Style.Numberformat.Format = "m/d/yyyy"

                    ws.Comments.Add(ws.Cells("N1"), "1- I1/International Distributor", "ckoch")
                    ws.Comments.Add(ws.Cells("O1"), "2- C1/Carrier/Master Distributor", "ckoch")
                    ws.Comments.Add(ws.Cells("P1"), "3- D1/Sub-Distribution", "ckoch")
                    ws.Comments.Add(ws.Cells("Q1"), "4- R2/Large Volume Retailer", "ckoch")
                    ws.Comments.Add(ws.Cells("R1"), "5- R1/Small Volume Retailer", "ckoch")
                    ws.Comments.Add(ws.Cells("S1"), "6- B1/B2B", "ckoch")
                    ws.Comments.Add(ws.Cells("T1"), "7- M1/Ecommerce", "ckoch")

                    ws.Cells.AutoFitColumns()
                End If
            Next

            excelPackage.Save()
		End Using

        Using file As New IO.FileStream(filename, IO.FileMode.Create)
            ms.WriteTo(file)
        End Using

        Console.WriteLine("send email")

        Dim smtp As New Net.Mail.SmtpClient("smtp.office365.com")
		smtp.Port = 587
		smtp.EnableSsl = True
        smtp.Credentials = New Net.NetworkCredential("kng@c2wireless.com", "Office365!")

        Dim email As New Net.Mail.MailMessage
		email.From = New Net.Mail.MailAddress(My.Settings.MailFromAddress, My.Settings.MailFromName)
        'email.To.Add("kng@C2wireless.com")
        For Each t As String In My.Settings.MailTo.Split(",")
            If t.Contains("@") Then email.To.Add(New Net.Mail.MailAddress(t))
        Next
        For Each t As String In My.Settings.MailCC.Split(",")
            If t.Contains("@") Then email.CC.Add(New Net.Mail.MailAddress(t))
        Next
        For Each t As String In My.Settings.MailBCC.Split(",")
            If t.Contains("@") Then email.Bcc.Add(New Net.Mail.MailAddress(t))
        Next
        email.Subject = My.Settings.MailSubject
		email.Body = My.Settings.MailBody
		Dim attachment As New Net.Mail.Attachment(filename)
		email.Attachments.Add(attachment)

        smtp.Send(email)

        Console.WriteLine("sent")

		Console.WriteLine("deleting")

		Try
			attachment.Dispose()
			IO.File.Delete(filename)
		Catch ex As Exception

		End Try

		Console.WriteLine("kbye")
	End Sub

    Private Sub GetNoActSkus(ByRef ns As NetSuiteService, ByRef sheet As DataTable, ByRef exclude As List(Of String))
        Console.Write("getting more skus for byron ... ")

        Dim search As New ItemSearchAdvanced
        search.savedSearchScriptId = "customsearch_dailyinv_2"

        Dim sr As SearchResult = ns.search(search)

        Console.WriteLine(sr.totalRecords & " results ... " & sr.totalPages & " pages")

#Region " fill tables "
        For page As Integer = 1 To SR.totalPages
            For Each row As ItemSearchRow In sr.searchRowList
                If exclude.Contains(row.basic.internalId(0).searchValue.internalId) Then
                    Continue For
                End If

                Dim l As DataRow = sheet.NewRow

                'l("Available") = ""
                'l("Avg Cost") = ""
                'l("Avg 2wk Selling") = ""
                'l("Backordered") = ""
                Try
                    l("C2PN") = row.basic.displayName(0).searchValue
                Catch ex As Exception
                    l("C2PN") = ""
                End Try
                Try
                    l("Category") = row.parentJoin.itemId(0).searchValue
                Catch ex As Exception
                    l("Category") = ""
                End Try
                Try
                    l("Class") = LOOKUP_CLASS(GetCustom(row.basic.customFieldList, "custitem_classification"))
                Catch ex As Exception
                    l("Class") = ""
                End Try
                'l("Committed") = ""
                'Try
                '    l("D1") = GetPrice(row.basic.otherPrices, "D1 Pricing")
                'Catch ex As Exception

                'End Try
                Try
                    l("Description") = row.basic.salesDescription(0).searchValue
                Catch ex As Exception
                    l("Description") = ""
                End Try
                Try
                    l("LastActivity") = row.basic.lastQuantityAvailableChange(0).searchValue
                Catch ex As Exception
                    l("LastActivity") = ""
                End Try
                Try
                    l("DateCreated") = row.basic.created(0).searchValue
                Catch ex As Exception
                    l("DateCreated") = ""
                End Try
                'l("Location") = ""
                'Try
                '    l("M1") = GetPrice(row.basic.otherPrices, "M1 Pricing")
                'Catch ex As Exception

                'End Try
                Try
                    l("Manufacturer") = LOOKUP_MFR(GetCustom(row.basic.customFieldList, "custitem13_2"))
                Catch ex As Exception
                    l("Manufacturer") = ""
                End Try
                Try
                    l("MPN") = row.basic.mpn(0).searchValue
                Catch ex As Exception
                    l("MPN") = ""
                End Try
                Try
                    l("MSRP") = GetPrice(row.basic.otherPrices, "Online Price")
                Catch ex As Exception

                End Try
                'l("On Hand") = ""
                'l("On Order") = ""
                'Try
                '    l("R1") = GetPrice(row.basic.otherPrices, "R1 Pricing")
                'Catch ex As Exception

                'End Try
                'Try
                '    l("R2") = GetPrice(row.basic.otherPrices, "R2 Pricing")
                'Catch ex As Exception

                'End Try
                Try
                    l("UPC") = row.basic.upcCode(0).searchValue
                Catch ex As Exception

                End Try

                Try
                    l("I1") = GetPrice(row.basic.otherPrices, "1- I1/International Distributor")
                Catch ex As Exception

                End Try
                Try
                    l("C1") = GetPrice(row.basic.otherPrices, "2- C1/Carrier/Master Distributor")
                Catch ex As Exception

                End Try
                Try
                    l("D1") = GetPrice(row.basic.otherPrices, "3- D1/Sub-Distribution")
                Catch ex As Exception

                End Try
                Try
                    l("R2") = GetPrice(row.basic.otherPrices, "4- R2/Large Volume Retailer")
                Catch ex As Exception

                End Try
                Try
                    l("R1") = GetPrice(row.basic.otherPrices, "5- R1/Small Volume Retailer")
                Catch ex As Exception

                End Try
                Try
                    l("B1") = GetPrice(row.basic.otherPrices, "6- B1/B2B")
                Catch ex As Exception

                End Try
                Try
                    l("M1") = GetPrice(row.basic.otherPrices, "7- M1/Ecommerce")
                Catch ex As Exception

                End Try
#End Region

                sheet.Rows.Add(l)
            Next

            If page < SR.totalPages Then
                Console.WriteLine("getting page " & page + 1 & " ...")

                SR = ns.searchMoreWithId(SR.searchId, page + 1)
            End If
        Next
    End Sub

    Private Function GetCustom(fields As SearchColumnCustomField(), scriptid As String)
		Dim output As String = ""

		For Each f As SearchColumnCustomField In fields
			If f.scriptId = scriptid Then
				output = CType(f, SearchColumnSelectCustomField).searchValue.internalId
			End If
		Next

		Return output
	End Function

	Private Function GetPrice(fields As SearchColumnDoubleField(), label As String)
		Dim output As Double = 0.00

		For Each f As SearchColumnDoubleField In fields
            If f.customLabel = label And f.searchValueSpecified = True Then
                output = f.searchValue
            End If
        Next

		Return output
	End Function

    Function GetSellPrices() As Dictionary(Of String, Double)
        Dim output As New Dictionary(Of String, Double)
        Dim prices As New Dictionary(Of String, List(Of Double))

        Dim search As New TransactionSearchAdvanced
        search.savedSearchId = "1732"

        Console.Write("getting avg selling prices ... ")

        Dim sr As SearchResult = ns.search(search)

        Console.WriteLine(sr.totalRecords & " results ... " & sr.totalPages & " pages")

        For page As Integer = 1 To sr.totalPages
            For Each row As TransactionSearchRow In sr.searchRowList
                Dim locid As String = row.basic.location(0).searchValue.internalId
                Dim itemid As String = row.basic.item(0).searchValue.internalId
                Dim key As String = locid & "-" & itemid
                Dim price As Double = row.basic.rate(0).searchValue

                If prices.ContainsKey(key) = False Then
                    prices.Add(key, New List(Of Double))
                End If

                prices(key).Add(price)
            Next

            If page < sr.totalPages Then
                Console.WriteLine("getting page " & page + 1)

                sr = ns.searchMoreWithId(sr.searchId, page + 1)
            End If
        Next

        Console.WriteLine("averaging prices ...")

        For Each item As KeyValuePair(Of String, List(Of Double)) In prices
            Dim total As Double = 0
            For Each p As Double In item.Value
                total += p
            Next

            output.Add(item.Key, total / item.Value.Count)
        Next

        Console.WriteLine("done")

        Return output
    End Function

End Module
