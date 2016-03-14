Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports FacadeFor3e

Module Program
    Private ReadOnly Rnd As Random = New Random()
    Private ReadOnly Today As DateTime = DateTime.Today

    Sub Main()
        Call DisplaySettingsInfo()
        Dim listOfIcbUnits As HashSet(Of String) = GetListOfIcbUnits()
        Dim mattersByUnits As Dictionary(Of String, HashSet(Of Integer)) = GetActiveDataAndTheirUnits(AddressOf BuildSqlCommandForRetrievingMatters)
        Dim timekeepersByUnits As Dictionary(Of String, HashSet(Of Integer)) = GetActiveDataAndTheirUnits(AddressOf BuildSqlCommandForRetrievingTimekeepers)

        Try
            Call RunTestingProcess(listOfIcbUnits, mattersByUnits, timekeepersByUnits)
        Catch ex As Exception
            Call Console.WriteLine(ex.Message)
        Finally
            If IsRunningInVisualStudio() Then
                Call Console.WriteLine()
                Call Console.WriteLine("Press Return to continue.")
                Call Console.ReadLine()
            End If
        End Try
    End Sub

    Private Sub RunTestingProcess(listOfIcbUnits As HashSet(Of String), mattersByUnits As Dictionary(Of String, HashSet(Of Integer)), timekeepersByUnits As Dictionary(Of String, HashSet(Of Integer)))
        Dim matterTimekeepers As New List(Of MattTkpr)
        For Each unit As String In listOfIcbUnits
            Dim matter As Integer
            If Not TrySelectItem(mattersByUnits, unit, matter) Then
                Call Console.WriteLine("No active matters found for unit " & unit)
                Continue For
            End If

            Dim matterUnit As String = unit
            For Each differentUnit As String In listOfIcbUnits.Where(Function(item) item <> matterUnit)
                Dim timekeeper As Integer
                If Not TrySelectItem(timekeepersByUnits, differentUnit, timekeeper) Then
                    Call Console.WriteLine("No active timekeepers found for unit " & differentUnit)
                    Continue For
                End If

                Call Console.WriteLine("Processing for matter in unit " & unit & ", timekeeper in unit " & differentUnit)
                Dim mt As New MattTkpr With {.Matter = matter, .Timekeeper = timekeeper}
                Call CreateAndBillCards(mt)
                Call matterTimekeepers.Add(mt)
            Next
        Next
        Dim proforma As Integer = CreateProforma(matterTimekeepers)
    End Sub

    Private Sub DisplaySettingsInfo()
        Call Console.WriteLine("Test InterCompany Billing")
        Call Console.WriteLine()

        ' todo: output settings
    End Sub

    Private Function GetListOfIcbUnits() As HashSet(Of String)
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Default").ConnectionString)
            Call conn.Open()

            Using command As SqlCommand = BuildSqlCommandForRetrievingIcbUnits(conn)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("There are no GL units set up for intercompany billing.")
                    End If

                    Dim result As New HashSet(Of String)
                    Do
                        Dim unit As String = reader.GetString(0)
                        Call result.Add(unit)
                    Loop While reader.Read()

                    Return result
                End Using
            End Using
        End Using
    End Function

    Private Function BuildSqlCommandForRetrievingIcbUnits(connection As SqlConnection) As SqlCommand
        Dim sql As String =
            "SELECT     NxUnit" & vbCrLf &
            "FROM       GLUnit" & vbCrLf &
            "WHERE      IntercoNat Is Not null" & vbCrLf &
            "AND		IntercoOffice is not null" & vbCrLf &
            "AND		ICBIntercoWIP is not null" & vbCrLf &
            "AND		ICBIntercoAR is not null" & vbCrLf &
            "AND		ICBIntercoARTo is not null"
        Dim result As New SqlCommand(sql, connection)
        Return result
    End Function

    Private Function GetActiveDataAndTheirUnits(dataRetrievalFunction As Func(Of SqlConnection, SqlCommand)) As Dictionary(Of String, HashSet(Of Integer))
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Default").ConnectionString)
            Call conn.Open()

            Using command As SqlCommand = dataRetrievalFunction(conn)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("No active records found.")
                    End If

                    Dim result As New Dictionary(Of String, HashSet(Of Integer))
                    Do
                        Dim matter As Integer = reader.GetInt32(0)
                        Dim unit As String = reader.GetString(1)
                        Dim items As HashSet(Of Integer) = Nothing
                        If Not result.TryGetValue(unit, items) Then
                            items = New HashSet(Of Integer)
                            Call result.Add(unit, items)
                        End If
                        Call items.Add(matter)
                    Loop While reader.Read()

                    Return result
                End Using
            End Using
        End Using
    End Function

    Private Function BuildSqlCommandForRetrievingMatters(connection As SqlConnection) As SqlCommand
        Dim sql As String =
            "SELECT	    MattIndex, NxUnit" & vbCrLf &
            "FROM	    Matter" & vbCrLf &
            "			    INNER JOIN MattDate ON Matter.MattIndex = MattDate.MatterLkUp AND Convert(date, GetDate()) BETWEEN NxStartDate AND NxEndDate" & vbCrLf &
            "			    INNER JOIN Office ON MattDate.Office = Office.Code" & vbCrLf &
            "			    INNER JOIN MattStatus ON Matter.MattStatus = MattStatus.Code" & vbCrLf &
            "WHERE	    MattStatus.IsTimeEntry = 1" & vbCrLf &
            "AND		MattStatus.IsCostEntry = 1" & vbCrLf &
            "AND		MattStatus.IsBilling = 1" & vbCrLf &
            "AND		MattStatus.IsPayment = 1" & vbCrLf &
            "AND        MattDate.PTAGroup is null" & vbCrLf &
            "AND        MattDate.PTAGroupCost is null"
        Dim result As New SqlCommand(sql, connection)
        Return result
    End Function

    Private Function BuildSqlCommandForRetrievingTimekeepers(connection As SqlConnection) As SqlCommand
        Dim sql As String =
            "SELECT	TkprIndex, Office.NxUnit" & vbCrLf &
            "FROM	Timekeeper" & vbCrLf &
            "			INNER JOIN TkprDate ON Timekeeper.TkprIndex = TkprDate.TimekeeperLkUp AND Convert(date, GetDate()) BETWEEN NxStartDate AND NxEndDate" & vbCrLf &
            "			INNER JOIN Office ON TkprDate.Office = Office.Code" & vbCrLf &
            "			INNER JOIN TkprStatus ON Timekeeper.TkprStatus = TkprStatus.Code" & vbCrLf &
            "WHERE	TkprStatus.IsAllowTime = 1" & vbCrLf &
            "AND		TkprStatus.IsAllowCost = 1"
        Dim result As New SqlCommand(sql, connection)
        Return result
    End Function

    Private Function TrySelectItem(itemsByUnits As Dictionary(Of String, HashSet(Of Integer)), unit As String, ByRef result As Integer) As Boolean
        Dim hashset As HashSet(Of Integer) = Nothing
        If Not itemsByUnits.TryGetValue(unit, hashset) Then
            result = 0
            Return False
        End If

        Dim countOfItems As Integer = hashset.Count
        result = hashset.ElementAt(Rnd.Next(countOfItems))
        Return True
    End Function

    Private Sub CreateAndBillCards(mt As MattTkpr)
        Call CreateTimecard(mt)
        Call CreateCostcard(mt)
    End Sub

    Private Function CreateTimecard(mt As MattTkpr) As Integer
        Dim p As Process = Process.NewEsbTimeCardLoadProcess()
        Dim a As OperationAdd = p.AddOperation()
        Call a.AddAttribute("WorkDate", Today)
        Call a.AddAttribute("Matter", mt.Matter)
        Call a.AddAttribute("Timekeeper", mt.Timekeeper)
        Call a.AddAttribute("WorkHrs", 0.1D)    ' minimal amount to avoid posting too much time in a single day for the timekeeper
        Call a.AddAttribute("Narrative", "ICB Test")

        Dim rpp As New RunProcessParameters(p)
        rpp.ThrowExceptionIfProcessDoesNotComplete = True
        rpp.GetKey = False
        Dim rpr As RunProcessResult = RunProcess.ExecuteProcess(rpp)
        Dim result As Integer = GetNewCardIndexFromProcessId(AddressOf BuildSqlCommandForRetrievingTimeIndex, rpr.ProcessId)
        Return result
    End Function

    Private Function CreateCostcard(mt As MattTkpr) As Integer
        Dim p As Process = Process.NewEsbCostCardLoadProcess()
        Dim a As OperationAdd = p.AddOperation()
        Call a.AddAttribute("WorkDate", Today)
        Call a.AddAttribute("Matter", mt.Matter)
        Call a.AddAttribute("Timekeeper", mt.Timekeeper)
        Call a.AddAttribute("CostType", "WALLET")
        Call a.AddAttribute("WorkQty", Rnd.Next(10) + 1)
        Call a.AddAttribute("Currency", "GBP")
        Call a.AddAttribute("WorkRate", 10)
        Call a.AddAttribute("Narrative", "ICB Test")

        Dim rpp As New RunProcessParameters(p)
        rpp.ThrowExceptionIfProcessDoesNotComplete = True
        rpp.GetKey = False
        Dim rpr As RunProcessResult = RunProcess.ExecuteProcess(rpp)
        Dim result As Integer = GetNewCardIndexFromProcessId(AddressOf BuildSqlCommandForRetrievingCostIndex, rpr.ProcessId)
        Return result
    End Function

    Private Function GetNewCardIndexFromProcessId(dataRetrievalFunction As Func(Of SqlConnection, Guid, SqlCommand), processId As Guid) As Integer
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Default").ConnectionString)
            Call conn.Open()

            Using command As SqlCommand = dataRetrievalFunction(conn, processId)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("Timecard could not be identified.")
                    End If

                    Dim result As Integer = reader.GetInt32(0)
                    Return result
                End Using
            End Using
        End Using
    End Function

    Private Function BuildSqlCommandForRetrievingTimeIndex(connection As SqlConnection, processId As Guid) As SqlCommand
        Dim sql As String =
            "SELECT		TOP 1 TimeIndex" & vbCrLf &
            "FROM		Timecard" & vbCrLf &
            "WHERE		OrigProcItemID = @ProcessItemId" & vbCrLf &
            "ORDER BY   TimeIndex DESC"
        Dim result As New SqlCommand(sql, connection)
        result.CommandTimeout = 0
        Dim parameter As SqlParameter = result.Parameters.Add("@ProcessItemId", SqlDbType.UniqueIdentifier)
        parameter.Value = processId
        Return result
    End Function

    Private Function BuildSqlCommandForRetrievingCostIndex(connection As SqlConnection, processId As Guid) As SqlCommand
        Dim sql As String =
            "SELECT		TOP 1 CostIndex" & vbCrLf &
            "FROM		Costcard" & vbCrLf &
            "WHERE		OrigProcItemID = @ProcessItemId" & vbCrLf &
            "ORDER BY   CostIndex DESC"
        Dim result As New SqlCommand(sql, connection)
        result.CommandTimeout = 0
        Dim parameter As SqlParameter = result.Parameters.Add("@ProcessItemId", SqlDbType.UniqueIdentifier)
        parameter.Value = processId
        Return result
    End Function

    Private Function CreateProforma(criteria As IEnumerable(Of MattTkpr)) As Integer
        Dim p As New Process("CCC_ProfGen_srv", "ProfGenerationRun")
        Dim a As OperationAdd = p.AddOperation()
        Call a.AddAttribute("Description", "ICB Testing")
        Call a.AddAttribute("ProfStatus", "Edit")
        Call a.AddAttribute("IsCreateSingleProforma", True)
        Call a.AddAttribute("IsIgnoreExcludedEntry", False)
        Call a.AddAttribute("IsIncludeOtherProforma", False)
        Call a.AddAttribute("ProformaDateSelectList", "WorkDate")
        Call a.AddAttribute("TimeStart", Today)
        Call a.AddAttribute("TimeEnd", Today)
        Call a.AddAttribute("CostStart", Today)
        Call a.AddAttribute("CostEnd", Today)
        Call a.AddAttribute("ChrgStart", Today)
        Call a.AddAttribute("ChrgEnd", Today)

        'Call a.AddAttribute("TemplateName", "TE_EDG_Proforma")
        'Call a.AddAttribute("TemplateFormat", "PDF")
        'Call a.AddAttribute("PrinterXml", "<PrintInfo PrintType=""Template"" Name=""Print Proforma"" PrintToScreen=""False"" SaveLocal=""False"" Condense=""False"" TemplateName=""TE_EDG_Proforma"" TemplateFormat=""PDF"" AllowReprint=""False""><AdvancedOptions PaperSize=""A4"" FitMethod=""Standard"" Orientation=""Default"" GridFormat=""6"" KeepGroups=""False"" IntelliBreak=""False"" PrintCoverPage=""False"" PrintHeaderAtTheTop=""False"" PrintGrandTotalOnNewPage=""False"" SuppressIfSame=""False"" /><Email From=""jon.saffron@ashurst.com"" ToList=""jon.saffron@ashurst.com"" /></PrintInfo>")

        Dim profGenerationChild As DataObject = a.AddChild("ProfGeneration")
        For Each mt As MattTkpr In criteria
            Dim c As OperationAdd = profGenerationChild.AddOperation()
            Call c.AddAttribute("BillingTkpr", mt.Timekeeper)
            Call c.AddAttribute("Matter", mt.Matter)
        Next

        Dim rpp As New RunProcessParameters(p)
        rpp.ThrowExceptionIfProcessDoesNotComplete = True
        rpp.GetKey = True
        Dim rpr As RunProcessResult = RunProcess.ExecuteProcess(rpp)
        Dim profGenerationRun As Guid = Guid.Parse(rpr.NewKey)
        Dim result As Integer = GetNewProformaIndexFromProfGenerationRun(profGenerationRun)
        Return result
    End Function

    Private Function GetNewProformaIndexFromProfGenerationRun(profGenerationRun As Guid) As Integer
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Default").ConnectionString)
            Call conn.Open()

            Using command As SqlCommand = BuildSqlCommandForRetrievingProformaIndex(conn, profGenerationRun)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("Proforma could not be identified.")
                    End If

                    Dim result As Integer = reader.GetInt32(0)
                    Return result
                End Using
            End Using
        End Using
    End Function

    Private Function BuildSqlCommandForRetrievingProformaIndex(connection As SqlConnection, profGenerationRun As Guid) As SqlCommand
        Dim sql As String =
            "SELECT		TOP 1 ProfIndex" & vbCrLf &
            "FROM		ProfMaster" & vbCrLf &
            "WHERE		ProfGenerationRun = @ProfGenerationRun" & vbCrLf &
            "ORDER BY   ProfIndex DESC"
        Dim result As New SqlCommand(sql, connection)
        result.CommandTimeout = 0
        Dim parameter As SqlParameter = result.Parameters.Add("@ProfGenerationRun", SqlDbType.UniqueIdentifier)
        parameter.Value = profGenerationRun
        Return result
    End Function

    Private Function IsRunningInVisualStudio() As Boolean
        Dim mainModule = Diagnostics.Process.GetCurrentProcess().MainModule
        If mainModule.ModuleName.EndsWith(".vshost.exe", StringComparison.OrdinalIgnoreCase) Then
            Dim basePath As String = Path.GetDirectoryName(mainModule.FileName)
            If basePath.EndsWith("\bin\debug", StringComparison.OrdinalIgnoreCase) OrElse basePath.EndsWith("\bin\release", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Class MattTkpr
        Public Matter As Integer
        Public Timekeeper As Integer
    End Class
End Module
