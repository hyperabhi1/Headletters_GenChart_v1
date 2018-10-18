Imports System.Data.SqlClient
Public Class PromiseRAKE
    Dim HID = ""
    Dim UID = ""
    'Public constr = "data source=DESKTOP-JBRFH9E;initial catalog=HEADLETTERS_ENGINE;integrated security=True;"
    'Public constr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS_ENGINE;integrated security=True;multipleactiveresultsets=True;"
    'Public constr = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Public constr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS_ENGINE;integrated security=False;multipleactiveresultsets=True;User Id=sa;password=pSI)TA1t0K[);"
    'Sub Main()
    '    Dim connection As SqlConnection = New SqlConnection(constr)
    '    Try
    '        Dim cmd As New SqlCommand($"SELECT HM.HUSERID, HM.HID FROM ASTROLOGYSOFTWARE_DB.DBO.HMAIN AS HM JOIN HEADLETTERS_ENGINE.DBO.REPORT AS R ON HM.HUSERID = R.UID AND HM.HID = R.HID WHERE HM.HSTATUS = '5' AND R.CHART_WORK = 1 AND R.PROMISE_RAKE_CALC IS NULL;", connection)
    '        Dim da As New SqlDataAdapter(cmd)
    '        Dim RowsData As New DataSet()
    '        da.Fill(RowsData)
    '        If RowsData.Tables(0).Rows.Count > 0 Then
    '        End If
    '        For i As Integer = 0 To RowsData.Tables(0).Rows.Count - 1
    '            Try
    '                UID = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
    '                HID = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
    '                MainHIDUID(UID, HID)
    '            Catch ex As Exception
    '                Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
    '                IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
    '                Continue For
    '            End Try
    '        Next
    '    Catch ex As Exception
    '        Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
    '        IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
    '    Finally
    '        connection.Close()
    '    End Try
    'End Sub
    Sub MainHIDUID(ByVal UID As String, ByVal HID As String)
        Dim HCUSPList As New List(Of HCUSP)
        Dim HRAKEList As New List(Of HRAKE)
        Dim HPLANETList As New List(Of HPLANET)
        Dim HCUSPCALCList As New List(Of HCUSP_CALC)
        DELETE_RECORDS(UID, HID)
        Select_From_HRAKE(UID, HID, HRAKEList)
        Select_From_HCUSP(UID, HID, HCUSPList)
        Select_From_HPLANET(UID, HID, HPLANETList)
        Conversion_of_HCUSP_RAKE(UID, HID, HCUSPList, HRAKEList)
        Conversion_of_HPLANET_RAKE(UID, HID, HPLANETList, HRAKEList)
        Insert_Into_HCUSP_CALC(UID, HID, HCUSPList, HPLANETList)
        Insert_Into_HPROMISE_CALC(UID, HID)
        UPDATE_REPORT(UID, HID)
        Dim promise_ As New Promise
        promise_.MainPromiseUIDHID(UID, HID)
    End Sub
    Sub DELETE_RECORDS(ByVal UID As String, ByVal HID As String)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = constr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = $"DELETE FROM HEADLETTERS_ENGINE.DBO.HCUSP_CALC WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "';
                                DELETE FROM HEADLETTERS_ENGINE.DBO.HPROMISE_CALC WHERE HPCUSPID = '" + UID + "' AND HPHID = '" + HID + "';"
            cmd.ExecuteNonQuery()
        Catch EX As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, EX.Message + vbCrLf + EX.StackTrace))
            con.Close()
        End Try
    End Sub
    Sub UPDATE_REPORT(ByVal UID As String, ByVal HID As String)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = constr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = $"UPDATE HEADLETTERS_ENGINE.DBO.REPORT SET PROMISE_RAKE_CALC = 1, HCUSP_CALC = 1, HPROMISE_CALC = 1 WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            cmd.ExecuteNonQuery()
        Catch EX As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, EX.Message + vbCrLf + EX.StackTrace))
            con.Close()
        End Try
    End Sub
    Function SplitInTwoChar(ByRef S1 As String) As String()
        Dim S1Split = New String(S1.Length / 2 + (If(S1.Length Mod 2 = 0, 0, 1)) - 1) {}
        For i As Integer = 0 To S1Split.Length - 1
            S1Split(i) = S1.Substring(i * 2, If(i * 2 + 2 > S1.Length, 1, 2))
        Next
        Return S1Split
    End Function
    Sub Select_From_HRAKE(ByVal UID As String, ByVal HID As String, ByRef HRAKEList As List(Of HRAKE))
        Dim connection As SqlConnection = New SqlConnection(constr)
        Dim query = $"SELECT * FROM HEADLETTERS_ENGINE.DBO.HRAKE WHERE UID = '" + UID + "' AND HID = '" + HID + "' AND HKEY = 'RA';
                                SELECT * FROM HEADLETTERS_ENGINE.DBO.HRAKE WHERE UID = '" + UID + "' AND HID = '" + HID + "' AND HKEY = 'KE';"
        Dim cmd As New SqlCommand(query, connection)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        Try
            connection.Open()
            da.Fill(ds)
            If ds.Tables(0).Rows.Count <> 0 Then
                Dim hRA = New HRAKE With {
                    .UID = ds.Tables(0).Rows(0).Item(0).ToString().Trim,
                    .HID = ds.Tables(0).Rows(0).Item(1).ToString().Trim,
                    .HKEY = ds.Tables(0).Rows(0).Item(2).ToString().Trim,
                    .S1 = ds.Tables(0).Rows(0).Item(3).ToString().Trim
                }
                HRAKEList.Add(hRA)
            End If
            If ds.Tables(1).Rows.Count <> 0 Then
                Dim hKE = New HRAKE With {
                    .UID = ds.Tables(1).Rows(0).Item(0).ToString().Trim,
                    .HID = ds.Tables(1).Rows(0).Item(1).ToString().Trim,
                    .HKEY = ds.Tables(1).Rows(0).Item(2).ToString().Trim,
                    .S1 = ds.Tables(1).Rows(0).Item(3).ToString().Trim
                }
                HRAKEList.Add(hKE)
            End If
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            connection.Close()
        End Try
    End Sub
    Sub Select_From_HCUSP(ByVal UID As String, ByVal HID As String, ByRef HCUSPList As List(Of HCUSP))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = constr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "';"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim hCusp = New HCUSP With {
                    .CUSPUSERID = reader("CUSPUSERID").ToString().Trim,
                    .CUSPHID = reader("CUSPHID").ToString().Trim,
                    .CUSP = reader("CUSP").ToString().Trim,
                    .SIGN = reader("SIGN").ToString().Trim,
                    .DMS = reader("DMS").ToString().Trim,
                    .CP1 = reader("CP1").ToString().ToUpper().Trim,
                    .CP2 = reader("CP2").ToString().ToUpper().Trim,
                    .CP3 = reader("CP3").ToString().ToUpper().Trim
                }
                HCUSPList.Add(hCusp)
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_HPLANET(ByVal UID As String, ByVal HID As String, ByRef HPLANETList As List(Of HPLANET))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = constr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HPLANET WHERE PLHUSERID = '" + UID + "' AND PLHID = '" + HID + "';"
            reader = cmd.ExecuteReader()
            While reader.Read()
                If reader("PLANET").ToString().Trim <> "NE" And reader("PLANET").ToString().Trim <> "UR" And reader("PLANET").ToString().Trim <> "PL" Then
                    Dim hPlanet = New HPLANET With {
                    .PLHUSERID = reader("PLHUSERID").ToString().Trim,
                    .PLHID = reader("PLHID").ToString().Trim,
                    .PLANET = reader("PLANET").ToString().Trim,
                    .SIGN = reader("SIGN").ToString().Trim,
                    .DMS = reader("DMS").ToString().Trim,
                    .HP1 = reader("HP1").ToString().Trim,
                    .HP2 = reader("HP2").ToString().Trim,
                    .HP3 = reader("HP3").ToString().Trim,
                    .HP4 = reader("HP4").ToString().Trim,
                    .HP5 = reader("HP5").ToString().Trim,
                    .HP6 = reader("HP6").ToString().Trim,
                    .HPHOUSE = reader("HPHOUSE").ToString().Trim
                }
                    HPLANETList.Add(hPlanet)
                End If
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Conversion_of_HCUSP_RAKE(ByVal UID As String, ByVal HID As String, ByVal HCUSPList As List(Of HCUSP), ByVal HRAKEList As List(Of HRAKE))
        Dim HcuspListCount = HCUSPList.Count - 1
        For i As Integer = 0 To HcuspListCount
            If HCUSPList.Item(i).CP1 = "RA" Then
                Dim CP1Expand = "RA"
                Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitRA.Length - 1
                    If S1SplitRA(j) = "KE" Then
                        CP1Expand = CP1Expand + "KE"
                        Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitKE.Length - 1
                            CP1Expand = CP1Expand + S1SplitKE(k)
                        Next
                    Else
                        CP1Expand = CP1Expand + S1SplitRA(j)
                    End If
                Next
                HCUSPList.Item(i).CP1 = CP1Expand
            ElseIf HCUSPList.Item(i).CP1 = "KE" Then
                Dim CP1Expand = "KE"
                Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitKE.Length - 1
                    If S1SplitKE(j) = "RA" Then
                        CP1Expand = CP1Expand + "RA"
                        Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitRA.Length - 1
                            CP1Expand = CP1Expand + S1SplitRA(k)
                        Next
                    Else
                        CP1Expand = CP1Expand + S1SplitKE(j)
                    End If
                Next
                HCUSPList.Item(i).CP1 = CP1Expand
            End If
            If HCUSPList.Item(i).CP2 = "RA" Then
                Dim CP2Expand = "RA"
                Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitRA.Length - 1
                    If S1SplitRA(j) = "KE" Then
                        CP2Expand = CP2Expand + "KE"
                        Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitKE.Length - 1
                            CP2Expand = CP2Expand + S1SplitKE(k)
                        Next
                    Else
                        CP2Expand = CP2Expand + S1SplitRA(j)
                    End If
                Next
                HCUSPList.Item(i).CP2 = CP2Expand
            ElseIf HCUSPList.Item(i).CP2 = "KE" Then
                Dim CP2Expand = "KE"
                Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitKE.Length - 1
                    If S1SplitKE(j) = "RA" Then
                        CP2Expand = CP2Expand + "RA"
                        Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitRA.Length - 1
                            CP2Expand = CP2Expand + S1SplitRA(k)
                        Next
                    Else
                        CP2Expand = CP2Expand + S1SplitKE(j)
                    End If
                Next
                HCUSPList.Item(i).CP2 = CP2Expand
            End If
            If HCUSPList.Item(i).CP3 = "RA" Then
                Dim CP3Expand = "RA"
                Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitRA.Length - 1
                    If S1SplitRA(j) = "KE" Then
                        CP3Expand = CP3Expand + "KE"
                        Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitKE.Length - 1
                            CP3Expand = CP3Expand + S1SplitKE(k)
                        Next
                    Else
                        CP3Expand = CP3Expand + S1SplitRA(j)
                    End If
                Next
                HCUSPList.Item(i).CP3 = CP3Expand
            ElseIf HCUSPList.Item(i).CP3 = "KE" Then
                Dim CP3Expand = "KE"
                Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitKE.Length - 1
                    If S1SplitKE(j) = "RA" Then
                        CP3Expand = CP3Expand + "RA"
                        Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitRA.Length - 1
                            CP3Expand = CP3Expand + S1SplitRA(k)
                        Next
                    Else
                        CP3Expand = CP3Expand + S1SplitKE(j)
                    End If
                Next
                HCUSPList.Item(i).CP3 = CP3Expand
            End If
        Next
    End Sub
    Sub Conversion_of_HPLANET_RAKE(ByVal UID As String, ByVal HID As String, ByVal HPLANETList As List(Of HPLANET), ByVal HRAKEList As List(Of HRAKE))
        Dim HplanetListCount = HPLANETList.Count - 1
        For i As Integer = 0 To HplanetListCount
            If HPLANETList.Item(i).PLANET = "RA" Then
                Dim PlanetExpand = "RA"
                Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitRA.Length - 1
                    If S1SplitRA(j) = "KE" Then
                        PlanetExpand = PlanetExpand + "KE"
                        Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitKE.Length - 1
                            PlanetExpand = PlanetExpand + S1SplitKE(k)
                        Next
                    Else
                        PlanetExpand = PlanetExpand + S1SplitRA(j)
                    End If
                Next
                HPLANETList.Item(i).PLANET = PlanetExpand
            ElseIf HPLANETList.Item(i).PLANET = "KE" Then
                Dim PlanetExpand = "KE"
                Dim S1SplitKE = SplitInTwoChar(HRAKEList.Item(1).S1.ToString().Trim)
                For j As Integer = 0 To S1SplitKE.Length - 1
                    If S1SplitKE(j) = "RA" Then
                        PlanetExpand = PlanetExpand + "RA"
                        Dim S1SplitRA = SplitInTwoChar(HRAKEList.Item(0).S1.ToString().Trim)
                        For k As Integer = 0 To S1SplitRA.Length - 1
                            PlanetExpand = PlanetExpand + S1SplitRA(k)
                        Next
                    Else
                        PlanetExpand = PlanetExpand + S1SplitKE(j)
                    End If
                Next
                HPLANETList.Item(i).PLANET = PlanetExpand
            End If
        Next
    End Sub
    Sub Insert_Into_HCUSP_CALC(ByVal UID As String, ByVal HID As String, ByVal HCUSPList As List(Of HCUSP), ByVal HPLANETList As List(Of HPLANET))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim HCUSPQuery = ""
        Dim HPLANETQuery = ""
        con.ConnectionString = constr
        Try
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM HEADLETTERS_ENGINE.DBO.HCUSP_CALC WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "';"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
        If HCUSPList.Count <> 0 Then
            For i As Integer = 0 To HCUSPList.Count - 1
                Dim HCUSPQueryCP1 = ""
                Dim HCUSPQueryCP2 = ""
                Dim HCUSPQueryCP3 = ""
                If HCUSPList.Item(i).CP1.Length > 2 Then
                    Dim HCUSPCP1 = SplitInTwoChar(HCUSPList.Item(i).CP1)
                    For j As Integer = 0 To HCUSPCP1.Count - 1
                        HCUSPQueryCP1 = HCUSPQueryCP1 + "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPCP1(j) + "','1','" + If(j = 0, "1", "3") + "');" + vbCrLf
                    Next
                Else
                    HCUSPQueryCP1 = "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPList.Item(i).CP1 + "','1','1');" + vbCrLf
                End If
                If HCUSPList.Item(i).CP2.Length > 2 Then
                    Dim HCUSPCP2 = SplitInTwoChar(HCUSPList.Item(i).CP2)
                    For j As Integer = 0 To HCUSPCP2.Count - 1
                        HCUSPQueryCP2 = HCUSPQueryCP2 + "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPCP2(j) + "','2','" + If(j = 0, "1", "3") + "');" + vbCrLf
                    Next
                Else
                    HCUSPQueryCP2 = "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPList.Item(i).CP2 + "','2','1');" + vbCrLf
                End If
                If HCUSPList.Item(i).CP3.Length > 2 Then
                    Dim HCUSPCP3 = SplitInTwoChar(HCUSPList.Item(i).CP3)
                    For j As Integer = 0 To HCUSPCP3.Count - 1
                        HCUSPQueryCP3 = HCUSPQueryCP3 + "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPCP3(j) + "','3','" + If(j = 0, "1", "3") + "');" + vbCrLf
                    Next
                Else
                    HCUSPQueryCP3 = "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HCUSPList.Item(i).CUSPUSERID + "','" + HCUSPList.Item(i).CUSPHID + "','" + HCUSPList.Item(i).CUSP + "','" + HCUSPList.Item(i).CP3 + "','3','1');" + vbCrLf
                End If
                HCUSPQuery = HCUSPQuery + HCUSPQueryCP1 + HCUSPQueryCP2 + HCUSPQueryCP3
            Next
            Try
                con.Open()
                cmd.Connection = con
                cmd.CommandText = HCUSPQuery
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
                IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
            Finally
                con.Close()
            End Try
        End If
        If HPLANETList.Count <> 0 Then
            For i As Integer = 0 To HPLANETList.Count - 1
                Dim MultiPlanetQuery = ""
                If HPLANETList.Item(i).PLANET.Length > 2 Then
                    Dim HPlanetArray = SplitInTwoChar(HPLANETList.Item(i).PLANET)
                    For j As Integer = 0 To HPlanetArray.Count - 1
                        MultiPlanetQuery = MultiPlanetQuery + "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HPLANETList.Item(i).PLHUSERID + "','" + HPLANETList.Item(i).PLHID + "','" + HPLANETList.Item(i).HPHOUSE + "','" + HPlanetArray(j) + "','7','" + If(j = 0, "2", "3") + "');" + vbCrLf
                    Next
                Else
                    MultiPlanetQuery = "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_CALC VALUES ('" + HPLANETList.Item(i).PLHUSERID + "','" + HPLANETList.Item(i).PLHID + "','" + HPLANETList.Item(i).HPHOUSE + "','" + HPLANETList.Item(i).PLANET + "','7','2');" + vbCrLf
                End If
                HPLANETQuery = HPLANETQuery + MultiPlanetQuery
            Next
            Try
                con.Open()
                cmd.Connection = con
                cmd.CommandText = HPLANETQuery
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
                IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
            Finally
                con.Close()
            End Try
        End If
    End Sub
    Sub Insert_Into_HPROMISE_CALC(ByVal UID As String, ByVal HID As String)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = constr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM HEADLETTERS_ENGINE.DBO.HPROMISE_CALC WHERE HPCUSPID = '" + UID + "' AND HPHID = '" + HID + "';"
            cmd.ExecuteNonQuery()
            con.Close()
            con.ConnectionString = constr
            cmd.Connection = con
            For i As Integer = 1 To 12
                Dim HCUSPCALCList As New List(Of HCUSP_CALC)
                cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_CALC WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "' AND CUSPID = '" + i.ToString("D2") + "';"
                con.Open()
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim hCuspCalc = New HCUSP_CALC With {
                    .CUSPUSERID = reader("CUSPUSERID").ToString().Trim,
                    .CUSPHID = reader("CUSPHID").ToString().Trim,
                    .CUSPID = reader("CUSPID").ToString().Trim,
                    .CUSPPLANET = reader("CUSPPLANET").ToString().Trim,
                    .CUSPTYPE = reader("CUSPTYPE").ToString().Trim,
                    .CUSPCAT = reader("CUSPCAT").ToString().ToUpper().Trim
                }
                    HCUSPCALCList.Add(hCuspCalc)
                End While
                con.Close()
                For j As Integer = 0 To HCUSPCALCList.Count - 1
                    cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_CALC WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "' AND CUSPPLANET = '" + HCUSPCALCList.Item(j).CUSPPLANET + "' AND CUSPCAT IN ('1','2');"
                    con.Open()
                    reader = cmd.ExecuteReader()
                    Dim HPROMSIE_CALCQuery = ""
                    While reader.Read()
                        HPROMSIE_CALCQuery = HPROMSIE_CALCQuery + "INSERT INTO HEADLETTERS_ENGINE.DBO.HPROMISE_CALC VALUES ('" + UID + "','" + HID + "','" + i.ToString("D2") + "','" + HCUSPCALCList.Item(j).CUSPPLANET + "','" + reader("CUSPID") + "','" + reader("CUSPPLANET") + "');" + vbCrLf
                    End While
                    con.Close()
                    If HPROMSIE_CALCQuery <> "" Then
                        con.Open()
                        cmd.CommandText = HPROMSIE_CALCQuery
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Next
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\PromiseRAKE_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        End Try
    End Sub
    Public Class HCUSP
        Property CUSPUSERID As String
        Property CUSPHID As String
        Property CUSP As String
        Property SIGN As String
        Property DMS As String
        Property CP1 As String
        Property CP2 As String
        Property CP3 As String
    End Class
    Public Class HRAKE
        Property UID As String
        Property HID As String
        Property HKEY As String
        Property S1 As String
    End Class
    Public Class HPLANET
        Property PLHUSERID As String
        Property PLHID As String
        Property PLANET As String
        Property SIGN As String
        Property DMS As String
        Property HP1 As String
        Property HP2 As String
        Property HP3 As String
        Property HP4 As String
        Property HP5 As String
        Property HP6 As String
        Property HPHOUSE As String
    End Class
    Public Class HCUSP_CALC
        Property CUSPUSERID As String
        Property CUSPHID As String
        Property CUSPID As String
        Property CUSPPLANET As String
        Property CUSPTYPE As String
        Property CUSPCAT As String
    End Class
End Class
