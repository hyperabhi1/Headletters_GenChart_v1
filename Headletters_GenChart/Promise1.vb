Imports System.Data.SqlClient

Public Class Promise1
    Dim UID = ""
    Dim HID = ""
    'Public connstr = "data source=DESKTOP-JBRFH9E;initial catalog=HEADLETTERS_ENGINE;integrated security=True;"
    'Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS_ENGINE;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;multipleactiveresultsets=True;User Id=sa;password=pSI)TA1t0K[);"
    Sub Main(ByVal UID_ As String, ByVal HID_ As String)
        UID = UID_
        HID = HID_
        Dim DAILY_RULES_SUBList As New List(Of DAILY_RULES_SUB)
        Dim DAILY_RULES_TEMPLATEList As New List(Of DAILY_RULES_TEMPLATE)
        Dim HCUSP_PROMISE_LINKList As New List(Of HCUSP_PROMISE_LINK)
        Dim DAILY_RULES_MAINList As New List(Of DAILY_RULES_MAIN)
        Select_From_DAILY_RULES_TEMPLATE(UID, HID, DAILY_RULES_TEMPLATEList)
        Update_DAILY_RULES_MAIN(UID, HID, DAILY_RULES_TEMPLATEList)
        Select_From_DAILY_RULES_MAIN(UID, HID, DAILY_RULES_MAINList)
        Select_From_DAILY_RULES_SUB(UID, HID, DAILY_RULES_SUBList)
        Select_From_HCUSP_PROMISE_LINK(UID, HID, HCUSP_PROMISE_LINKList)
        Process_Promise1(UID, HID, DAILY_RULES_MAINList, DAILY_RULES_SUBList, HCUSP_PROMISE_LINKList)
    End Sub
    Sub Select_From_DAILY_RULES_TEMPLATE(ByVal UID As String, ByVal HID As String, ByRef DAILY_RULES_TEMPLATEList As List(Of DAILY_RULES_TEMPLATE))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.DAILY_RULES_TEMPLATE;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim Daily_Rules_Template = New DAILY_RULES_TEMPLATE With {
                    .MAIN = reader("MAIN").ToString().Trim,
                    .S1 = reader("S1").ToString().Trim,
                    .S2 = reader("S2").ToString().Trim,
                    .S3 = reader("S3").ToString().Trim,
                    .S4 = reader("S4").ToString().Trim,
                    .S5 = reader("S5").ToString().Trim,
                    .S6 = reader("S6").ToString().Trim,
                    .RULESID = reader("RULESID").ToString().Trim,
                    .CCOUNT = If(IsDBNull(reader("CCOUNT")), 0, reader("CCOUNT")),
                    .DESCRIPTION = reader("DESCRIPTION").ToString().Trim,
                    .CATEGORY = reader("CATEGORY").ToString().Trim,
                    .RTOTAL = If(IsDBNull(reader("RTOTAL")), 0, reader("RTOTAL")),
                    .CUSPINDEX = reader("CUSPINDEX").ToString().Trim,
                    .DUP = reader("DUP").ToString().Trim
                }
                DAILY_RULES_TEMPLATEList.Add(Daily_Rules_Template)
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Update_DAILY_RULES_MAIN(ByVal UID As String, ByVal HID As String, ByRef DAILY_RULE_TEMPLATEList As List(Of DAILY_RULES_TEMPLATE))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = $";"
            cmd.ExecuteNonQuery()
            cmd.CommandText = $"DELETE FROM HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            cmd.ExecuteNonQuery()
            For i As Integer = 0 To DAILY_RULE_TEMPLATEList.Count - 1
                cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN VALUES('" + UID + "', '" + HID + "','" + DAILY_RULE_TEMPLATEList(i).MAIN + "',
                                    '" + DAILY_RULE_TEMPLATEList(i).S1 + "', '" + DAILY_RULE_TEMPLATEList(i).S2 + "', '" + DAILY_RULE_TEMPLATEList(i).S3 + "',
                                    '" + DAILY_RULE_TEMPLATEList(i).S4 + "', '" + DAILY_RULE_TEMPLATEList(i).S5 + "', '" + DAILY_RULE_TEMPLATEList(i).S6 + "',
                                    '" + DAILY_RULE_TEMPLATEList(i).RULESID + "'," + DAILY_RULE_TEMPLATEList(i).CCOUNT.ToString() + ",'" + DAILY_RULE_TEMPLATEList(i).DESCRIPTION + "',
                                    NULL, NULL, '" + DAILY_RULE_TEMPLATEList(i).CATEGORY + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_DAILY_RULES_MAIN(ByVal UID As String, ByVal HID As String, ByRef DAILY_RULES_MAINList As List(Of DAILY_RULES_MAIN))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim Daily_Rules_Main = New DAILY_RULES_MAIN With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .MAIN = reader("MAIN").ToString().Trim,
                    .S1 = reader("S1").ToString().Trim,
                    .S2 = reader("S2").ToString().Trim,
                    .S3 = reader("S3").ToString().Trim,
                    .S4 = reader("S4").ToString().Trim,
                    .S5 = reader("S5").ToString().Trim,
                    .S6 = reader("S6").ToString().Trim,
                    .RULESID = reader("RULESID").ToString().Trim,
                    .CCOUNT = If(IsDBNull(reader("CCOUNT")), 0, reader("CCOUNT")),
                    .DESCRIPTION = reader("DESCRIPTION").ToString().Trim,
                    .RTOTAL = If(IsDBNull(reader("RTOTAL")), 0, reader("RTOTAL")),
                    .FREQ = reader("FREQ").ToString().Trim,
                    .CATEGORY = reader("CATEGORY").ToString().Trim
                }
                DAILY_RULES_MAINList.Add(Daily_Rules_Main)
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_DAILY_RULES_SUB(ByVal UID As String, ByVal HID As String, ByRef DAILY_RULES_SUBList As List(Of DAILY_RULES_SUB))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.DAILY_RULES_SUB;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim Daily_Rules_Sub = New DAILY_RULES_SUB With {
                    .RULESID = reader("RULESID").ToString().Trim,
                    .C1 = reader("C1").ToString().Trim,
                    .C1PM = reader("C1PM").ToString().Trim,
                    .C2 = reader("C2").ToString().Trim,
                    .C2PM = reader("C2PM").ToString().Trim,
                    .C3 = reader("C3").ToString().Trim,
                    .C3PM = reader("C3PM").ToString().Trim,
                    .C4 = reader("C4").ToString().Trim,
                    .C4PM = reader("C4PM").ToString().Trim,
                    .C5 = reader("C5").ToString().Trim,
                    .C5PM = reader("C5PM").ToString().Trim
                }
                DAILY_RULES_SUBList.Add(Daily_Rules_Sub)
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_HCUSP_PROMISE_LINK(ByVal UID As String, ByVal HID As String, ByRef HCUSP_PROMISE_LINKList As List(Of HCUSP_PROMISE_LINK))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim HCusp_Promise_Link = New HCUSP_PROMISE_LINK With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .CUSPKEY = reader("CUSPKEY").ToString().Trim,
                    .TCOUNT = If(IsDBNull(reader("TCOUNT")), 0, reader("TCOUNT")),
                    .SKEY = reader("SKEY").ToString().Trim
                }
                HCUSP_PROMISE_LINKList.Add(HCusp_Promise_Link)
            End While
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub Process_Promise1(ByVal UID As String, ByVal HID As String, ByVal DAILY_RULES_MAINList As List(Of DAILY_RULES_MAIN), ByVal DAILY_RULES_SUBList As List(Of DAILY_RULES_SUB), ByVal HCUSP_PROMISE_LINKList As List(Of HCUSP_PROMISE_LINK))
        For i As Integer = 0 To DAILY_RULES_MAINList.Count - 1
            Dim mr = DAILY_RULES_MAINList.Item(i).RULESID
            Dim FoundDRS = (From DAILY_RULES_SUB In DAILY_RULES_SUBList
                            Where DAILY_RULES_SUB.RULESID = mr
                            Select DAILY_RULES_SUB).ToList()
            If FoundDRS.Count > 0 Then
                Dim pos = 0
                Dim neg = 0
                Dim mc1 = FoundDRS.Item(0).C1
                Dim mc1pm = FoundDRS.Item(0).C1PM
                Dim mc2 = FoundDRS.Item(0).C2
                Dim mc2pm = FoundDRS.Item(0).C2PM
                Dim mc3 = FoundDRS.Item(0).C3
                Dim mc3pm = FoundDRS.Item(0).C3PM
                Dim mc4 = FoundDRS.Item(0).C4
                Dim mc4pm = FoundDRS.Item(0).C4PM
                Dim mc5 = FoundDRS.Item(0).C5
                Dim mc5pm = FoundDRS.Item(0).C5PM

                Dim FoundHPL_MC1 = (From HCUSP_PROMISE_LINK In HCUSP_PROMISE_LINKList
                                    Where HCUSP_PROMISE_LINK.UID = UID And HCUSP_PROMISE_LINK.HID = HID And HCUSP_PROMISE_LINK.CUSPKEY = mc1
                                    Select HCUSP_PROMISE_LINK).ToList()
                If FoundHPL_MC1.Count > 0 Then
                    Dim mc = FoundHPL_MC1.Item(0).TCOUNT
                    If mc1pm = "+" Then
                        pos = pos + mc
                    Else
                        neg = neg + mc
                    End If
                End If

                Dim FoundHPL_MC2 = (From HCUSP_PROMISE_LINK In HCUSP_PROMISE_LINKList
                                    Where HCUSP_PROMISE_LINK.UID = UID And HCUSP_PROMISE_LINK.HID = HID And HCUSP_PROMISE_LINK.CUSPKEY = mc2
                                    Select HCUSP_PROMISE_LINK).ToList()
                If FoundHPL_MC2.Count > 0 Then
                    Dim mc = FoundHPL_MC2.Item(0).TCOUNT
                    If mc2pm = "+" Then
                        pos = pos + mc
                    Else
                        neg = neg + mc
                    End If
                End If

                Dim FoundHPL_MC3 = (From HCUSP_PROMISE_LINK In HCUSP_PROMISE_LINKList
                                    Where HCUSP_PROMISE_LINK.UID = UID And HCUSP_PROMISE_LINK.HID = HID And HCUSP_PROMISE_LINK.CUSPKEY = mc3
                                    Select HCUSP_PROMISE_LINK).ToList()
                If FoundHPL_MC3.Count > 0 Then
                    Dim mc = FoundHPL_MC3.Item(0).TCOUNT
                    If mc3pm = "+" Then
                        pos = pos + mc
                    Else
                        neg = neg + mc
                    End If
                End If

                Dim FoundHPL_MC4 = (From HCUSP_PROMISE_LINK In HCUSP_PROMISE_LINKList
                                    Where HCUSP_PROMISE_LINK.UID = UID And HCUSP_PROMISE_LINK.HID = HID And HCUSP_PROMISE_LINK.CUSPKEY = mc4
                                    Select HCUSP_PROMISE_LINK).ToList()
                If FoundHPL_MC4.Count > 0 Then
                    Dim mc = FoundHPL_MC4.Item(0).TCOUNT
                    If mc4pm = "+" Then
                        pos = pos + mc
                    Else
                        neg = neg + mc
                    End If
                End If

                Dim FoundHPL_MC5 = (From HCUSP_PROMISE_LINK In HCUSP_PROMISE_LINKList
                                    Where HCUSP_PROMISE_LINK.UID = UID And HCUSP_PROMISE_LINK.HID = HID And HCUSP_PROMISE_LINK.CUSPKEY = mc5
                                    Select HCUSP_PROMISE_LINK).ToList()
                If FoundHPL_MC5.Count > 0 Then
                    Dim mc = FoundHPL_MC5.Item(0).TCOUNT
                    If mc5pm = "+" Then
                        pos = pos + mc
                    Else
                        neg = neg + mc
                    End If
                End If

                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                Try
                    con.ConnectionString = connstr
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = $"UPDATE HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN SET RTOTAL = " + (pos - neg).ToString() + " WHERE RULESID = '" + mr + "' AND UID  = '" + UID + "' AND HID = '" + HID + "';"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
                    IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
                Finally
                    con.Close()
                End Try
            End If
        Next
        For i As Integer = 0 To DAILY_RULES_MAINList.Count - 1
            Dim Diff = DAILY_RULES_MAINList.Item(i).CCOUNT - DAILY_RULES_MAINList.Item(i).RTOTAL
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            Try
                con.ConnectionString = connstr
                con.Open()
                cmd.Connection = con

                If Diff > 2 Then
                    cmd.CommandText = $"UPDATE HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN SET FREQ = 'low/medium' WHERE RULESID = '" + DAILY_RULES_MAINList.Item(i).RULESID + "' AND UID  = '" + UID + "' AND HID = '" + HID + "';"
                    cmd.ExecuteNonQuery()
                Else
                    cmd.CommandText = $"UPDATE HEADLETTERS_ENGINE.DBO.DAILY_RULES_MAIN SET FREQ = 'high' WHERE RULESID = '" + DAILY_RULES_MAINList.Item(i).RULESID + "' AND UID  = '" + UID + "' AND HID = '" + HID + "';"
                    cmd.ExecuteNonQuery()
                End If
            Catch ex As Exception
                Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\Promise_Error.txt")
                IO.File.AppendAllText(strFile, String.Format(vbCrLf + "UID=> " + UID + "  ::  HID=> " + HID + vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
            Finally
                con.Close()
            End Try
        Next
    End Sub

    Class DAILY_RULES_SUB
        Property RULESID As String
        Property C1 As String
        Property C1PM As String
        Property C2 As String
        Property C2PM As String
        Property C3 As String
        Property C3PM As String
        Property C4 As String
        Property C4PM As String
        Property C5 As String
        Property C5PM As String
    End Class
    Class DAILY_RULES_TEMPLATE
        Property MAIN As String
        Property S1 As String
        Property S2 As String
        Property S3 As String
        Property S4 As String
        Property S5 As String
        Property S6 As String
        Property RULESID As String
        Property CCOUNT As Integer
        Property DESCRIPTION As String
        Property CATEGORY As String
        Property RTOTAL As Integer
        Property CUSPINDEX As String
        Property DUP As String
    End Class
    Class HCUSP_SIGN_SUB
        Property UID As String
        Property HID As String
        Property CUSP As String
        Property PLANET As String
    End Class
    Class HCUSP_PROMISE_LINK
        Property UID As String
        Property HID As String
        Property CUSPKEY As String
        Property TCOUNT As Integer
        Property SKEY As String
    End Class
    Public Class DAILY_RULES_MAIN
        Property UID As String
        Property HID As String
        Property MAIN As String
        Property S1 As String
        Property S2 As String
        Property S3 As String
        Property S4 As String
        Property S5 As String
        Property S6 As String
        Property RULESID As String
        Property CCOUNT As Integer
        Property DESCRIPTION As String
        Property RTOTAL As Integer
        Property FREQ As String
        Property CATEGORY As String
    End Class

End Class
