Imports System.Data.SqlClient

Public Class UpdateCusp
    Dim ST(12) As String
    Dim C As String() = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"}
    Sub UpdateCusp(ByRef HID, ByRef UID)
        GetCuspString(HID, UID)
        UpdateCuspString(HID, UID)
        WriteToCUSP(HID, UID)
        WriteToHRAKE(HID, UID)
    End Sub
    Sub GetCuspString(ByRef HID, ByRef UID)
        Dim connection As SqlConnection = New SqlConnection(Connstr.connstr)
        Try
            Dim SelectHCUSP = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "';"
            connection.Open()
            Dim cmd As New SqlCommand(SelectHCUSP, connection)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)
            ST(0) = (ds.Tables(0).Rows(0)).ItemArray(5).ToString() + (ds.Tables(0).Rows(0)).ItemArray(6).ToString() + (ds.Tables(0).Rows(0)).ItemArray(7).ToString()
            ST(1) = (ds.Tables(0).Rows(1)).ItemArray(5).ToString() + (ds.Tables(0).Rows(1)).ItemArray(6).ToString() + (ds.Tables(0).Rows(1)).ItemArray(7).ToString()
            ST(2) = (ds.Tables(0).Rows(2)).ItemArray(5).ToString() + (ds.Tables(0).Rows(2)).ItemArray(6).ToString() + (ds.Tables(0).Rows(2)).ItemArray(7).ToString()
            ST(3) = (ds.Tables(0).Rows(3)).ItemArray(5).ToString() + (ds.Tables(0).Rows(3)).ItemArray(6).ToString() + (ds.Tables(0).Rows(3)).ItemArray(7).ToString()
            ST(4) = (ds.Tables(0).Rows(4)).ItemArray(5).ToString() + (ds.Tables(0).Rows(4)).ItemArray(6).ToString() + (ds.Tables(0).Rows(4)).ItemArray(7).ToString()
            ST(5) = (ds.Tables(0).Rows(5)).ItemArray(5).ToString() + (ds.Tables(0).Rows(5)).ItemArray(6).ToString() + (ds.Tables(0).Rows(5)).ItemArray(7).ToString()
            ST(6) = (ds.Tables(0).Rows(6)).ItemArray(5).ToString() + (ds.Tables(0).Rows(6)).ItemArray(6).ToString() + (ds.Tables(0).Rows(6)).ItemArray(7).ToString()
            ST(7) = (ds.Tables(0).Rows(7)).ItemArray(5).ToString() + (ds.Tables(0).Rows(7)).ItemArray(6).ToString() + (ds.Tables(0).Rows(7)).ItemArray(7).ToString()
            ST(8) = (ds.Tables(0).Rows(8)).ItemArray(5).ToString() + (ds.Tables(0).Rows(8)).ItemArray(6).ToString() + (ds.Tables(0).Rows(8)).ItemArray(7).ToString()
            ST(9) = (ds.Tables(0).Rows(9)).ItemArray(5).ToString() + (ds.Tables(0).Rows(9)).ItemArray(6).ToString() + (ds.Tables(0).Rows(9)).ItemArray(7).ToString()
            ST(10) = (ds.Tables(0).Rows(10)).ItemArray(5).ToString() + (ds.Tables(0).Rows(10)).ItemArray(6).ToString() + (ds.Tables(0).Rows(10)).ItemArray(7).ToString()
            ST(11) = (ds.Tables(0).Rows(11)).ItemArray(5).ToString() + (ds.Tables(0).Rows(11)).ItemArray(6).ToString() + (ds.Tables(0).Rows(11)).ItemArray(7).ToString()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            connection.Close()
        End Try
    End Sub
    Sub UpdateCuspString(ByRef HID As String, ByRef UID As String)
        Dim connection As SqlConnection = New SqlConnection(Connstr.connstr)
        Try
            Dim SelectHPLANET = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HPLANET WHERE PLHUSERID = '" + UID + "' AND PLHID = '" + HID + "';"
            connection.Open()
            Dim cmd As New SqlCommand(SelectHPLANET, connection)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)
            Dim FL As Boolean
            For Each Row As DataRow In ds.Tables(0).Rows
                Select Case Row.Item(11).ToString()
                    Case "01"
                        FL = CHECKSTR(ST(0), Row.Item(2).ToString())
                        If FL = False Then
                            ST(0) = ST(0) + Row.Item(2).ToString()
                        End If
                    Case "02"
                        FL = CHECKSTR(ST(1), Row.Item(2).ToString())
                        If FL = False Then
                            ST(1) = ST(1) + Row.Item(2).ToString()
                        End If
                    Case "03"
                        FL = CHECKSTR(ST(2), Row.Item(2).ToString())
                        If FL = False Then
                            ST(2) = ST(2) + Row.Item(2).ToString()
                        End If
                    Case "04"
                        FL = CHECKSTR(ST(3), Row.Item(2).ToString())
                        If FL = False Then
                            ST(3) = ST(3) + Row.Item(2).ToString()
                        End If
                    Case "05"
                        FL = CHECKSTR(ST(4), Row.Item(2).ToString())
                        If FL = False Then
                            ST(4) = ST(4) + Row.Item(2).ToString()
                        End If
                    Case "06"
                        FL = CHECKSTR(ST(5), Row.Item(2).ToString())
                        If FL = False Then
                            ST(5) = ST(5) + Row.Item(2).ToString()
                        End If
                    Case "07"
                        FL = CHECKSTR(ST(6), Row.Item(2).ToString())
                        If FL = False Then
                            ST(6) = ST(6) + Row.Item(2).ToString()
                        End If
                    Case "08"
                        FL = CHECKSTR(ST(7), Row.Item(2).ToString())
                        If FL = False Then
                            ST(7) = ST(7) + Row.Item(2).ToString()
                        End If
                    Case "09"
                        FL = CHECKSTR(ST(8), Row.Item(2).ToString())
                        If FL = False Then
                            ST(8) = ST(8) + Row.Item(2).ToString()
                        End If
                    Case "10"
                        FL = CHECKSTR(ST(9), Row.Item(2).ToString())
                        If FL = False Then
                            ST(9) = ST(9) + Row.Item(2).ToString()
                        End If
                    Case "11"
                        FL = CHECKSTR(ST(10), Row.Item(2).ToString())
                        If FL = False Then
                            ST(10) = ST(10) + Row.Item(2).ToString()
                        End If
                    Case "12"
                        FL = CHECKSTR(ST(11), Row.Item(2).ToString())
                        If FL = False Then
                            ST(11) = ST(11) + Row.Item(2).ToString()
                        End If
                    Case Else
                End Select
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            connection.Close()
        End Try
    End Sub
    Sub WriteToCUSP(ByRef HID, ByRef UID)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = Connstr.connstr
            con.Open()
            cmd.Connection = con
            Dim flag = False

            For i As Integer = 0 To 11
                cmd.CommandText = "INSERT INTO HEADLETTERS_ENGINE.DBO.CUSP VALUES ('" + UID + "','" + HID + "','" + C(i) + "','" + ST(i).ToUpper() + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub WriteToHRAKE(ByRef HID As String, ByRef UID As String)
        Dim FR1 As String = ""
        Dim FK1 As String = ""
        Dim R1 As String = ""
        Dim R2 As String = ""
        Dim K1 As String = ""
        Dim K2 As String = ""
        Dim SelectHPLANET_RA = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HPLANET WHERE PLHUSERID = '" + UID + "' AND PLHID = '" + HID + "' AND PLANET = 'RA'"
        Dim SelectHPLANET_KE = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HPLANET WHERE PLHUSERID = '" + UID + "' AND PLHID = '" + HID + "' AND PLANET = 'KE'"
        Try
            Dim connection As SqlConnection = New SqlConnection(Connstr.connstr)
            connection.Open()
            Dim cmd As New SqlCommand(SelectHPLANET_RA, connection)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)
            R1 = (ds.Tables(0).Rows(0)).ItemArray(5).ToString()
            R2 = (ds.Tables(0).Rows(0)).ItemArray(6).ToString()
            Dim cmd0 As New SqlCommand(SelectHPLANET_KE, connection)
            Dim da0 As New SqlDataAdapter(cmd0)
            Dim ds0 As New DataSet()
            da0.Fill(ds0)
            K1 = (ds0.Tables(0).Rows(0)).ItemArray(5).ToString()
            K2 = (ds0.Tables(0).Rows(0)).ItemArray(6).ToString()
            connection.Close()

            Dim FLX As Boolean = True
            FLX = CHECKSTR(FR1, R1)
            If FLX = False And R1 IsNot "RA" Then
                FR1 = FR1 + R1
            End If
            FLX = CHECKSTR(FR1, R2)
            If FLX = False And R1 IsNot "RA" Then
                FR1 = FR1 + R2
            End If
            FLX = CHECKSTR(FK1, K1)
            If FLX = False And K1 IsNot "KE" Then
                FK1 = FK1 + K1
            End If
            FLX = CHECKSTR(FK1, K2)
            If FLX = False And K1 IsNot "KE" Then
                FK1 = FK1 + K2
            End If
            'CHECK FOR RA AND KE AS DUPLICATES
            FLX = CHECKSTR(FR1, "KE")
            If FLX = True Then
                FLX = CHECKSTR(FR1, K1)
                If FLX = False And K1 IsNot "RA" Then
                    FR1 = FR1 + K1
                End If
                FLX = CHECKSTR(FR1, K2)
                If FLX = False And K2 IsNot "RA" Then
                    FR1 = FR1 + K2
                End If
            End If
            FLX = CHECKSTR(FK1, "RA")
            If FLX = True Then
                FLX = CHECKSTR(FK1, R1)
                If FLX = False And R1 IsNot "KE" Then
                    FK1 = FK1 + R1
                End If
                If FLX = False And R2 IsNot "KE" Then
                    FK1 = FK1 + R2
                End If
            End If
            Dim InsertHRAKE_RA = "INSERT INTO HEADLETTERS_ENGINE.DBO.HRAKE VALUES ('" + UID + "','" + HID + "','RA','" + FR1.ToUpper() + "');"
            Dim InsertHRAKE_KE = "INSERT INTO HEADLETTERS_ENGINE.DBO.HRAKE VALUES ('" + UID + "','" + HID + "','KE','" + FK1.ToUpper() + "');"
            connection.Open()
            Dim cmd1 As New SqlCommand(InsertHRAKE_RA + InsertHRAKE_KE, connection)
            Dim da1 As New SqlDataAdapter(cmd1)
            Dim ds1 As New DataSet()
            da1.Fill(ds1)
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog.txt")
            IO.File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        End Try
    End Sub
    Function CHECKSTR(S As String, P As String) As Boolean
        Dim X = S.Length / 2
        Dim OV = False
        Dim START = 0
        For I As Integer = 0 To X - 1
            If (S.Substring(START, 2) = P) Then
                OV = True
                Exit For
            End If
            START = START + 2
        Next
        Return OV
    End Function
End Class
