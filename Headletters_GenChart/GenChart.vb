Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Module Connstr
    'Public connstr = "data source=DESKTOP-JBRFH9E;initial catalog=HEADLETTERS_ENGINE;integrated security=True;"
    'Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS_ENGINE;integrated security=True;multipleactiveresultsets=True;"
    'Public connstr = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;multipleactiveresultsets=True;User Id=sa;password=pSI)TA1t0K[);"

End Module
Public Class GenChart
    Public Shared Sub Main()
        Dim UID = ""
        Dim HID = ""
        Dim personalDetails As PersonalDetails = New PersonalDetails()
        Dim P_list(12) As String
        Dim H_List(12) As String
        Dim HstrCusp(12) As String
        Dim Hplanets(12) As String
        Dim BirthLagna(12, 2) As String
        Dim BirthBhav(12, 2) As String
        Dim BirthSouth(12) As String
        Dim Horo As New TASystem.TrueAstro

        Dim connection As SqlConnection = New SqlConnection(Connstr.connstr)
        Try
            Dim cmd As New SqlCommand($"SELECT HUSERID, HID, RECTIFIEDDATE, RECTIFIEDTIME, RECTIFIEDDST, RECTIFIEDPLACE, RECTIFIEDLONGTITUDE, RECTIFIEDLONGTITUDEEW, RECTIFIEDLATITUDE,
                                        RECTIFIEDLATITUDENS, RECTIFIEDTIMEZONE, HNAME, HDOBNATIVE, HPLACE, HHOURS, HMIN, HSS, HAMPM, HMARRIAGEDATE, HFIRSTCHILDPLACE, HCRDATE, HDRR, HDRRD
	                                    FROM ASTROLOGYSOFTWARE_DB.DBO.HMAIN where HMAIN.HPDF IS NULL AND RECTIFIEDTIME IS NOT NULL AND RECTIFIEDDATE IS NOT NULL AND RECTIFIEDDST IS NOT NULL 
                                        AND RECTIFIEDTIMEZONE IS NOT NULL AND HSTATUS = '2';", connection)
            Dim da As New SqlDataAdapter(cmd)
            Dim RowsData As New DataSet()
            da.Fill(RowsData)
            For i As Integer = 0 To RowsData.Tables(0).Rows.Count - 1
                Try
                    Dim DateTimeB As New NameValueCollection
                    Dim PlaceDataB As New NameValueCollection
                    Dim BData As New NameValueCollection
                    Dim PlaceData As New NameValueCollection
                    UID = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                    HID = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                    Dim RECTIFIEDDATE = CType(RowsData.Tables(0).Rows(i)(2), DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                    Dim RECTIFIEDTIME = CType(RowsData.Tables(0).Rows(i)(3).ToString(), DateTime).ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)
                    Dim RECTIFIEDDST = RowsData.Tables(0).Rows(i)(4).ToString().Trim
                    Dim RECTIFIEDPLACE = RowsData.Tables(0).Rows(i)(5).ToString().Trim
                    Dim RECTIFIEDLONGTITUDE = RowsData.Tables(0).Rows(i)(6).Trim.ToString().Trim
                    Dim RECTIFIEDLONGTITUDEEW = RowsData.Tables(0).Rows(i)(7).ToString().Trim
                    Dim RECTIFIEDLATITUDE = RowsData.Tables(0).Rows(i)(8).ToString().Trim
                    Dim RECTIFIEDLATITUDENS = RowsData.Tables(0).Rows(i)(9).ToString().Trim
                    Dim RECTIFIEDTZ = RowsData.Tables(0).Rows(i)(10).ToString().Trim
                    Dim Year = RECTIFIEDDATE.Split("-")(0)
                    Dim Month = RECTIFIEDDATE.Split("-")(1)
                    Dim Day = RECTIFIEDDATE.Split("-")(2).Substring(0, 2)
                    Dim Hour = RECTIFIEDTIME.Split(":")(0)
                    Dim Min = RECTIFIEDTIME.Split(":")(1)
                    Dim Sec = RECTIFIEDTIME.Split(":")(2).Substring(0, 2)
                    Dim MSec = RECTIFIEDTIME.Split(".")(1)
                    BData.Add("Year", Year)
                    BData.Add("Month", Month)
                    BData.Add("Day", Day)
                    BData.Add("Hour", Hour)
                    BData.Add("Min", Min)
                    BData.Add("Sec", Sec)
                    BData.Add("mSec", MSec)
                    DateTimeB = BData
                    Dim Place = ""
                    Dim State = ""
                    Dim Country = ""
                    If RECTIFIEDPLACE.Split("-").Length = 3 Then
                        Place = RECTIFIEDPLACE.Split("-")(0)
                        State = RECTIFIEDPLACE.Split("-")(1)
                        Country = RECTIFIEDPLACE.Split("-")(2)
                    ElseIf RECTIFIEDPLACE.Split("-").Length = 2 Then
                        Place = RECTIFIEDPLACE.Split("-")(0)
                        State = "NA"
                        Country = RECTIFIEDPLACE.Split("-")(2)
                    ElseIf RECTIFIEDPLACE.Split("-").Length = 1 Then
                        Place = RECTIFIEDPLACE.Split("-")(0)
                        State = "NA"
                        Country = "NA"
                    ElseIf RECTIFIEDPLACE.Contains("-") <> False Then
                        Place = RECTIFIEDPLACE
                        State = "NA"
                        Country = "NA"
                    Else
                        Throw New System.Exception("Invalid RECTIFIEDPLACE format exception.")
                    End If
#Region "Personal Details"
                    personalDetails.DateofBirth = RECTIFIEDDATE.Split(" ")(0)
                    personalDetails.DayofBirth = GetDayOfTheWeek(DateTime.Parse(RECTIFIEDDATE.Split(" ")(0).Replace("-", "/")).DayOfWeek)
                    personalDetails.TimeofBirth = RECTIFIEDDATE.Split(" ")(1)
                    personalDetails.PlaceofBirth = RECTIFIEDPLACE
                    personalDetails.Latitude = RECTIFIEDLATITUDE.Remove(2, 1).Insert(2, "°").Remove(5, 1).Insert(5, "'") & """"
                    personalDetails.Longitude = RECTIFIEDLONGTITUDE.Remove(2, 1).Insert(2, "°").Remove(5, 1).Insert(5, "'") & """"
                    personalDetails.NameoftheChartOwner = RowsData.Tables(0).Rows(i)(11).ToString().Trim
                    personalDetails.Rashi = "Not yet provided by DLL"
                    personalDetails.Star = "Not yet provided by DLL"
                    personalDetails.SunSign = "Not yet provided by DLL"
                    personalDetails.BalanceDasa = "Not yet provided by DLL"
                    personalDetails.OriginalDOBByChartOwner = RowsData.Tables(0).Rows(i)(12).ToShortDateString
                    personalDetails.OriginalPOBByChartOwner = RowsData.Tables(0).Rows(i)(13).ToString().Trim
                    personalDetails.OriginalTOBByChartOwner = RowsData.Tables(0).Rows(i)(14).ToString() + ":" + RowsData.Tables(0).Rows(i)(15).ToString() + ":" + RowsData.Tables(0).Rows(i)(16).ToString() + " " + RowsData.Tables(0).Rows(i)(17).Trim
                    personalDetails.Marriage = RowsData.Tables(0).Rows(i)(18).ToString().Trim
                    personalDetails.FirstChild = RowsData.Tables(0).Rows(i)(19).ToString().Trim
                    personalDetails.LastCallRecieved = RowsData.Tables(0).Rows(i)(20).ToString().Trim
                    personalDetails.DemiseOfRelatives = RowsData.Tables(0).Rows(i)(21).ToString().Trim & " at: " + RowsData.Tables(0).Rows(i)(22).ToString().Trim
#End Region
                    Dim lon, lat, geoLat, Tz As Double
                    lon = RECTIFIEDLONGTITUDE.Split("^")(0) + RECTIFIEDLONGTITUDE.Split("^")(1) / 60 + RECTIFIEDLONGTITUDE.Split("^")(2) / 3600
                    geoLat = RECTIFIEDLATITUDE.Split("^")(0) + RECTIFIEDLATITUDE.Split("^")(1) / 60 + RECTIFIEDLATITUDE.Split("^")(2) / 3600
                    Dim eW = RECTIFIEDLONGTITUDEEW
                    Dim nS = RECTIFIEDLATITUDENS
                    If eW = "W" Then
                        lon = -lon
                        Tz = -Tz
                    End If
                    lat = (Math.Atan(Math.Tan(geoLat * Math.PI / 180) * 0.99330546)) * 180 / Math.PI
                    If nS = "S" Then
                        lat = -lat
                    End If
                    Dim TzB = Val(RECTIFIEDTZ.Split(":")(0).Substring(1, 2)) + RECTIFIEDTZ.Split(":")(1) / 60
                    Dim DstB = RECTIFIEDDST
                    PlaceData.Add("Place", Place)
                    PlaceData.Add("State", State)
                    PlaceData.Add("Country", Country)
                    PlaceData.Add("lonB", lon)
                    PlaceData.Add("latB", lat)
                    PlaceData.Add("eW", eW)
                    PlaceData.Add("nS", nS)
                    PlaceData.Add("TzB", TzB)
                    PlaceData.Add("DstB", DstB)
                    PlaceDataB = PlaceData
                    Horo.getBirthPlanetCusp(DateTimeB, PlaceDataB, P_list, H_List)
                    Horo.getBirthLagnaBhav(DateTimeB, PlaceDataB, BirthLagna, BirthBhav, BirthSouth)
                    Dim DasaListP As New DataTable
                    DasaListP.Columns.Add("Dasha")
                    DasaListP.Columns.Add("Bhukti")
                    DasaListP.Columns.Add("Antara")
                    DasaListP.Columns.Add("Suk")
                    DasaListP.Columns.Add("Pra")
                    DasaListP.Columns.Add("Cl_Date")
                    Horo.getBirthDasaDBA(DateTimeB, PlaceDataB, DasaListP)
                    Dim genChart As GenChart = New GenChart()
                    genChart.UpdateHCUSP(HID, UID, H_List)
                    genChart.UpdateHPLANET(HID, UID, P_list)
                    genChart.UpdateHDASA(HID, UID, DasaListP)
                    Dim updateCusp As UpdateCusp = New UpdateCusp()
                    updateCusp.UpdateCusp(HID, UID)
                    Dim genHTMLChart As GenHTMLChart = New GenHTMLChart()
                    genHTMLChart.GenHTMLChartMain(HID, UID, P_list, H_List, BirthLagna, BirthBhav, BirthSouth, DasaListP, personalDetails)
                    genChart.UpdateStatus(HID, UID)
                Catch ex As Exception
                    Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
                    File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
                    Continue For
                End Try
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
            File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            connection.Close()
        End Try
    End Sub
    Public Shared Function GetDayOfTheWeek(ByVal day As String) As String
        Dim DayOfWeek() = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        Return DayOfWeek(Convert.ToInt32(day))
    End Function
    Sub UpdateHCUSP(ByRef HID As String, ByRef UID As String, ByRef H_List As String())
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = Connstr.connstr
            con.Open()
            cmd.Connection = con
            Dim flag = False
            For Each rowData In H_List
                If flag = False Then
                    flag = True
                    Continue For
                End If
                cmd.CommandText = cmd.CommandText + "INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP VALUES ('" + UID + "','" + HID + "','" + GetCuspNo(rowData.Split("|")(0)) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "');" + vbCrLf
            Next
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
            File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub UpdateHPLANET(ByRef HID As String, ByRef UID As String, ByRef P_List As String())
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = Connstr.connstr
            con.Open()
            cmd.Connection = con
            Dim flag = False
            For Each rowData In P_List
                If flag = False Then
                    flag = True
                    Continue For
                End If
                cmd.CommandText = cmd.CommandText + "INSERT INTO HEADLETTERS_ENGINE.DBO.HPLANET VALUES ('" + UID + "','" + HID + "','" + GetPlanetShortName(rowData.Split("|")(0)) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "','" + rowData.Split("|")(6) + "','" + rowData.Split("|")(7) + "','" + rowData.Split("|")(8) + "','" + GetCuspPrefix(rowData.Split("|")(10)) + "');" + vbCrLf
            Next
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
            File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub UpdateHDASA(ByRef HID As String, ByRef UID As String, ByRef DasaListP As DataTable)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = Connstr.connstr
            con.Open()
            cmd.Connection = con
            Dim flag = False
            For i As Integer = 0 To DasaListP.Rows.Count - 1
                cmd.CommandText = "INSERT INTO HEADLETTERS_ENGINE.DBO.HDASA VALUES ('" + UID + "','" + HID + "','" + ((DasaListP.Rows.Item(i)).ItemArray(0) + (DasaListP.Rows.Item(i)).ItemArray(1) + (DasaListP.Rows.Item(i)).ItemArray(2) + (DasaListP.Rows.Item(i)).ItemArray(3) + (DasaListP.Rows.Item(i)).ItemArray(4)).ToString().ToUpper() + "','" + Convert.ToString(DasaListP.Rows.Item(i).ItemArray(5)).Split(" :: ")(0) + "','" + Convert.ToString(DasaListP.Rows.Item(i).ItemArray(5)).Split(" :: ")(2) + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
            File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Sub UpdateStatus(ByRef HID As String, ByRef UID As String)
        Dim con As New SqlConnection
        Dim command As New SqlCommand
        Try
            con.ConnectionString = Connstr.connstr
            con.Open()
            command.Connection = con
            command.CommandText = $"UPDATE ASTROLOGYSOFTWARE_DB.DBO.HREQUEST SET REQCAT = '7' WHERE RQUSERID = '" + UID + "' AND RQHID = '" + HID + "'  AND HREQUEST.REQCAT = '9' AND HREQUEST.RQUNREAD = 'Y';
                                    UPDATE ASTROLOGYSOFTWARE_DB.DBO.HMAIN SET HSTATUS = '5' WHERE HUSERID = '" + UID + "' AND HID = '" + HID + "';"
            command.ExecuteNonQuery()
        Catch ex As Exception
            Dim strFile As String = String.Format("C:\Astro\ServiceLogs\ChartGeneration\ErrorLog_{1}_{0}.txt", HID + UID, DateTime.Today.ToString("ddMMMyyyy"))
            File.AppendAllText(strFile, String.Format(vbCrLf + "Error Occured at-- {0}{1}{2}", Environment.NewLine + DateTime.Now, Environment.NewLine, ex.Message + vbCrLf + ex.StackTrace))
        Finally
            con.Close()
        End Try
    End Sub
    Function GetCuspNo(ByVal s As String)
        Dim ov = "NA"
        Select Case s
            Case "Cusp I"
                ov = "01"
            Case "Cusp II"
                ov = "02"
            Case "Cusp III"
                ov = "03"
            Case "Cusp IV"
                ov = "04"
            Case "Cusp V"
                ov = "05"
            Case "Cusp VI"
                ov = "06"
            Case "Cusp VII"
                ov = "07"
            Case "Cusp VIII"
                ov = "08"
            Case "Cusp IX"
                ov = "09"
            Case "Cusp X"
                ov = "10"
            Case "Cusp XI"
                ov = "11"
            Case "Cusp XII"
                ov = "12"
            Case Else
        End Select
        Return ov
    End Function
    Function GetPlanetShortName(ByVal Pl As String)
        Dim Planet As String = "Not_A_Planet"
        Select Case Pl
            Case "Mars"
                Planet = "MA"
            Case "Venus"
                Planet = "VE"
            Case "Saturn"
                Planet = "SA"
            Case "Jupiter"
                Planet = "JU"
            Case "Sun"
                Planet = "SU"
            Case "Moon"
                Planet = "MO"
            Case "Mercury"
                Planet = "ME"
            Case "T.Rahu"
                Planet = "RA"
            Case "T.Ketu"
                Planet = "KE"
            Case "Neptune"
                Planet = "NE"
            Case "Pluto"
                Planet = "PL"
            Case "Uranus"
                Planet = "UR"
        End Select
        Return Planet
    End Function
    Function GetCuspPrefix(ByVal s As String)
        Dim ov = "NA"
        Select Case s
            Case "1"
                ov = "01"
            Case "2"
                ov = "02"
            Case "3"
                ov = "03"
            Case "4"
                ov = "04"
            Case "5"
                ov = "05"
            Case "6"
                ov = "06"
            Case "7"
                ov = "07"
            Case "8"
                ov = "08"
            Case "9"
                ov = "09"
            Case "10"
                ov = "10"
            Case "11"
                ov = "11"
            Case "12"
                ov = "12"
            Case Else
        End Select
        Return ov
    End Function
End Class
