Imports System
Imports System.IO
Imports System.Security.Cryptography

Imports System.Runtime.InteropServices
Module functions
    'For Databases
    Public db As New ADODB.Connection
    Public cdb As New ADODB.Connection

    Public enc As System.Text.UTF8Encoding
    Public encryptor As ICryptoTransform
    Public decryptor As ICryptoTransform

    'For Recordset

    Public rs As ADODB.Recordset
    Public rs1 As ADODB.Recordset

    Public mrs As ADODB.Recordset
    Public nocmrs As ADODB.Recordset
    Public cmenu As String = Nothing
    Public sFmat As String = "###,###,###,##0.00"
    Public svid As Integer = 0
    Public mytab As Integer = 0
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const CTRL As Integer = &H2
    Public Const WM_HOTKEY As Integer = &H312

    <DllImport("User32.dll")> _
    Public Function RegisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer, ByVal fsModifiers As Integer, _
                        ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")> _
    Public Function UnregisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer) As Integer
    End Function
    Public Sub connecttodb()


        'Try
        'Process.Start(Application.StartupPath + "\ip.txt")
        Try
            Dim sr As New System.IO.StreamReader("ip2.txt")
            Dim mline As String = Nothing
            Dim mline2 As String = Nothing
            Dim mline3 As String = Nothing
            mline = sr.ReadLine
            mline2 = sr.ReadLine
            mline3 = sr.ReadLine
            sr.Close()


            db = New ADODB.Connection

            rs = New ADODB.Recordset
            rs1 = New ADODB.Recordset

            'readtext()

            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs1.CursorLocation = ADODB.CursorLocationEnum.adUseClient

            'db.Open()
            db.Open("Driver={MySQL ODBC 5.1 Driver};Database=dsegsurvey;Server=" & mline & ";User=" & mline2 & ";Password=" & mline3 & "")
            DSEGMAIN.tstat1.Text = "Database Connection Successful"
            DSEGMAIN.tstat1.ForeColor = Color.Blue


            rs.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
            rs.LockType = ADODB.LockTypeEnum.adLockOptimistic

            rs1.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
            rs1.LockType = ADODB.LockTypeEnum.adLockOptimistic
        Catch ex As Exception
            DSEGMAIN.tstat1.Text = ex.ToString
            DSEGMAIN.tstat1.ForeColor = Color.Red
            connecttodb()
            Exit Sub
        End Try

        'Catch e As Exception
        '    MsgBox(Err.Description)
        'End Try

        'ts.Close()

    End Sub
    Public Function GetCurrentAge(ByVal dob As Date) As Integer
        Dim age As Integer
        age = Today.Year - dob.Year
        If (dob > Today.AddYears(-age)) Then age -= 1
        Return age
    End Function
    Public Function SafeImageFromFile(path As String) As Image
        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read)
            Dim img = Image.FromStream(fs)
            Return img
        End Using
    End Function


    Public Sub connecttomdb()
        On Error GoTo errh

        cdb = New ADODB.Connection

        cdb.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:/myreports.mdb;Jet OLEDB:Database Password=;")
errh:
        DSEGMAIN.tstat1.Text = Err.Description
        '  DSEGMAIN.lblerror.Text = Err.Description
    End Sub
    Function myd(str As String)

        myd = Replace(str, ",", "")
        Return myd

    End Function
    Function cq(str As String)
        cq = Trim$(Replace(str, "'", "''"))
        Return cq
    End Function
    Function gcyear()

        Dim cyrr As New ADODB.Recordset
        sopen(cyrr)
        cyrr.Open("SELECT DATE_FORMAT(CURDATE(),'%Y') myy", db)
        Dim y As String = Nothing
        y = cyrr.Fields("myy").Value
        cyrr.Close()
        Return y
    End Function
    Function checkdbase()


        Dim cd As New ADODB.Recordset
        sopen(cd)
        cd.Open("SELECT DATE_FORMAT(CURDATE(),'%Y') myy", db)
        Dim y As String = Nothing
        y = cd.Fields("myy").Value
        cd.Close()

        Dim mydb As String = Nothing
        mydb = "maritime" & y

        cd.Open("SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & mydb & "' ", db)
        If cd.EOF = True Then
            Return 1
        Else
            Return 2
        End If
        cd.Close()



    End Function
    Function sopen(mrec As ADODB.Recordset)
        mrec.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        mrec.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        mrec.LockType = ADODB.LockTypeEnum.adLockOptimistic
        Return True

    End Function
    Function onlyletters(y As Integer, x As Keys)
        Dim st As Integer = 0
        'MsgBox(x)
        If Not (x = 8) Then
            If (x >= 48 And x <= 57) Then
                st = 1
            End If
        End If
        Return st


    End Function
    Function getdatetime(stat As String)
        Dim dtg As New ADODB.Recordset
        Dim data As String = Nothing
        Dim gdate, gtime As String
        Dim gdy As String

        sopen(dtg)

        dtg.Open("SELECT DATE_FORMAT(CURDATE(),'%Y-%m-%d') mdate,TIME_FORMAT(CURTIME(),'%h:%i:%s %p') mtime", db)
        gtime = dtg.Fields("mdate").Value & " " & dtg.Fields("mtime").Value
        gdate = dtg.Fields("mdate").Value
        gdy = Strings.Left(dtg.Fields("mdate").Value, 4)
        dtg.Close()

        If stat = 1 Then
            data = gtime
        ElseIf stat = 2 Then
            data = gdate
        ElseIf stat = 3 Then
            data = gdy
        End If

        Return data

    End Function
    Function getimeonly()
        Dim dtm As New ADODB.Recordset
        Dim cdata As String = Nothing

        sopen(dtm)

        dtm.Open("SELECT TIME_FORMAT(CURTIME(),'%h:%i:%s') ctime", db)
        cdata = dtm.Fields("ctime").Value
        dtm.Close()

        Return cdata
    End Function
    Function createdbase()
        Dim cd As New ADODB.Recordset
        sopen(cd)
        cd.Open("SELECT DATE_FORMAT(CURDATE(),'%Y') myy", db)
        Dim y As String = Nothing
        y = cd.Fields("myy").Value
        cd.Close()

        Dim mydbase As String = Nothing
        mydbase = "maritime" & y
        db.Execute("create database " & mydbase & "")
        cd.Open("select table_name from information_schema.tables where table_schema ='maritime'")
        Dim mytab As String = Nothing
        While Not cd.EOF
            mytab = cd.Fields("table_name").Value
            db.Execute("CREATE TABLE " & mydbase & "." & mytab & " LIKE maritime." & mytab & "")
            cd.MoveNext()
        End While

        cd.Close()
        Return True

    End Function
    Function centerform(theobject)

        theobject.Left = (Screen.PrimaryScreen.WorkingArea.Width - theobject.Width) / 2
        theobject.Top = (Screen.PrimaryScreen.WorkingArea.Height - theobject.Height) / 2
        Return theobject
    End Function

    Function Shuffle(ByRef sDeck() As String)

        Dim alRand As New ArrayList
        Dim iCount As Int32 = 0
        Dim iRand As Int32
        'MsgBox(sDeck.Length)

        'Do While iCount <= UBound(sDeck) - 1
        Do While iCount <= UBound(sDeck) - 1
            Dim rand As New Random
            iRand = rand.Next(0, 51)
            If alRand.Contains(iRand) = False Then
                alRand.Add(iRand)
                iCount += 1
            End If
        Loop

        Array.Sort(alRand.ToArray, sDeck)

        Return sDeck

    End Function

    Public Sub closeallchild()
        For Each frm As Form In DSEGMAIN.MdiChildren()
            frm.Close()
        Next
    End Sub


End Module
