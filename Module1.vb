Option Strict Off
Option Explicit On
Module Module1

    Public DiscID, Runout, DiscCondi As String
    Public CVPath, LVFile, CVFile, LVPath, PathNam, LVTempPath As String
    Public Prg01, Prg02, Prg03, Prg04, Prg05, Prg06, Prg07, Prg08, Prg09, Prg10 As String
    Public Prg11, Prg12, Prg13, Prg14, Prg15, Prg16, Prg17, Prg18, Prg19, Prg20 As String
    Public State, Usernam, Prg(0), Prg00(0), excelview, kopieview, MeasurStation, MessInstID As String
    Public DBOrExcel As Integer
    Public LVCVTimer, PrgCnt, MFCnt, MessCnt, txtCnt, intColIndex0, lngNumCols01, LVNewCnt, h As Integer
    Public Sheets02, Sheets01, Sheets03, tttemp, testnumber, ViewCVLV, msgMeasurend As String
    Public CVSheets01, CVSheets02, CVSheets03, LVSheets01, LVSheets02, LVSheets03 As String
    'Const xlDateiName As String = "Beispiel.xls"
    'Const xlWS_Name As String = "AdressDaten"
    Public PortMitutoyo, PortWeigh, selectmess As String
    Public xlWS_Name, xlDateiName, LastProj1, tempchr, tempstr0, tempstr1, StartLVTestNr, StartCVTestNr As String
    Public PadInside, Padcnt, Padoutside, canchelPad As Short
    Public PadsDir As String
    Public Disccnt, cntBack, startWriteRow, startWriteColum As Short
    Public strUsername As String
    Public cntStartRow, cntStartCol, valfailed As Short
    Public MitutoyoInstr1ID, MitutoyoInstr2ID, MitutoyoInstr3ID, MitutoyoInstr4ID As Integer
    Public weightID, weight1ID, xlwscnt As Integer
    Public Const Creatxt As String = "Create by H. Volk"
    Public Const Chr0 As String = " "
    Public Const Adminpwd As String = "FMO"
    Public Chr1, Chr2, Chr3, Chr4, Chr5 As String
    Public AdminUser As String
    Public DTVAnw, DTVPath As String

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    Public INIPath As String

    Public Property sPath() As String
        Get
            sPath = INIPath
        End Get
        Set(ByVal Value As String)
            INIPath = Value
        End Set
    End Property

    Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)
        WritePrivateProfileString(Section, Key, sValue, INIPath)
    End Sub

    Public Sub WriteValue(ByVal Section As String, ByVal Key As String, ByVal vValue As Object)
        WriteString(Section, Key, CStr(vValue))
    End Sub

    Public Function GetIniString(ByVal Section As String, ByVal Key As String, Optional ByVal Default_Renamed As String = "") As String

        Dim sTemp As String

        sTemp = New String(Chr(0), 256)
        GetPrivateProfileString(Section, Key, "", sTemp, Len(sTemp), INIPath)
        If InStr(sTemp, Chr(0)) Then
            sTemp = Left(sTemp, InStr(sTemp, vbNullChar) - 1)
        Else
            sTemp = Default_Renamed
        End If

        GetIniString = sTemp
    End Function

    Public Function GetIniLong(ByVal Section As String, ByVal Key As String, Optional ByVal Default_Renamed As Integer = -1) As Integer
        Dim sTemp As String

        sTemp = GetIniString(Section, Key, CStr(Default_Renamed))
        If IsNumeric(sTemp) Then
            GetIniLong = CShort(sTemp)
            'Else
            'Evtl. Fehlermeldung ausgeben
        End If
    End Function

    Public Function GetIniBool(ByVal Section As String, ByVal Key As String, Optional ByVal Default_Renamed As Boolean = False) As Boolean
        GetIniBool = CBool(GetIniLong(Section, Key, CShort(Default_Renamed)))
    End Function

    Sub IniTal()
        Dim ININame As String

        valfailed = 0
        DBOrExcel = 0
        ' 1 = CV ; 2 = LV
        'LVCVTimer = 0
        '##############################################################################
        '# Ini Datei auslesen;
        '# 
        '##############################################################################
        frmLogin.Combo1.Items.Clear()

        INIPath = My.Application.Info.DirectoryPath ' "C:\"
        ININame = "\Resources\config.ini"
        INIPath = INIPath & ININame

        '        LastProj1 = GetIniString("Path", "LVPath", INIPath) '***
        LastProj1 = GetIniString("Path", "FilePath", INIPath) '***
        LVPath = LastProj1

        LastProj1 = GetIniString("Path", "FilePath", INIPath) '***
        CVPath = LastProj1

        LVPath = CVPath

        LastProj1 = GetIniString("Path", "LVTempPath", INIPath) '***
        LVTempPath = LastProj1

        LastProj1 = GetIniString("Files", "CVFile", INIPath) '***
        CVFile = LastProj1

        LastProj1 = GetIniString("Files", "LVFile", INIPath) '***
        LVFile = LastProj1

        LastProj1 = GetIniString("Location", "State", INIPath) '***
        State = LastProj1

        'Public Chr1, Chr2, Chr3, Chr4, Chr5 As String

        LastProj1 = GetIniString("NotUsedChar", "Char1", INIPath) '***
        Chr1 = LastProj1

        LastProj1 = GetIniString("NotUsedChar", "Char2", INIPath) '***
        Chr2 = LastProj1

        LastProj1 = GetIniString("NotUsedChar", "Char3", INIPath) '***
        Chr3 = LastProj1

        LastProj1 = GetIniString("NotUsedChar", "Char4", INIPath) '***
        Chr4 = LastProj1

        LastProj1 = GetIniString("NotUsedChar", "Char5", INIPath) '***
        Chr5 = LastProj1

        ' [Config] excelview

        LastProj1 = GetIniString("Config", "excelview", INIPath) '***
        If LastProj1 = "True" Or LastProj1 = "False" Then
            excelview = LastProj1
        Else
            excelview = "False"
        End If

        LastProj1 = GetIniString("Config", "Kopieview", INIPath) '***
        If LastProj1 = "True" Or LastProj1 = "False" Then
            kopieview = LastProj1
        Else
            kopieview = "False"
        End If

        LastProj1 = GetIniString("Config", "DBOrExcel", INIPath) '***
        If LastProj1 = "Excel" Then
            DBOrExcel = 2
        ElseIf LastProj1 = "DB" Then
            DBOrExcel = 1
        Else
            DBOrExcel = 0
        End If

        LastProj1 = GetIniString("Config", "SelectCVandLV", INIPath) '***
        ViewCVLV = LastProj1

        'PadsDir

        LastProj1 = GetIniString("Config", "Padsdirector", INIPath) '***
        PadsDir = LastProj1

        '    Public AdminUser As String
        LastProj1 = GetIniString("Config", "AdminUser", INIPath) '***
        AdminUser = LastProj1

        ' DTV Values
        LastProj1 = GetIniString("Config", "DTVPath", INIPath) '***
        DTVPath = LastProj1

        LastProj1 = GetIniString("Config", "DTVExec", INIPath) '***
        DTVAnw = LastProj1

        'Public PortMitutoyo, PortWeigh As String

        LastProj1 = GetIniString("Config", "PORTMitutoyo", INIPath) '***
        PortMitutoyo = LastProj1

        LastProj1 = GetIniString("Config", "PORTWeigh", INIPath) '***
        PortWeigh = LastProj1

        'Public MitutoyoInstr1ID, MitutoyoInstr2ID, weightID As Integer

        LastProj1 = GetIniString("Config", "MitutoyoInstrumentNr1", INIPath) '***
        If LastProj1 <> "" Then
            MitutoyoInstr1ID = CInt(LastProj1)
        End If
        LastProj1 = GetIniString("Config", "MitutoyoInstrumentNr2", INIPath) '***
        If LastProj1 <> "" Then
            MitutoyoInstr2ID = CInt(LastProj1)
        End If

        LastProj1 = GetIniString("Config", "MitutoyoInstrumentNr3", INIPath) '***
        If LastProj1 <> "" Then
            MitutoyoInstr3ID = CInt(LastProj1)
        End If
        LastProj1 = GetIniString("Config", "MitutoyoInstrumentNr4", INIPath) '***
        If LastProj1 <> "" Then
            MitutoyoInstr4ID = CInt(LastProj1)
        End If

        LastProj1 = GetIniString("Config", "WaageInstrumentNr1", INIPath) '***
        If LastProj1 <> "" Then
            weightID = CInt(LastProj1)
        End If

        LastProj1 = GetIniString("Config", "WaageInstrumentNr2", INIPath) '***
        If LastProj1 <> "" Then
            weight1ID = CInt(LastProj1)
        End If

        'StartLVTestNr,StartCVTestNr As String
        LastProj1 = GetIniString("Config", "BeginLVTestNr", INIPath) '***
        StartLVTestNr = LastProj1

        LastProj1 = GetIniString("Config", "BeginCVTestNr", INIPath) '***
        StartCVTestNr = LastProj1

        LastProj1 = GetIniString("Sheets", "CVSheets01", INIPath) '***
        CVSheets01 = LastProj1

        LastProj1 = GetIniString("Sheets", "CVSheets02", INIPath) '***
        CVSheets02 = LastProj1

        LastProj1 = GetIniString("Sheets", "CVSheets03", INIPath) '***
        CVSheets03 = LastProj1

        LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
        LVSheets01 = LastProj1

        LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
        LVSheets02 = LastProj1

        LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
        LVSheets03 = LastProj1

    End Sub

    Sub ReadNewIniCancel()
        If LVCVTimer = 1 Then
            'Public Padcnt As Integer, PadInside As Integer, Padoutside As Integer
            'Public Disccnt As Integer
            ' 0 = Keine Auswahl, 1 = Auswahl
            If Disccnt > 0 And Padcnt < 1 Then
                LastProj1 = GetIniString("Sheets", "CVSheets01", INIPath) '***
                Sheets01 = LastProj1
            ElseIf Padcnt > 0 And PadInside > 0 Then
                LastProj1 = GetIniString("Sheets", "CVSheets02", INIPath) '***
                Sheets02 = LastProj1
            ElseIf Padcnt > 0 And Padoutside > 0 Then
                LastProj1 = GetIniString("Sheets", "CVSheets03", INIPath) '***
                Sheets03 = LastProj1
            End If
        ElseIf LVCVTimer = 2 And tttemp <> "" Then
            If MFCnt = 1 And PrgCnt >= 1 Then
                If Disccnt > 0 And Padcnt < 1 Then
                    LastProj1 = GetIniString(tttemp, "Disc", INIPath) '***
                    Sheets01 = LastProj1
                    LastProj1 = GetIniString(tttemp, "DiscMeasurePoints", INIPath) '***
                    MessCnt = LastProj1
                ElseIf Padcnt > 0 And PadInside > 0 Then
                    LastProj1 = GetIniString(tttemp, "PadInSide", INIPath) '***
                    Sheets02 = LastProj1
                    LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                    MessCnt = LastProj1
                ElseIf Padcnt > 0 And Padoutside > 0 Then
                    LastProj1 = GetIniString(tttemp, "PadOutSide", INIPath) '***
                    Sheets03 = LastProj1
                    LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                    MessCnt = LastProj1
                End If
            End If

        ElseIf LVCVTimer = 2 And tttemp = "" Then
            If Disccnt > 0 And Padcnt < 1 Then
                LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
                Sheets01 = LastProj1
            ElseIf Padcnt > 0 And PadInside > 0 Then
                LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
                Sheets02 = LastProj1
            ElseIf Padcnt > 0 And Padoutside > 0 Then
                LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
                Sheets03 = LastProj1
            End If
        End If

        ' 1 = CV ; 2 = LV
        If LVCVTimer = 1 Then
            readExcel()
        ElseIf LVCVTimer = 2 Then
            LVMircofaceExcel()
        End If
    End Sub

    Sub CheckcmdAuswahl()
        '* frmLogin.TopMost = False
        '*  Form2.Hide()
        '*  frmLogin.Hide()
        'Form2.Close()

        '* MsgBox("Messung beendet, Kein Messzeitpunkt mehr gefunden, Datei speichern !", MsgBoxStyle.Critical, "Messung beendet")
        '* frmLogin.SaveFileDialog1.FileName = xlDateiName '""
        '* frmLogin.SaveFileDialog1.ShowDialog()

        '* Form2.Close()
        '* frmLogin.Show()

        'Me.Close()
        'frmLogin.Show()


    End Sub
End Module
