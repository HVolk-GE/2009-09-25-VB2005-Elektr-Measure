Option Strict Off
Option Explicit On
Friend Class cIniFile

    ' =========================================================
	
	Private m_sPath As String
	Private m_sKey As String
	Private m_sSection As String
	Private m_sDefault As String
	Private m_lLastReturnCode As Integer
	
#If Win32 Then
    Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#Else
	' Profile String functions:
	Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If
	
	ReadOnly Property LastReturnCode() As Integer
		Get
			LastReturnCode = m_lLastReturnCode
		End Get
	End Property
	ReadOnly Property Success() As Boolean
		Get
			Success = (m_lLastReturnCode <> 0)
		End Get
	End Property

	Property Default_Renamed() As String
		Get
			Default_Renamed = m_sDefault
		End Get
		Set(ByVal Value As String)
			m_sDefault = Value
		End Set
	End Property

    Property Path() As String
        Get
            Path = m_sPath
        End Get
        Set(ByVal Value As String)
            m_sPath = Value
        End Set
    End Property

    Property Key() As String
        Get
            Key = m_sKey
        End Get
        Set(ByVal Value As String)
            m_sKey = Value
        End Set
    End Property

    Property Section() As String
        Get
            Section = m_sSection
        End Get
        Set(ByVal Value As String)
            m_sSection = Value
        End Set
    End Property

    Property Value() As String
        Get
            Dim sBuf As String
            Dim iSize As String
            Dim iRetCode As Short

            sBuf = Space(255)
            iSize = CStr(Len(sBuf))
            iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, CInt(iSize), m_sPath)
            If (CDbl(iSize) > 0) Then
                Value = Left(sBuf, iRetCode)
            Else
                Value = ""
            End If

        End Get

        Set(ByVal Value As String)
            Dim iPos As Short

            iPos = InStr(Value, Chr(0))
            Do While iPos <> 0
                Value = Left(Value, iPos - 1) & Mid(Value, iPos + 1)
                iPos = InStr(Value, Chr(0))
            Loop
            m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, Value, m_sPath)
        End Set
    End Property

    Property INISection() As String
        Get
            Dim sBuf As String
            Dim iSize As String
            Dim iRetCode As Short

            sBuf = Space(8192)
            iSize = CStr(Len(sBuf))
            iRetCode = GetPrivateProfileString(m_sSection, 0, m_sDefault, sBuf, CInt(iSize), m_sPath)
            If (CDbl(iSize) > 0) Then
                INISection = Left(sBuf, iRetCode)
            Else
                INISection = ""
            End If

        End Get
        Set(ByVal Value As String)
            m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0, Value, m_sPath)
        End Set
    End Property

    ReadOnly Property Sections() As String
        Get
            Dim sBuf As String
            Dim iSize As String
            Dim iRetCode As Short

            sBuf = Space(8192)
            iSize = CStr(Len(sBuf))
            iRetCode = GetPrivateProfileString(0, 0, m_sDefault, sBuf, CInt(iSize), m_sPath)
            If (CDbl(iSize) > 0) Then
                Sections = Left(sBuf, iRetCode)
            Else
                Sections = ""
            End If

        End Get
    End Property

    Public Sub DeleteKey()
        m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0, m_sPath)
    End Sub

    Public Sub DeleteSection()
        m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0, 0, m_sPath)
    End Sub

    Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef iCount As Integer)
        Dim sSection As String
        Dim iPos As Integer
        Dim iNextPos As Integer
        Dim sCur As String

        iCount = 0

        Erase sKey
        sSection = INISection
        If (Len(sSection) > 0) Then
            iPos = 1
            iNextPos = InStr(iPos, sSection, Chr(0))
            Do While iNextPos <> 0
                sCur = Mid(sSection, iPos, iNextPos - iPos)
                If (sCur <> Chr(0)) Then
                    iCount = iCount + 1

                    ReDim Preserve sKey(iCount)
                    sKey(iCount) = Mid(sSection, iPos, iNextPos - iPos)
                    iPos = iNextPos + 1
                    iNextPos = InStr(iPos, sSection, Chr(0))
                End If
            Loop
        End If
    End Sub

	Public Sub EnumerateAllSections(ByRef sSections() As String, ByRef iCount As Integer)
		Dim sIniFile As String
		Dim iPos As Integer
		Dim iNextPos As Integer
		Dim sCur As String
		
		iCount = 0
		Erase sSections
		sIniFile = Sections
		If (Len(sIniFile) > 0) Then
			iPos = 1
			iNextPos = InStr(iPos, sIniFile, Chr(0))
			Do While iNextPos <> 0
				If (iNextPos <> iPos) Then
					sCur = Mid(sIniFile, iPos, iNextPos - iPos)
					iCount = iCount + 1
					ReDim Preserve sSections(iCount)
					sSections(iCount) = sCur
				End If
				iPos = iNextPos + 1
				iNextPos = InStr(iPos, sIniFile, Chr(0))
			Loop 
		End If
		
    End Sub

	Public Sub SaveFormPosition(ByRef frmThis As Object)
		Dim sSaveKey As String
		Dim sSaveDefault As String
		On Error GoTo SaveError
		sSaveKey = Key

		If Not (frmThis.WindowState = System.Windows.Forms.FormWindowState.Minimized) Then
			Key = "Maximised"

			Value = CStr(CShort(frmThis.WindowState = System.Windows.Forms.FormWindowState.Maximized) * -1)

			If (frmThis.WindowState <> System.Windows.Forms.FormWindowState.Maximized) Then
				Key = "Left"

				Value = frmThis.Left
				Key = "Top"

				Value = frmThis.Top
				Key = "Width"

				Value = frmThis.Width
				Key = "Height"

				Value = frmThis.Height
			End If
		End If
		Key = sSaveKey
		Exit Sub
SaveError: 
		Key = sSaveKey
		m_lLastReturnCode = 0
		Exit Sub
	End Sub

    Public Sub LoadFormPosition(ByRef frmThis As Object, Optional ByRef lMinWidth As Object = 3000, Optional ByRef lMinHeight As Object = 3000)
        Dim sSaveKey As String
        Dim sSaveDefault As String
        Dim lLeft As Integer
        Dim lTOp As Integer
        Dim lWidth As Integer
        Dim lHeight As Integer
        On Error GoTo LoadError
        sSaveKey = Key
        sSaveDefault = Default_Renamed
        Default_Renamed = "FAIL"
        Key = "Left"
        lLeft = CLngDefault(Value, frmThis.Left)
        Key = "Top"
        lTOp = CLngDefault(Value, frmThis.Top)
        Key = "Width"
        lWidth = CLngDefault(Value, frmThis.Width)
        If (lWidth < lMinWidth) Then lWidth = lMinWidth
        Key = "Height"
        lHeight = CLngDefault(Value, frmThis.Height)
        If (lHeight < lMinHeight) Then lHeight = lMinHeight
        If (lLeft < 4 * VB6.TwipsPerPixelX) Then lLeft = 4 * VB6.TwipsPerPixelX
        If (lTOp < 4 * VB6.TwipsPerPixelY) Then lTOp = 4 * VB6.TwipsPerPixelY
        If (lLeft + lWidth > VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 4 * VB6.TwipsPerPixelX) Then
            lLeft = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 4 * VB6.TwipsPerPixelX - lWidth
            If (lLeft < 4 * VB6.TwipsPerPixelX) Then lLeft = 4 * VB6.TwipsPerPixelX
            If (lLeft + lWidth > VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - 4 * VB6.TwipsPerPixelX) Then
                lWidth = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - lLeft - 4 * VB6.TwipsPerPixelX
            End If
        End If
        If (lTOp + lHeight > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - 4 * VB6.TwipsPerPixelY) Then
            lTOp = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - 4 * VB6.TwipsPerPixelY - lHeight
            If (lTOp < 4 * VB6.TwipsPerPixelY) Then lTOp = 4 * VB6.TwipsPerPixelY
            If (lTOp + lHeight > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - 4 * VB6.TwipsPerPixelY) Then
                lHeight = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - lTOp - 4 * VB6.TwipsPerPixelY
            End If
        End If
        If (lWidth >= lMinWidth) And (lHeight >= lMinHeight) Then
            frmThis.Move(lLeft, lTOp, lWidth, lHeight)
        End If
        Key = "Maximised"
        If (CLngDefault(Value, 0) <> 0) Then
            frmThis.WindowState = System.Windows.Forms.FormWindowState.Maximized
        End If
        Key = sSaveKey
        Default_Renamed = sSaveDefault
        Exit Sub
LoadError:
        Key = sSaveKey
        Default_Renamed = sSaveDefault
        m_lLastReturnCode = 0
        Exit Sub
    End Sub

	Public Function CLngDefault(ByVal sString As String, Optional ByVal lDefault As Integer = 0) As Integer
		Dim lR As Integer
		On Error Resume Next
		lR = CInt(sString)
		If (Err.Number <> 0) Then
			CLngDefault = lDefault
		Else
			CLngDefault = lR
		End If
    End Function

End Class
