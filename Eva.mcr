Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "coordinate_systems.lib"
'#include "exports.lib"

Sub Main

' -------------------------------------------------------------------------------------------------
' Main: This function serves as a main program for testing purposes.
'       You need to rename this function to "Main" for debugging the result template.
'
'		PLEASE NOTE that a result template file must not contain a main program for
'       proper execution by the framework. Therefore please ensure to rename this function
'       to e.g. "Main2" before the result template can be used by the framework.
' -------------------------------------------------------------------------------------------------

	' Activate the StoreScriptSetting / GetScriptSetting functionality. Clear the data in order to
	' provide well defined environment for testing.

	ActivateScriptSettings True
	ClearScriptSettings

	' Now call the define method and check whether it is completed successfully

	If (Define("test", True, False)) Then

		' If the define method is executed properly, call the evaluate0d or 1d/c method
		Dim r1d As Object, stmpfile As String
		Dim ncount As Long, sName As String, sTableName As String, dvalue As Double

		stmpfile = "Test1D_temp.txt"
		ncount = 1

		Select Case GetScriptSetting("TemplateType", "1D")
			Case "0D"
				MsgBox CStr(Evaluate0D())
			Case "1D"
				Set r1d = Evaluate1D
				r1d.Save stmpfile
				r1d.AddToTree "1D Results\Test 1D"
				SelectTreeItem "1D Results\Test 1D"
			Case "1DC"
				Set r1d = Evaluate1DComplex
				r1d.Save stmpfile
				r1d.AddToTree "1D Results\Test 1DC"
				SelectTreeItem "1D Results\Test 1DC"
			Case "M0D"
				sName = "1D Results\Test M0D"
				dvalue = EvaluateMultiple0D(ncount, sName, sTableName)
				While (sTableName <> "")
					MsgBox(CStr(dvalue))
					ncount = ncount + 1
					dvalue = EvaluateMultiple0D(ncount, sName, sTableName)
				Wend
			Case "M1D"
				sName = "1D Results\Test M1D"
				Set r1d = EvaluateMultiple1D(ncount, sName, sTableName)
				While (sTableName <> "")
					r1d.Save stmpfile & CStr(ncount)
					r1d.AddToTree sTableName
					ncount = ncount + 1
					Set r1d = EvaluateMultiple1D(ncount, sName, sTableName)
				Wend
			Case "M1DC"
				sName = "1D Results\Test M1DC"
				Set r1d = EvaluateMultiple1DComplex(ncount, sName, sTableName)
				While (sTableName <> "")
					r1d.Save CSTr(ncount) & stmpfile
					r1d.AddToTree sTableName
					ncount = ncount + 1
					Set r1d = EvaluateMultiple1DComplex(ncount, sName, sTableName)
				Wend
		End Select


		With Resulttree
		    .UpdateTree
		    .RefreshView
		End With


	End If

	' Deactivate the StoreScriptSetting / GetScriptSetting functionality.

	ActivateScriptSettings False

End Sub

' *** global variables

Dim bDSTemplate As Boolean
Dim bInfoAlreadyShown As Boolean
Dim bIgnoreZcomp As Boolean

Dim acoordinates() As String

Dim sdir As String

Dim uvwlabel(3,2) As String
Dim iuvw As Integer

Dim a0DValue() As String

Dim aComponent() As String
Dim aComplex() As String

Dim aSolidArray_CST() As String, nSolids_CST As Integer
Dim iSolid_CST As Integer, bSolids_CST As Boolean

'------------------------------------------
	Dim b1Dplot_CST_arbitr As Boolean
	Dim cst_value_arbitr As Double
'------------------------------------------
	Dim x1box_CST_arbitr As Double, x2box_CST_arbitr As Double
	Dim y1box_CST_arbitr As Double, y2box_CST_arbitr As Double
	Dim z1box_CST_arbitr As Double, z2box_CST_arbitr As Double

	Dim dUVWvalue(3,3) As Double ' first index=u,v,w   second index=low,high,step
	Dim iDir_CST_arbitr As Integer

	Dim dVoxel_Unit As Double, dVoxel_SI As Double
	Dim bMultiplyRadiusLater As Boolean, bMultiplyRsinThetaLater As Boolean
	Dim dRef_Voxel_Unit As Double, dRef_Voxel_SI As Double, dRsinTheta As Double

	Dim icoordsystem_CST_arbitr As Integer, iWCS_CST_arbitr As Integer
	Dim bWriteFile_CST_arbitr As Boolean, iDataFileID As Integer

	Dim iNowSolid_CST_arbitr As Integer
	Dim dSolid_Integral_CST_arbitr() As Double
	Dim dSolid_Volume_CST_arbitr() As Double
	Dim nSolid_Data_CST_arbitr() As Long
	Dim dSolid_Maximum_CST_arbitr() As Double
	Dim dSolid_Minimum_CST_arbitr() As Double

	Dim bScalar_CST_arbitr As Boolean
	Dim sComponent_CST_arbitr As String
	Dim sComplex_CST_arbitr As String

	Dim bLokVA_CST_arbitr As Boolean, bLokVB_CST_arbitr As Boolean, bLokVC_CST_arbitr As Boolean
	Dim d_va_CST_arbitr(2) As Double, d_vb_CST_arbitr(2) As Double, d_vc_CST_arbitr(2) As Double

	Dim dSumVoxel_Unit_CST_arbitr As Double
	Dim dSumIntegral_CST_arbitr As Double
	Dim dMax_CST_arbitr As Double
	Dim dMin_CST_arbitr As Double
	Dim dMax_X_CST_arbitr As Double, dMax_Y_CST_arbitr As Double, dMax_Z_CST_arbitr As Double
	Dim dMin_X_CST_arbitr As Double, dMin_Y_CST_arbitr As Double, dMin_Z_CST_arbitr As Double

	' Arrays used to store x/y/z and potentially solid ID of all items added to calculation list
	Dim dListItemU() As Double, dListItemV() As Double, dListItemW() As Double
	Dim dListItemX() As Double, dListItemY() As Double, dListItemZ() As Double
	Dim iListItemSolidID() As Long, iListItemShapeID() As Long
'------------------------------------------

Dim bCheckMin As Boolean, bCheckMax As Boolean, dLimitMin As Double, dLimitMax As Double, bSkippedPointsDueToClamping As Boolean

Dim	sLogFilename_CST As String, bLogFileFirstEval As Boolean

Dim nCurrentLocale As Long ' the current locale at start time

Public Const sNameDefaultText = "Please enter result name or browse for it."
Public Const sWarningTemplateName_Local = "RTP-Evaluate in arbitrary coordinates: "
Public Const bDebugOutput = False

Public Const sAction0D = Array( _
	"Field Value")

Public Const sAction1D = Array( _
	"1D Plot of Field Values", _
	"Integral-1D", _
	"Integral f(x)-1D", _
	"Maximum-1D", _
	"Minimum-1D", _
	"Mean Value-1D", _
	"Deviation-1D", _
	"Length-1D")

Public Const sAction2D = Array( _
	"Integral-2D", _
	"Integral f(x)-2D", _
	"Maximum-2D", _
	"Minimum-2D", _
	"Mean Value-2D", _
	"Deviation-2D", _
	"Area-2D")

Public Const sAction3D = Array( _
	"Statistics-3D (Min/Max/Mean/Deviation...)", _
	"Integral-3D", _
	"Integral f(x)-3D", _
	"Volume-3D")

Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------

	Dim newname As String, sText As String

	' Log file needs to be active also if data file is requested
	If (DlgValue("WriteFile") = 1) Then
		DlgValue("GenerateLogfileCB", 1)
		DlgEnable("GenerateLogFileCB", False)
	Else
		DlgEnable("GenerateLogFileCB", True)
	End If

	If (Action% = 1 Or Action% = 2) Then

		DlgEnable "PushBrowseAll", 1

		If (DlgItem = "PushSpecials") Then
			DialogFunction = True       ' Don't close the dialog box.
			PushSpecials()
		End If

		If (DlgItem = "PushSetFrqTime") Then
			DialogFunction = True       ' Don't close the dialog box.
			sText=DlgText("TextFrqTime")
			PushSetFrqTime2D3D(sText)
			DlgText("TextFrqTime"),sText
			StoreScriptSetting("sTextFrqTime",sText)
		End If

		SetLabels DlgValue("edim"),DlgValue("coordsystem"),DlgValue("wcs"),DlgValue("coordinates")

		Dim iccc As Integer
		iccc = DlgValue("coordinates")
		DlgListBoxArray "coordinates", acoordinates
		DlgValue "coordinates", iccc

		DlgEnable("PushDatafile", (DlgValue("WriteFile")=1) Or (DlgValue("edim")<2))

		If (DlgItem = "edim") Then

			' if dimension has been changed, fill new actions into 0dvalue-list and select first one

			Dim n0DValue As Integer
			n0DValue = DlgValue("a0DValue")

			DlgListBoxArray "a0DValue", a0DValue

			If (n0DValue>UBound(a0DValue)) Then n0DValue=0
			DlgValue "a0DValue", n0DValue

		End If

		If (DlgItem = "Help") Then

			StartHelp "common_preloadedmacro_0d_evaluate_field_in_arbitrary_coordinates_(0d,_1d,_2d,_3d)"
			DialogFunction = True

		End If

		Dim s2exclude As String
		s2exclude = "Surface Current\"+cExcludeSeperator+"\Surface Current"

		If (DlgItem = "PushBrowseResults") Then
			DialogFunction = True       ' Don't close the dialog box.
			newname = ""
			PushBrowse2D3DResults(newname,s2exclude)
			If newname <> "" Then
				DlgText("ResultTreeNameArbitr"),newname
			End If
		End If

		If (DlgItem = "PushBrowseAll") Then
			DialogFunction = True       ' Don't close the dialog box.
			newname = ""
			PushBrowseAll2D3D(newname,s2exclude)
			If newname <> "" Then
				DlgText("ResultTreeNameArbitr"),newname
			End If
		End If

		' now scalar field detection and component adjustment

		If (bScalarField(DlgText("ResultTreeNameArbitr"))) Then

			FillArray aComponent() ,  Array("Scalar")
			DlgListBoxArray "aComponent", aComponent
			DlgValue "aComponent", 0
			DlgEnable "aComponent", 0

		Else
			If (DlgItem = "PushBrowseResults") Or (DlgItem = "PushBrowseAll") Or (DlgItem = "ResultTreeNameArbitr") Or (DlgItem = "edim") Or (DlgItem = "coordsystem") Or (DlgItem = "wcs") Then

				' if dimension has been changed, fill new components

				Dim iComp_Max As Integer
				Dim iComp_Last As Integer

				iComp_Max = UBound(aComponent)
				iComp_Last = DlgValue("aComponent")

				DlgListBoxArray "aComponent", aComponent

				If iComp_Last <= iComp_Max Then
					DlgValue "aComponent", iComp_Last
				Else
					DlgValue "aComponent", 0
				End If

			End If
			DlgEnable "aComponent", 1
		End If

		DlgText "u1text", uvwlabel(1,1)
		DlgText "v1text", uvwlabel(2,1)
		DlgText "w1text", uvwlabel(3,1)
		DlgText "u2text", uvwlabel(1,2)
		DlgText "v2text", uvwlabel(2,2)
		DlgText "w2text", uvwlabel(3,2)

		DlgText "idir", sdir

		DlgEnable "a0DValue", 1

		Select Case DlgValue("edim")

		Case 0 ' 0D
			DlgEnable "a0DValue", 0
			DlgEnable "coordinates", 0
			DlgEnable "idir", 0
			DlgEnable "Sampling", 0
			DlgEnable "stepsize", 0
       		DlgEnable "maxrange", 0
       		DlgEnable "u2", 0
       		DlgEnable "v2", 0
       		DlgEnable "w2", 0
       		DlgEnable "CheckSolids", 0
       		DlgEnable "BrowseSolids", 0

       		DlgEnable "u1", 1
       		DlgEnable "v1", 1
       		DlgEnable "w1", 1
       		DlgEnable "u2", 0
       		DlgEnable "v2", 0
       		DlgEnable "w2", 0
		Case 1 ' 1D
			DlgEnable "coordinates", 1
			DlgEnable "idir", 1
			DlgEnable "Sampling", 1
			DlgEnable "stepsize", 1
       		DlgEnable "maxrange", 1
       		DlgEnable "CheckSolids", 1
       		DlgEnable "BrowseSolids", IIf(DlgValue("CheckSolids") = 1, 1, 0)

       		DlgEnable "u1", 1
       		DlgEnable "v1", 1
       		DlgEnable "w1", 1
       		DlgEnable "u2", 0
       		DlgEnable "v2", 0
       		DlgEnable "w2", 0
       		If DlgValue("maxrange") = 0 Then
       			If DlgValue("coordinates") = 0 Then DlgEnable "u2", 1
       			If DlgValue("coordinates") = 1 Then DlgEnable "v2", 1
       			If DlgValue("coordinates") = 2 Then DlgEnable "w2", 1
       		Else
       			If DlgValue("coordinates") = 0 Then DlgEnable "u1", 0
       			If DlgValue("coordinates") = 1 Then DlgEnable "v1", 0
       			If DlgValue("coordinates") = 2 Then DlgEnable "w1", 0
       		End If
		Case 2 ' 2D
			DlgEnable "coordinates", 1
			DlgEnable "idir", 1
			DlgEnable "Sampling", 1
			DlgEnable "stepsize", 1
       		DlgEnable "maxrange", 1
       		DlgEnable "CheckSolids", 1
       		DlgEnable "BrowseSolids", IIf(DlgValue("CheckSolids") = 1, 1, 0)

       		If DlgValue("maxrange") = 0 Then
	       		DlgEnable "u1", 1
	       		DlgEnable "v1", 1
	       		DlgEnable "w1", 1
	       		DlgEnable "u2", 1
	       		DlgEnable "v2", 1
	       		DlgEnable "w2", 1
       			If DlgValue("coordinates") = 0 Then DlgEnable "u2", 0
       			If DlgValue("coordinates") = 1 Then DlgEnable "v2", 0
       			If DlgValue("coordinates") = 2 Then DlgEnable "w2", 0
       		Else
	       		DlgEnable "u1", 0
	       		DlgEnable "v1", 0
	       		DlgEnable "w1", 0
	       		DlgEnable "u2", 0
	       		DlgEnable "v2", 0
	       		DlgEnable "w2", 0
       			If DlgValue("coordinates") = 0 Then DlgEnable "u1", 1
       			If DlgValue("coordinates") = 1 Then DlgEnable "v1", 1
       			If DlgValue("coordinates") = 2 Then DlgEnable "w1", 1
       		End If
		Case 3 ' 3D
			DlgEnable "coordinates", 0
			DlgEnable "idir", 0
			DlgEnable "Sampling", 1
			DlgEnable "stepsize", 1
       		DlgEnable "maxrange", 1
       		DlgEnable "u2", 1
       		DlgEnable "v2", 1
       		DlgEnable "w2", 1
       		DlgEnable "CheckSolids", 1
       		DlgEnable "BrowseSolids", IIf(DlgValue("CheckSolids") = 1, 1, 0)
       		If DlgValue("maxrange") = 0 Then
	       		DlgEnable "u1", 1
	       		DlgEnable "v1", 1
	       		DlgEnable "w1", 1
	       		DlgEnable "u2", 1
	       		DlgEnable "v2", 1
	       		DlgEnable "w2", 1
       		Else
	       		DlgEnable "u1", 0
	       		DlgEnable "v1", 0
	       		DlgEnable "w1", 0
	       		DlgEnable "u2", 0
	       		DlgEnable "v2", 0
	       		DlgEnable "w2", 0
       		End If
		End Select

		Dim stmpfile As String
		If (DlgItem = "PushLogfile") Then
			DialogFunction = True       ' Don't close the dialog box.
			If (sLogFilename_CST <> "") Then
				stmpfile = GetProjectPath("Result") + sLogFilename_CST + ".log"
				If Dir$(stmpfile) <> "" Then
					Shell("notepad.exe " + stmpfile, 1)
				Else
					MsgBox "Option is only available after evaluation.",vbInformation
				End If
			Else
				MsgBox "Option is only available after evaluation.",vbInformation
			End If
		End If

		If (DlgItem = "PushDatafile") Then
			DialogFunction = True       ' Don't close the dialog box.
			If (sLogFilename_CST <> "") Then
				stmpfile = GetProjectPath("Result") + sLogFilename_CST + ".dat"
				If Dir$(stmpfile) <> "" Then
					Shell("notepad.exe " + stmpfile, 1)
				Else
					MsgBox "Option is only available after evaluation.",vbInformation
				End If
			Else
				MsgBox "Option is only available after evaluation.",vbInformation
			End If
		End If

		If (DlgItem = "DrawPoints") Then
			DialogFunction = True       ' Don't close the dialog box.

			' DrawPoints requires a mesh if automatic step size is selected
			If (Evaluate(DlgText("stepsize")) = 0) And (Mesh.GetNumberOfMeshCells  = 0) Then
				If MsgBox("A valid mesh is required to continue. Would you like to start the mesh generator now?", vbYesNo, "Mesh required") = vbYes Then
					Mesh.Update
				Else
					Exit Function
				End If
			End If

			DlgEnable("DrawPoints", False)
			If (sLogFilename_CST <> "") Then
				stmpfile = GetProjectPath("Result") + sLogFilename_CST + ".xyz"
				If Dir$(stmpfile) <> "" Then
					Select Case MsgBox("Found xyz data file. Would you like to: " & vbNewLine _
										& "Display existing data set (YES), " & vbNewLine _
										& "preview points and keep existing data (NO), " & vbNewLine _
										& "or preview points and delete existing data (CANCEL)?", vbYesNoCancel, "Display Data Points")
						Case vbYes
							DrawXYZPickPoints(stmpfile, 10000)
						Case vbNo
							StoreAllScriptSettings()
							PerformEvaluation(1, True) ' run in preview mode
							DrawXYZPickPoints(GetProjectPath("Result") + "points_preview.xyz", 1000)
						Case vbCancel
							Kill(stmpfile)
							StoreAllScriptSettings()
							PerformEvaluation(1, True) ' run in preview mode
							DrawXYZPickPoints(GetProjectPath("Result") + "points_preview.xyz", 1000)
					End Select
				Else
					StoreAllScriptSettings()
					PerformEvaluation(1, True) ' run in preview mode
					DrawXYZPickPoints(GetProjectPath("Result") + "points_preview.xyz", 1000)
				End If
			Else
				'MsgBox("Option is only available after evaluation.",vbInformation)
				' Create preview
				' If log file name exists, write data file with corresponding name.
				StoreAllScriptSettings()
				PerformEvaluation(1, True) ' run in preview mode
				DrawXYZPickPoints(GetProjectPath("Result") + "points_preview.xyz", 1000)
			End If
			DlgEnable("DrawPoints", True)
		End If

		If (DlgItem = "BrowseSolids") Then
			DialogFunction = True       ' Don't close the dialog box.
			SelectSolids aSolidArray_CST(), nSolids_CST
		End If

		If (DlgItem = "OK") Then
			' The user pressed the Ok button. Check the settings and display an error message if some required
			' fields have been left blank.
			StoreAllScriptSettings()
			'
			If (DlgText("ResultTreeNameArbitr")=sNameDefaultText) Then
				MsgBox "Field Result not specified." + vbCrLf + vbCrLf + sNameDefaultText, vbExclamation, "Field Result not specified."
				DialogFunction = True       ' There is an error in the settings -> Don't close the dialog box.
			ElseIf ((DlgValue("CheckSolids") = 1) And (nSolids_CST < 1)) Then
				MsgBox "Solid option is checked but no solids selected, please check your settings.", vbExclamation, "Solid Settings Check"
				DialogFunction = True       ' There is an error in the settings -> Don't close the dialog box.
			Else
				Dim iTFsweep As Integer
				iTFsweep = CInt(GetScriptSetting("GroupConstSweep",0))

				If ( (aComplex(DlgValue("aComplex"))="Complex") And (a0DValue(DlgValue("a0DValue"))<>"1D Plot of Field Values")) Then
					MsgBox "Complex results only supported with Result Value ""1D Plot of Field Values""" + vbCrLf + vbCrLf + "Please change your settings.", vbExclamation
					DialogFunction = True       ' There is an error in the settings -> Don't close the dialog box.
				Else
					If ( ((aComplex(DlgValue("aComplex"))="Average") Or (aComplex(DlgValue("aComplex"))="RMS")) And (aComponent(DlgValue("aComponent"))<>"Abs")) Then
						MsgBox """Average"" and ""RMS"" results only supported with Component: ""Abs""" + vbCrLf + vbCrLf + "Please change your settings.", vbExclamation
						DialogFunction = True       ' There is an error in the settings -> Don't close the dialog box.
					End If
				End If
			End If
			If CInt(GetScriptSetting("CheckUseFixedPointlist","0"))=1 Then
				' make sure that datafile is created, if special fixed pointlist is chosen in the specials dialogue
				StoreScriptSetting("WriteFile","1")
			End If
		End If

	End If
End Function

Sub SetLabels (idim As Integer, iCartPolar As Integer, iWCS As Integer, idir As Integer)

	Select Case iCartPolar

	Case 0 ' cartesian

		If iWCS = 0 Then
			FillArray acoordinates() ,  Array("X", "Y", "Z")
		Else
			FillArray acoordinates() ,  Array("U", "V", "W")
		End If

	Case 1 ' cylindrical

		If iWCS = 0 Then
			FillArray acoordinates() ,  Array("R", "F", "Z")
		Else
			FillArray acoordinates() ,  Array("R", "F", "W")
		End If

	Case 2 ' spherical

		' do not distinguish between global-spherical and local-spherical

		FillArray acoordinates() ,  Array("R", "Theta", "Phi")

	End Select

	sdir = "Normal:"

	Select Case idim

	Case 0 ' 0D
		FillArray a0DValue() ,  sAction0D
		FillArray aComponent() ,  Array("Abs")
		For iuvw=1 To 3
			uvwlabel(iuvw,1) = acoordinates(iuvw-1) + ":"
			uvwlabel(iuvw,2) = ""
		Next iuvw
	Case 1 ' 1D
		FillArray a0DValue() ,  sAction1D
		FillArray aComponent() ,  Array("Abs", "Tangential")
		sdir = "Direction:"
		For iuvw=1 To 3
			If iuvw = idir+1 Then
				uvwlabel(iuvw,1) = acoordinates(iuvw-1) + "min:"
				uvwlabel(iuvw,2) = acoordinates(iuvw-1) + "max:"
			Else
				uvwlabel(iuvw,1) = acoordinates(iuvw-1) + ":"
				uvwlabel(iuvw,2) = ""
			End If
		Next iuvw
	Case 2 ' 2D
		FillArray a0DValue() ,  sAction2D
		FillArray aComponent() ,  Array("Normal", "Abs")
		For iuvw=1 To 3
			If iuvw <> idir+1 Then
				uvwlabel(iuvw,1) = acoordinates(iuvw-1) + "min:"
				uvwlabel(iuvw,2) = acoordinates(iuvw-1) + "max:"
			Else
				uvwlabel(iuvw,1) = acoordinates(iuvw-1) + ":"
				uvwlabel(iuvw,2) = ""
			End If
		Next iuvw
	Case 3 ' 3D
		FillArray a0DValue() ,  sAction3D
		FillArray aComponent() ,  Array("Abs")
		For iuvw=1 To 3
			uvwlabel(iuvw,1) = acoordinates(iuvw-1) + "min:"
			uvwlabel(iuvw,2) = acoordinates(iuvw-1) + "max:"
		Next iuvw
	End Select

	Append2Array aComponent() ,  acoordinates()

	If (iCartPolar = 0) Then
		If (iWCS > 0) Then
			Append2Array aComponent() ,  Array("X", "Y", "Z")
		End If
	Else
		Append2Array aComponent() ,  Array("X", "Y", "Z")
		If (iWCS > 0) Then
			Append2Array aComponent() ,  Array("U", "V", "W")
		End If
	End If


End Sub
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean) As Boolean

	Dim sPreselectedField As String, sPreselectedFieldComponent As String, sPreselectedFieldArray() As String
	Dim i As Long

	StoreScriptSetting("bNameChanged", CStr(bNameChanged))
	bDSTemplate = Left(GetApplicationName,2)="DS"

	sWarningTemplateName = sWarningTemplateName_Local

	Dim sTextFrqTime As String
	sTextFrqTime = GetScriptSetting("sTextFrqTime","")

	FillArray aComplex(),  Array("Real Part", "Imag. Part", "Mag", "Phase", "Complex", "Average", "RMS")

	Dim adim() As String
	FillArray adim() ,  Array("0D", "1D", "2D", "3D")

	Dim acoordsystem() As String
	FillArray acoordsystem() ,  Array("cartesian", "cylindrical", "spherical")

	FillWCSArray

	' Initialize the global arrays first

	SetLabels 0,0,0,0

	Begin Dialog UserDialog 600,423,"Evaluate 0D/1D/2D/3D",.DialogFunction ' %GRID:3,3,1,1

		' *** Evaluated Field

		GroupBox 9,6,579,105,"Field Result",.GroupBox2

		TextBox 30,27,414,21,.ResultTreeNameArbitr
		PushButton 459,27,123,21,"Browse Results...",.PushBrowseResults
		PushButton 459,54,123,21,"Browse All...",.PushBrowseAll
		PushButton 321,54,123,21,"Set Frq / Time...",.PushSetFrqTime
		Text 321,81,123,15,sTextFrqTime,.TextFrqTime

		Text 30,57,81,15,"Component:",.Text13
		DropListBox 30,72,120,192,aComponent(),.aComponent
		Text 162,57,69,15,"Complex:",.Text14
		DropListBox 162,72,120,192,aComplex(),.aComplex

		' *** Calculation Range

		GroupBox 10,120,580,189,"Calculation Range",.GroupBox1
		'Text 40,238,50,14,"Dim:",.Text1
		'DropListBox 30,259,150,192,adim(),.edim

		Text 200,133,110,14,"Coord.System:",.Text2
		DropListBox 200,147,110,192,acoordsystem(),.coordsystem
		Text 200,174,40,14,"WCS:",.Text3
		DropListBox 200,189,110,192,aWCS(),.wcs
		Text 200,216,70,14,sdir,.idir
		DropListBox 200,231,110,192,acoordinates(),.coordinates
		Text 200,258,120,14,"Stepsize (0=auto):",.Sampling
		TextBox 200,273,110,21,.stepsize

		CheckBox 350,147,100,14,"max. range",.maxrange

		Text 350,174,90,14,uvwlabel(1,1),.u1text
		Text 350,216,100,14,uvwlabel(2,1),.v1text
		Text 350,258,100,14,uvwlabel(3,1),.w1text
		Text 470,174,90,14,uvwlabel(1,2),.u2text
		Text 470,216,100,14,uvwlabel(2,2),.v2text
		Text 471,258,108,15,uvwlabel(3,2),.w2text

		TextBox 350,189,100,21,.u1
		TextBox 470,189,100,21,.u2
		TextBox 350,231,100,21,.v1
		TextBox 470,231,100,21,.v2
		TextBox 350,273,100,21,.w1

		' *** Result Value

		GroupBox 9,315,579,69,"Result Value",.GroupBox3
		TextBox 470,273,100,21,.w2

		CheckBox 30,255,111,15,"Select Solids",.CheckSolids

		DropListBox 30,336,280,192,a0DValue(),.a0DValue
		CheckBox 327,339,249,15,"Log file (decreases performance)",.GenerateLogfileCB
		CheckBox 327,363,243,15,"Data file (decreases performance)",.WriteFile
		PushButton 48,273,90,21,"Solids...",.BrowseSolids

		OKButton 10,396,90,21
		CancelButton 105,396,90,21
		PushButton 200,396,90,21,"Help",.Help
		PushButton 310,396,90,21,"DrawPoints",.DrawPoints
		PushButton 405,396,90,21,"Datafile...",.PushDatafile
		PushButton 500,396,90,21,"Logfile...",.PushLogfile
		OptionGroup .edim
			OptionButton 30,147,110,14,"0D = Point",.OptionButton1
			OptionButton 30,168,130,14,"1D = Curve/Line",.OptionButton2
			OptionButton 30,189,140,14,"2D = Area/Face",.OptionButton3
			OptionButton 30,210,130,14,"3D = Volume",.OptionButton4
		PushButton 27,360,90,21,"Specials...",.PushSpecials

	End Dialog
	Dim dlg As UserDialog

	' Pre-enter selected tree item, if item resides in 2D/3D results
	sPreselectedField = GetSelectedTreeItem
	sPreselectedFieldArray = Split(sPreselectedField, "\")
	If ((UBound(sPreselectedFieldArray)>0) _
			And (Left(sPreselectedField, 13) = "2D/3D Results") _
			And (sPreselectedField <> "2D/3D Results\E-Field") _
			And (sPreselectedField <> "2D/3D Results\H-Field") _
			And (sPreselectedField <> "2D/3D Results\Surface Current")) Then' Extract field path and component; exclude some specific folders
		' Remove "2D/3D Results" at beginning
		For i = 1 To UBound(sPreselectedFieldArray)
			sPreselectedFieldArray(i-1) = sPreselectedFieldArray(i)
		Next
		ReDim Preserve sPreselectedFieldArray(UBound(sPreselectedFieldArray)-1)
		' Identify component at end, if present. Also truncate array accordingly
		Select Case sPreselectedFieldArray(UBound(sPreselectedFieldArray))
			Case "X", "Y", "Z", "Abs"
				sPreselectedFieldComponent = sPreselectedFieldArray(UBound(sPreselectedFieldArray))
				ReDim Preserve sPreselectedFieldArray(UBound(sPreselectedFieldArray)-1)
			Case "Normal", "Tangential"
				' ignore component and use default, but strip it from tree entry
				sPreselectedFieldComponent = aComponent(0)
				ReDim Preserve sPreselectedFieldArray(UBound(sPreselectedFieldArray)-1)
			Case Else
				' ignore component and use default
				sPreselectedFieldComponent = aComponent(0)
		End Select
		sPreselectedField = Join(sPreselectedFieldArray, "\")
	Else
		sPreselectedField = sNameDefaultText
		sPreselectedFieldComponent = aComponent(0)
	End If
	dlg.ResultTreeNameArbitr = GetScriptSetting("ResultTreeNameArbitr",sPreselectedField)

	dlg.WriteFile 	= CInt(GetScriptSetting("WriteFile","0"))

	dlg.edim 		= CInt(GetScriptSetting("edim","1"))
	dlg.coordsystem = CInt(GetScriptSetting("coordsystem","0"))
	dlg.wcs 		= CInt(GetScriptSetting("wcs","0"))
	dlg.coordinates	= CInt(GetScriptSetting("coordinates","0"))
	dlg.stepsize	= GetScriptSetting("stepsize","0.0")
	dlg.maxrange	= CInt(GetScriptSetting("maxrange","1"))

	Dim npicks As Long
	Dim du1 As Double, dv1 As Double, dw1 As Double
	Dim du2 As Double, dv2 As Double, dw2 As Double
	Dim su1 As String, sv1 As String, sw1 As String
	Dim su2 As String, sv2 As String, sw2 As String

	Boundary.GetCalculationBox du1, du2, dv1, dv2, dw1, dw2
	su1 = CStr(du1)
	sv1 = CStr(dv1)
	sw1 = CStr(dw1)
	su2 = CStr(du2)
	sv2 = CStr(dv2)
	sw2 = CStr(dw2)
	npicks = Pick.GetNumberOfPickedPoints
	If npicks=1 Or npicks=2 Then
		Pick.GetPickpointCoordinates (1, du1, dv1, dw1)
		su1 = cstr(du1)
		sv1 = cstr(dv1)
		sw1 = cstr(dw1)
	End If
	If npicks=2 Then
		Pick.GetPickpointCoordinates (2, du2, dv2, dw2)
		su2 = cstr(du2)
		sv2 = cstr(dv2)
		sw2 = cstr(dw2)
		If du1>du2 Then
			su1 = cstr(du2)
			su2 = cstr(du1)
		End If
		If dv1>dv2 Then
			sv1 = cstr(dv2)
			sv2 = cstr(dv1)
		End If
		If dw1>dw2 Then
			sw1 = cstr(dw2)
			sw2 = cstr(dw1)
		End If
	End If

	dlg.u1 = GetScriptSetting("u1",su1)
	dlg.v1 = GetScriptSetting("v1",sv1)
	dlg.w1 = GetScriptSetting("w1",sw1)
	dlg.u2 = GetScriptSetting("u2",su2)
	dlg.v2 = GetScriptSetting("v2",sv2)
	dlg.w2 = GetScriptSetting("w2",sw2)

	SetLabels dlg.edim,dlg.coordsystem,dlg.wcs,dlg.coordinates

	dlg.a0DValue=FindListIndex(a0DValue(), GetScriptSetting("a0DValue",a0DValue(0)))
	dlg.GenerateLogfileCB = CInt(GetScriptSetting("GenerateLogfile", "1"))

	dlg.aComponent	= FindListIndex(aComponent(), GetScriptSetting("aComponent",sPreselectedFieldComponent))
	dlg.aComplex	= FindListIndex(aComplex(), GetScriptSetting("aComplex",aComplex(0)))

	' read solid information

	dlg.CheckSolids = CInt(GetScriptSetting("bSolids","0"))
	nSolids_CST = CInt(GetScriptSetting("nSolids","0"))

	If (nSolids_CST > 0) Then
		ReDim aSolidArray_CST(nSolids_CST-1)

		For iSolid_CST = 1 To nSolids_CST
			aSolidArray_CST(iSolid_CST-1) = GetScriptSetting("Solid" + CStr(iSolid_CST),"")
		Next
	End If

	sLogFilename_CST = GetScriptSetting("sLogFilename_CST","")

	If (Not Dialog(dlg, -1)) Then

		' The user left the dialog box without pressing Ok. Assigning False to the function
		' will cause the framework to cancel the creation or modification without storing
		' anything.

		Define = False
	Else

		' The user properly left the dialog box by pressing Ok. Assigning True to the function
		' will cause the framework to complete the creation or modification and store the corresponding
		' settings.

		If Not bNameChanged Then sName = GetScriptSetting("sName", "Evaluate fields in arbitrary coordinates")
		Define = True

		' Store the script settings into the database for later reuse by either the define function (for modifications)
		' or the evaluate function.

	End If

End Function

Function StoreAllScriptSettings() As Integer

		' Store all dialog settings as script settings
		' Most variables used are global and do not need to be dim'ed

		StoreScriptSetting("WriteFile",CStr(DlgValue("WriteFile")))
		StoreScriptSetting("edim",CStr(DlgValue("edim")))
		StoreScriptSetting("coordsystem",CStr(DlgValue("coordsystem")))
		StoreScriptSetting("wcs",CStr(DlgValue("wcs")))
		StoreScriptSetting("coordinates",CStr(DlgValue("coordinates")))
		StoreScriptSetting("stepsize",DlgText("stepsize"))
		StoreScriptSetting("maxrange",CStr(DlgValue("maxrange")))

		StoreScriptSetting("u1",DlgText("u1"))
		StoreScriptSetting("v1",DlgText("v1"))
		StoreScriptSetting("w1",DlgText("w1"))
		StoreScriptSetting("u2",DlgText("u2"))
		StoreScriptSetting("v2",DlgText("v2"))
		StoreScriptSetting("w2",DlgText("w2"))

		StoreScriptSetting("a0DValue",a0DValue(DlgValue("a0DValue")))
		StoreScriptSetting("GenerateLogfile", DlgValue("GenerateLogfileCB"))

		StoreScriptSetting("ResultTreeNameArbitr",DlgText("ResultTreeNameArbitr"))
		StoreScriptSetting("aComponent",aComponent(DlgValue("aComponent")))
		StoreScriptSetting("aComplex",aComplex(DlgValue("aComplex")))

		StoreScriptSetting("aCoordsystem",DlgText("coordsystem"))
		StoreScriptSetting("coord1",acoordinates(0))
		StoreScriptSetting("coord2",acoordinates(1))
		StoreScriptSetting("coord3",acoordinates(2))

		StoreScriptSetting("VaryingCoordinate", CStr(acoordinates(DlgValue("coordinates"))))

		' write solid information

		StoreScriptSetting("bSolids", CStr(DlgValue("CheckSolids")))
		StoreScriptSetting("nSolids", CStr(nSolids_CST))

		For iSolid_CST = 1 To nSolids_CST
			StoreScriptSetting("Solid" + CStr(iSolid_CST),aSolidArray_CST(iSolid_CST-1))
		Next iSolid_CST

		b1Dplot_CST_arbitr = (GetScriptSetting("a0DValue","") = "1D Plot of Field Values")

		If b1Dplot_CST_arbitr Then
			If aComplex(DlgValue("aComplex")) = "Complex" Then
				StoreTemplateSetting("TemplateType", "1DC")
			Else
				StoreTemplateSetting("TemplateType", "1D")
			End If
			' Add "M" in front if all frequencies are swept
			If (CInt(GetScriptSetting("CheckBoxFrqTimeActive",0))=1) Then
				Select Case CInt(GetScriptSetting("GroupConstSweep",0))
					Case 1,4 ' all frequencies or all times
						StoreTemplateSetting("TemplateType", "M"+GetScriptSetting("TemplateType", "1D"))
					Case Else
						' Do nothing
						' ReportWarning("Evaluate field in arbitrary coordinates: The selected frequency or time sweep settings are currently not supported for option 'Statistics-3D'. Please contact support.")
				End Select

			End If
		ElseIf (GetScriptSetting("a0DValue","") = "Statistics-3D (Min/Max/Mean/Deviation...)") Then
			StoreTemplateSetting("TemplateType", "M0D")
			If (CInt(GetScriptSetting("CheckBoxFrqTimeActive",0))=1) Then
				Select Case CInt(GetScriptSetting("GroupConstSweep",0))
					Case 1,4 ' all frequency or time samples
						StoreTemplateSetting("TemplateType", "M1D")
					Case Else
						' Do nothing
						' ReportWarning("Evaluate field in arbitrary coordinates: The selected frequency or time sweep settings are currently not supported for option 'Statistics-3D'. Please contact support.")
				End Select
			End If
		Else
			Dim iTFsweep As Integer
			iTFsweep = CInt(GetScriptSetting("GroupConstSweep",0))

			If CInt(GetScriptSetting("CheckBoxFrqTimeActive",0))=1 And iTFsweep<>0 And iTFsweep<>3 Then
				StoreTemplateSetting("TemplateType","1D")
			Else
				StoreTemplateSetting("TemplateType","0D")
			End If
		End If

	    Dim sFieldtmp As String, sName As String
		sFieldtmp = DlgText("ResultTreeNameArbitr")

		If (Not CBool(GetScriptSetting("bNameChanged", "0"))) Then

		    sName = Mid$(sFieldtmp,1+InStrRev(sFieldtmp,"\"))

			If Not bScalarField(sFieldtmp) Then
		    	sName = sName + "_" + aComponent(DlgValue("aComponent"))
		    End If

			If b1Dplot_CST_arbitr Then
			    sName = sName + " (" + CStr(acoordinates(DlgValue("coordinates"))) + ")"
			Else
			    sName = sName + "_" + CStr(DlgValue("edim")) + "D"
			End If

		    sName = NoForbiddenFilenameCharacters(sName)

		End If
		StoreScriptSetting("sName", sName)

End Function


Private Function DialogFunctionPushSpecials(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------

	If (Action%=1) Or (Action%=2) Then

       	DlgEnable "FilePointlist", IIf(DlgValue("CheckUseFixedPointlist") = 1, 1, 0)

		If (DlgItem = "OK") Then
		    ' The user pressed the Ok button. Check the settings and display an error message if necessary
			If (False) Then
				MsgBox("Please correct your settings.", "Input Check")
				DialogFunctionPushSpecials = True
			End If
		End If

	End If
End Function
Sub PushSpecials()

	Begin Dialog UserDialog 460,285,"Specials",.DialogFunctionPushSpecials ' %GRID:10,5,1,1
		GroupBox 10,5,440,65,"'Integral f(x)' Result values only:",.GroupBox1
		OKButton 20,260,90,20
		CancelButton 120,260,90,20
		Text 30,25,40,15,"f(x)=",.Text1
		TextBox 70,23,360,20,.Integ_fx
		Text 80,48,240,15,"(x represents field value)",.Text2
		CheckBox 30,80,380,15,"Ignore z-Component of field vector (will be set to 0)",.bIgnoreZcomp
		GroupBox 10,110,440,75,"Clamp result range",.GroupBox2
		CheckBox 30,132,270,15,"Only consider field values larger than",.bCheckMin
		TextBox 310,130,130,20,.dLimitMin
		CheckBox 30,158,270,15,"Only consider field values smaller than",.bCheckMax
		TextBox 310,155,130,20,.dLimitMax
		GroupBox 10,190,440,60,"",.GroupBox3
		CheckBox 20,205,260,15,"Use fixed FileName for xyz-pointlist:",.CheckUseFixedPointlist
		Text 50,230,250,15,"(stored in Projectdir\Export\3d\...)",.Text3
		TextBox 280,205,150,20,.FilePointlist

	End Dialog
	Dim dlg As UserDialog

	dlg.Integ_fx		= GetScriptSetting("Integ_fx","x")
	dlg.bIgnoreZcomp 	= CInt(GetScriptSetting("bIgnoreZcomp","0"))
	dlg.bCheckMin 		= CInt(GetScriptSetting("bCheckMin","0"))
	dlg.bCheckMax 		= CInt(GetScriptSetting("bCheckMax","0"))
	dlg.dLimitMin		= GetScriptSetting("dLimitMin","0.0")
	dlg.dLimitMax		= GetScriptSetting("dLimitMax","0.0")
	dlg.CheckUseFixedPointlist 	= CInt(GetScriptSetting("CheckUseFixedPointlist","0"))
	dlg.FilePointlist		= GetScriptSetting("FilePointlist","xyz-pointlist.txt")

	If (Dialog(dlg) <> 0) Then
		StoreScriptSetting("Integ_fx",CStr(dlg.Integ_fx))
		StoreScriptSetting("bIgnoreZcomp",CStr(dlg.bIgnoreZcomp))
		StoreScriptSetting("bCheckMin",CStr(dlg.bCheckMin))
		StoreScriptSetting("bCheckMax",CStr(dlg.bCheckMax))
		StoreScriptSetting("dLimitMin",CStr(dlg.dLimitMin))
		StoreScriptSetting("dLimitMax",CStr(dlg.dLimitMax))
		StoreScriptSetting("CheckUseFixedPointlist",CStr(dlg.CheckUseFixedPointlist))
		StoreScriptSetting("FilePointlist",CStr(dlg.FilePointlist))
	End If

End Sub

Function ReplaceVariables(ByRef sFieldCST As String) As String

	Dim nParameters As Long, i As Long, sParameterName As String, dParameterValue As Double

	' Replace $#VARNAME#$ with Evaluate(VARNAME) in sFieldCST to allow manual parameterization
	If (InStr(sFieldCST, "$#") < InStr(sFieldCST, "#$")) Then
		nParameters = IIf(Left(GetApplicationName, 2) = "DS", DS.GetNumberOfParameters, GetNumberOfParameters)
		For i = 0 To nParameters-1
			sParameterName = IIf(Left(GetApplicationName, 2) = "DS", DS.GetParameterName(i), GetParameterName(i))
			' DS and MWS parameters not synched during parameter sweep, need to manually distinguish here
			dParameterValue = IIf(Left(GetApplicationName, 2) = "DS", DS.RestoreDoubleParameter(sParameterName), RestoreDoubleParameter(sParameterName))
			sFieldCST = Replace(sFieldCST, "$#"+sParameterName+"#$", CStr(dParameterValue))
			If Not (InStr(sFieldCST, "$#") < InStr(sFieldCST, "#$")) Then Exit For
		Next
	End If

	ReplaceVariables = sFieldCST

End Function

Function FillPointsInList() As Long

	Dim bTake_this_point As Boolean
	Dim sTempName As String

	Dim vxam As Double, vxph As Double
	Dim vyam As Double, vyph As Double
	Dim vzam As Double, vzph As Double

	Dim vxTmp As Double, vyTmp As Double, vzTmp As Double, vNowCST As Double
	Dim vxTmpim As Double, vyTmpim As Double, vzTmpim As Double, vNowCSTim As Double

	Dim vTmpReal As Double, vTmpImag As Double

	If CInt(GetScriptSetting("bIgnoreZcomp","0")) = 0 Then
		bIgnoreZcomp = False
	Else
		bIgnoreZcomp = True
	End If

	Dim uvw_CST(2) As Double, bbb(2) As Double, xyz_CST(2) As Double
	Dim u_CST As Double, v_CST As Double, w_CST As Double
	Dim x_CST As Double, y_CST As Double, z_CST As Double

	Dim i As Long, j As Long
	Dim u As Long, v As Long, w As Long
	Dim nXSamples As Long, nYSamples As Long, nZSamples As Long, nMaxPoints As Long

	Dim dStartTime As Double
	dStartTime = Timer()

	VectorPlot3D.Reset

	' prevent that low=high and step=0
	Dim iyy As Long
	For iyy = 1 To 3
		If ( dUVWvalue(iyy,1)=dUVWvalue(iyy,2) ) And (dUVWvalue(iyy,3)=0.0) Then
			dUVWvalue(iyy,3) = 1.0
		End If
	Next

	FillPointsInList = 0

	' Estimate max number of points and give a warning if number is very high
	' dUVWvalues HAVE to be > 0 or the loop below would run endlessly
	nXSamples = Fix((dUVWvalue(1,2)-dUVWvalue(1,1))/dUVWvalue(1,3)+1)
	nYSamples = Fix((dUVWvalue(2,2)-dUVWvalue(2,1))/dUVWvalue(2,3)+1)
	nZSamples = Fix((dUVWvalue(3,2)-dUVWvalue(3,1))/dUVWvalue(3,3)+1)

	nMaxPoints = nXSamples*nYSamples*nZSamples
	If ((nMaxPoints >= 1e6) And (Not bSolids_CST)) _
		Or ((nMaxPoints >=5e6) And (bSolids_CST)) Then
		ReportWarning("Evaluate field in arbitrary coordinates: The number of samples is very large ("+CStr(nMaxPoints)+"), evaluation might take a long time.")
	End If

	' Allocate memory for iArraySizeIncrement list items initially, then increase later if needed.
	' Allocating max number from beginning is slow, so is redim with every step --> compromise
	ReDim dListItemX(nMaxPoints - 1)
	ReDim dListItemY(nMaxPoints - 1)
	ReDim dListItemZ(nMaxPoints - 1)
	ReDim dListItemU(nMaxPoints - 1)
	ReDim dListItemV(nMaxPoints - 1)
	ReDim dListItemW(nMaxPoints - 1)

	If bDebugOutput Then ReportInformationToWindow("FillPointsInList: Start loop after " + CStr(Timer()-dStartTime))

	' If(For) is more code than For(If), but faster; also tried a single loop instead 3 nested loops, but that was not faster...
	For u = 0 To nXSamples-1
		u_CST = dUVWvalue(1,1) + dUVWvalue(1,3) * u
		uvw_CST(0) = u_CST
		For v = 0 To nYSamples-1
			v_CST = dUVWvalue(2,1) + dUVWvalue(2,3) * v
			If (GetTemplateAborted) Then
				FillPointsInList = -1
				Exit Function
			End If
			
			uvw_CST(1) = v_CST
			For w = 0 To nZSamples-1
				w_CST = dUVWvalue(3,1) + dUVWvalue(3,3) * w
				uvw_CST(2) = w_CST
				' if not cartesian : transfer uvw into cartesian xyz_CST
				If (iWCS_CST_arbitr = 0 ) Then
					' global xyz
					Select Case icoordsystem_CST_arbitr
						Case 0 ' cartesian   xyz
							xyz_CST(0) = u_CST
							xyz_CST(1) = v_CST
							xyz_CST(2) = w_CST
						Case 1 ' cylindrical  rFz
							Convert_Point_Cylindrical2Cartesian uvw_CST, xyz_CST
						Case 2 ' spherical   rTP
							Convert_Point_Spherical2Cartesian uvw_CST, xyz_CST
					End Select
				Else
					' local uvw
					Select Case icoordsystem_CST_arbitr
						Case 0 ' cartesian   xyz
							Convert_Point2Global uvw_CST, xyz_CST
						Case 1 ' cylindrical  rFz
							Convert_Point_Cylindrical2Cartesian uvw_CST, bbb
							Convert_Point2Global bbb, xyz_CST
						Case 2 ' spherical   rTP
							Convert_Point_Spherical2Cartesian uvw_CST, bbb
							Convert_Point2Global bbb, xyz_CST
					End Select
				End If
				x_CST = xyz_CST(0)
				y_CST = xyz_CST(1)
				z_CST = xyz_CST(2)
				FillPointsInList = FillPointsInList + 1
				dListItemX(FillPointsInList-1) = x_CST
				dListItemY(FillPointsInList-1) = y_CST
				dListItemZ(FillPointsInList-1) = z_CST
				dListItemU(FillPointsInList-1) = u_CST
				dListItemV(FillPointsInList-1) = v_CST
				dListItemW(FillPointsInList-1) = w_CST
			Next w
		Next v
	Next u

	If bDebugOutput Then ReportInformationToWindow("FillPointsInList: Finish loop after " + CStr(Timer()-dStartTime))

	VectorPlot3D.SetPoints(dListItemX, dListItemY, dListItemZ)

	If bSolids_CST Then
		FillPointsInList = 0
		iNowSolid_CST_arbitr = 1

		' Go through full list and determine ID of shape in which points are located
		VectorPlot3D.DetermineShapeIDs
		iListItemShapeID = VectorPlot3D.GetShapeList
		ReDim iListItemSolidID(UBound(iListItemShapeID))
		Solid.GenerateShapeIDTable

		If bDebugOutput Then ReportInformationToWindow("FillPointsInList: Shape ID setup completed after " + CStr(Timer()-dStartTime))

		For i = 0 To UBound(iListItemShapeID)
			If (i Mod 1000 = 0) Then
				If (GetTemplateAborted) Then
					FillPointsInList = -1
					Exit Function
				End If
			End If
			bTake_this_point = False
			' Check last found solid first. Good chance that neighboring points are in same solid
			sTempName = Mid(Solid.GetShapeNameFromID(iListItemShapeID(i)), 12) ' remove the first 11 letters, typically 'Components\' in front
			If (sTempName = "") Then
				' Do nothing, point is in background or a face port with a single digit number.
			ElseIf ( sTempName = aSolidArray_CST(iNowSolid_CST_arbitr-1) ) Then
				bTake_this_point = True
			Else
				For iSolid_CST = 1 To nSolids_CST
					If ( sTempName = aSolidArray_CST(iSolid_CST-1) ) Then
						bTake_this_point = True
						iNowSolid_CST_arbitr = iSolid_CST
						Exit For
					End If
				Next
			End If

			If bTake_this_point Then
				' ReportInformationToWindow("Point found!")
				dListItemX(FillPointsInList) = dListItemX(i)
				dListItemY(FillPointsInList) = dListItemY(i)
				dListItemZ(FillPointsInList) = dListItemZ(i)
				dListItemU(FillPointsInList) = dListItemU(i)
				dListItemV(FillPointsInList) = dListItemV(i)
				dListItemW(FillPointsInList) = dListItemW(i)
				iListItemSolidID(FillPointsInList) = iNowSolid_CST_arbitr
				FillPointsInList = FillPointsInList + 1
			End If

		Next

		If FillPointsInList > 0 Then
			' ReDim to final size
			ReDim Preserve dListItemX(FillPointsInList - 1)
			ReDim Preserve dListItemY(FillPointsInList - 1)
			ReDim Preserve dListItemZ(FillPointsInList - 1)
			ReDim Preserve dListItemU(FillPointsInList - 1)
			ReDim Preserve dListItemV(FillPointsInList - 1)
			ReDim Preserve dListItemW(FillPointsInList - 1)
			ReDim Preserve iListItemSolidID(FillPointsInList - 1)
		Else
			Exit Function
		End If

		' Update point list for calculation
		VectorPlot3D.Reset
		VectorPlot3D.SetPoints(dListItemX, dListItemY, dListItemZ)

		If bDebugOutput Then ReportInformationToWindow("FillPointsInList: Points in Solids identified: " + CStr(Timer()-dStartTime))

	End If

	If bDebugOutput Then ReportInformationToWindow("Number of points: " & CStr(FillPointsInList))
	If bDebugOutput Then ReportInformationToWindow("FillPointsInList: Set points after " + CStr(Timer()-dStartTime))

End Function

Function WriteXYZFile(sPointxyzFile_CST As String, nNumberOfPoints As Long) As Integer

	' Write the XYZ coordinates of the points in list to a file; dListItemX/Y/Z are global variables
	' This function uses its own buffer as the buffered file writer in exports.lib only allows one output at a time, which is already used for the dat file in this template

	Dim i As Long, sOutputFormat As String, iFileID As Long, sBuffer As String

	sOutputFormat = "  0.0000E+00; -0.0000E+00"
	iFileID = FreeFile()
	sBuffer = ""

	SetLocale(&H409)
	Open sPointxyzFile_CST For Output As #iFileID

	For i = 0 To nNumberOfPoints-1
		' This can be extremely slow on network drives - allow abort before writing buffer
		' Old version using vba_globals_all:PP12(...) function was slow; PP12 can also handle non-numerical values but is not needed here since values are known to be numerical
		sBuffer = sBuffer & Format(dListItemX(i), sOutputFormat) & Format(dListItemY(i), sOutputFormat) & Format(dListItemZ(i), sOutputFormat) & vbNewLine
		If (i + 1 Mod 2500 = 0) Then
			If (GetTemplateAborted) Then
				Exit For
			End If
			Print #iFileID, sBuffer ;
			sBuffer = ""
		End If
	Next
	Print #iFileID, sBuffer ;
	sBuffer = ""

	Close #iFileID
	SetLocale(nCurrentLocale) ' switch back to original locale

End Function

Function CalculatePointsInList(nNumberOfPoints As Long) As Integer

	Dim dStart As Double
	dStart = Timer()

	bSkippedPointsDueToClamping = False

	bCheckMin = CInt(GetScriptSetting("bCheckMin","0"))
	bCheckMax = CInt(GetScriptSetting("bCheckMax","0"))
	dLimitMin = Evaluate(GetScriptSetting("dLimitMin","0.0"))
	dLimitMax = Evaluate(GetScriptSetting("dLimitMax","0.0"))

	Dim EvaluateTmp As Object
	Set EvaluateTmp = Result1DComplex("")

	If CInt(GetScriptSetting("bIgnoreZcomp","0")) = 0 Then
		bIgnoreZcomp = False
	Else
		bIgnoreZcomp = True
	End If

	Dim vxam As Double, vxph As Double
	Dim vyam As Double, vyph As Double
	Dim vzam As Double, vzph As Double

	Dim vxTmp As Double, vyTmp As Double, vzTmp As Double, vNowCST As Double
	Dim vxTmpim As Double, vyTmpim As Double, vzTmpim As Double, vNowCSTim As Double

	Dim vTmpReal As Double, vTmpImag As Double

	Dim s1v As String

	Dim sFxCST As String, bFxCST As Boolean
	sFxCST = GetScriptSetting("Integ_fx","x")
	If (Left((GetScriptSetting("a0DValue","Integral")),14) = "Integral f(x)-") Then
		bFxCST = True
	Else
		bFxCST = False
	End If

	Dim uvw_CST(2) As Double, xyz_CST(2) As Double

	Dim iyy As Long
	Dim bGenerateLogFile As Boolean
	Dim sActionCST As String

	Dim sOutputFormat As String ' Output format of numbers for data file
	sOutputFormat = "  0.0000E+00; -0.0000E+00"

	sActionCST 		= GetScriptSetting("a0DValue","")
	bGenerateLogFile = CBool(GetScriptSetting("GenerateLogfile", "1"))

	vNowCST = 0
	vNowCSTim = 0

	' If bDebugOutput Then ReportInformationToWindow("start copying all results: " + CSTr(Timer()-dStart) + "s")
		
	' Read out list first, store in array
	Dim vxreArr As Variant, vximArr As Variant
	Dim vyreArr As Variant, vyimArr As Variant
	Dim vzreArr As Variant, vzimArr As Variant

	vxreArr = VectorPlot3D.GetList("xre")
	vyreArr = VectorPlot3D.GetList("yre")
	vzreArr = VectorPlot3D.GetList("zre")
	vximArr = VectorPlot3D.GetList("xim")
	vyimArr = VectorPlot3D.GetList("yim")
	vzimArr = VectorPlot3D.GetList("zim")

	Dim vxre As Double, vxim As Double
	Dim vyre As Double, vyim As Double
	Dim vzre As Double, vzim As Double

	Dim dtmp2 As Double, dtmp3 As Double
	Dim dxtmp As Double, dytmp As Double, dztmp As Double, davg As Double, ia As Integer
	Dim dx2 As Double, dy2 As Double, dz2 As Double
	
	' If bDebugOutput Then ReportInformationToWindow("stop copying all results: " + CSTr(Timer()-dStart) + "s")

	For iyy = 0 To nNumberOfPoints-1

		xyz_CST(0) = dListItemX(iyy)
		xyz_CST(1) = dListItemY(iyy)
		xyz_CST(2) = dListItemZ(iyy)

		uvw_CST(0) = dListItemU(iyy)
		uvw_CST(1) = dListItemV(iyy)
		uvw_CST(2) = dListItemW(iyy)

		' Calculate voxel
		If (bGenerateLogFile Or InStr("FielMeanInteLengAreaVoluStat", Left(sActionCST, 4))>0) Then
			If (bMultiplyRadiusLater) And (Not bMultiplyRsinThetaLater) Then
				dVoxel_Unit = uvw_CST(0) * dRef_Voxel_Unit
				dVoxel_SI   = uvw_CST(0) * dRef_Voxel_SI
			ElseIf (bMultiplyRsinThetaLater) Then
				dRsinTheta = uvw_CST(0) * sinD(uvw_CST(1))
				If (bMultiplyRadiusLater) Then
					dVoxel_Unit = uvw_CST(0) * dRsinTheta * dRef_Voxel_Unit
					dVoxel_SI   = uvw_CST(0) * dRsinTheta * dRef_Voxel_SI
				Else
					dVoxel_Unit = dRsinTheta * dRef_Voxel_Unit
					dVoxel_SI   = dRsinTheta * dRef_Voxel_SI
				End If
			End If
		End If
		'
		vxre = vxreArr(iyy)
		vyre = vyreArr(iyy)
		vzre = vzreArr(iyy)
		vxim = vximArr(iyy)
		vyim = vyimArr(iyy)
		vzim = vzimArr(iyy)
		'
		If bScalar_CST_arbitr Then
			vyre = 0.0
			vyim = 0.0
			vzre = 0.0
			vzim = 0.0
		ElseIf bIgnoreZcomp Then
			vzre = 0.0
			vzim = 0.0
		End If

		Select Case sComplex_CST_arbitr
			Case "Complex"
				vxTmp = vxre
				vyTmp = vyre
				vzTmp = vzre
				vxTmpim = vxim
				vyTmpim = vyim
				vzTmpim = vzim
			Case "Real Part"
				vxTmp = vxre
				vyTmp = vyre
				vzTmp = vzre
			Case "Imag. Part"
				vxTmp = vxim
				vyTmp = vyim
				vzTmp = vzim
			Case "Mag"
				vxTmp = Sqr(vxre^2+vxim^2)
				vyTmp = Sqr(vyre^2+vyim^2)
				vzTmp = Sqr(vzre^2+vzim^2)
			Case "Phase"
				vxTmp = ATn2D(vxim,vxre)
				vyTmp = ATn2D(vyim,vyre)
				vzTmp = ATn2D(vzim,vzre)
		End Select
		'
		If bScalar_CST_arbitr Then
			vNowCST = vxTmp
			If (sComplex_CST_arbitr = "Complex") Then
				vNowCSTim = vxTmpim
			End If
		Else
			If sComponent_CST_arbitr = "Abs" Then
				Select Case sComplex_CST_arbitr
					Case "Real Part", "Imag. Part", "Complex"
						vNowCST = Sqr(vxTmp^2+vyTmp^2+vzTmp^2)
						If (sComplex_CST_arbitr = "Complex") Then
							vNowCSTim = Sqr(vxTmpim^2+vyTmpim^2+vzTmpim^2)
						End If
					Case "Mag"
						vNowCST = GetMaxVectorLength (vxre, vxim,  vyre,  vyim,  vzre,  vzim)
					Case "Phase"
						If vxTmp=vyTmp And vxTmp=vzTmp Then
							vNowCST = vxTmp
						Else
							vNowCST = 0.0
						End If
					Case "Average"
						' Average only evaluated, if component is "abs"
						vxTmp = IIf (vxre=0.0 And vxim=0.0, 0.0, ATn2D(vxim,vxre))
						vyTmp = IIf (vyre=0.0 And vyim=0.0, 0.0, ATn2D(vyim,vyre))
						vzTmp = IIf (vzre=0.0 And vzim=0.0, 0.0, ATn2D(vzim,vzre))
						'
						dxtmp = vxre*vxre + vxim*vxim
						dytmp = vyre*vyre + vyim*vyim
						dztmp = vzre*vzre + vzim*vzim
						'
						davg = 0.0
						For ia = 0 To 17
							dx2 = cosD(vxTmp + (ia*10))
							dy2 = cosD(vyTmp + (ia*10))
							dz2 = cosD(vzTmp + (ia*10))
							'
							davg = davg + Sqr( dxtmp*dx2*dx2 + dytmp*dy2*dy2 + dztmp*dz2*dz2 )
						Next ia
						vNowCST = davg / 18.0
					Case "RMS"
						' RMS only evaluated, if component is "abs"
						dtmp2 = vxre*vxre + vyre*vyre + vzre*vzre
						dtmp3 = vxim*vxim + vyim*vyim + vzim*vzim
						vNowCST = Sqr( 0.5 * (dtmp2 + dtmp3) )
				End Select
			Else
				If (Not bLokVC_CST_arbitr) Then
					' transform actual needed component into xyz (vc(0),vc(1),vc(2))
					' for all cartesian coordinates (global and local), vc-Vector already exists (constant for all points)
					If (bLokVB_CST_arbitr) Then
						MsgBox "Vector should have been calculated." + vbCrLf + _
								"Please contact technical support.", vbExclamation
					Else
						Select Case icoordsystem_CST_arbitr
							Case 1 ' cylindrical  rFz
								Convert_Vector_Cylindrical2Cartesian uvw_CST(1), d_va_CST_arbitr, d_vb_CST_arbitr
							Case 2 ' spherical   rTP
								Convert_Vector_Spherical2Cartesian uvw_CST(1), uvw_CST(2), d_va_CST_arbitr, d_vb_CST_arbitr
						End Select
						If (iWCS_CST_arbitr = 0) Then
							d_vc_CST_arbitr(0) = d_vb_CST_arbitr(0)
							d_vc_CST_arbitr(1) = d_vb_CST_arbitr(1)
							d_vc_CST_arbitr(2) = d_vb_CST_arbitr(2)
						Else
							Convert_Vector2Global d_vb_CST_arbitr, d_vc_CST_arbitr
						End If
					End If
				End If
				'
				If (sComplex_CST_arbitr <> "Phase") Then
					If (sComplex_CST_arbitr = "Mag") Then
						' new "magnitude" part temf may-2012
						vTmpReal = vxre * d_vc_CST_arbitr(0) + vyre * d_vc_CST_arbitr(1) + vzre * d_vc_CST_arbitr(2)
						vTmpImag = vxim * d_vc_CST_arbitr(0) + vyim * d_vc_CST_arbitr(1) + vzim * d_vc_CST_arbitr(2)
						vNowCST = Sqr(vTmpReal^2 + vTmpImag^2)
					Else
						vNowCST = vxTmp * d_vc_CST_arbitr(0) + vyTmp * d_vc_CST_arbitr(1) + vzTmp * d_vc_CST_arbitr(2)
						If (sComplex_CST_arbitr = "Complex") Then
							vNowCSTim = vxTmpim * d_vc_CST_arbitr(0) + vyTmpim * d_vc_CST_arbitr(1) + vzTmpim * d_vc_CST_arbitr(2)
						End If
					End If
				Else
					' new "phase" part temf may-2012
					vTmpReal = vxre * d_vc_CST_arbitr(0) + vyre * d_vc_CST_arbitr(1) + vzre * d_vc_CST_arbitr(2)
					vTmpImag = vxim * d_vc_CST_arbitr(0) + vyim * d_vc_CST_arbitr(1) + vzim * d_vc_CST_arbitr(2)
					vNowCST = ATn2D(vTmpImag,vTmpReal)
				End If
			End If
		End If
		'
		If bCheckMin Then
			If vNowCST < dLimitMin Then
				bSkippedPointsDueToClamping = True
				GoTo SkipThisPoint
			End If
		End If
		If bCheckMax Then
			If vNowCST > dLimitMax  Then
				bSkippedPointsDueToClamping = True
				GoTo SkipThisPoint
			End If
		End If
		'
		If bFxCST Then
			s1v = Replace(Evaluate(vNowCST),",",".")
			vNowCST=Evaluate(Replace$(sFxCST,"x",s1v)) 'jsw integral f(x)
		End If
		'
		If (bGenerateLogFile Or InStr("FielMeanInteLengAreaVoluStat", Left(sActionCST, 4))>0) Then
			dSumVoxel_Unit_CST_arbitr = dSumVoxel_Unit_CST_arbitr + dVoxel_Unit
			dSumIntegral_CST_arbitr = dSumIntegral_CST_arbitr + dVoxel_SI * vNowCST
		End If
		'
		If b1Dplot_CST_arbitr Then
			EvaluateTmp.AppendXYDouble(uvw_CST(iDir_CST_arbitr-1), vNowCST, vNowCSTim)
		End If
		'
		If bWriteFile_CST_arbitr And bLogFileFirstEval Then
			' Old version using vba_globals_all:PP12(...) was slow; Format() is faster for numerical values
			If b1Dplot_CST_arbitr Then
				BufferedFileWriteLine_LIB(iDataFileID, Format(uvw_CST(iDir_CST_arbitr-1), sOutputFormat) + Format(vNowCST, sOutputFormat))
			Else
				BufferedFileWriteLine_LIB(iDataFileID, Format(uvw_CST(0), sOutputFormat) + Format(uvw_CST(1), sOutputFormat) + Format(uvw_CST(2), sOutputFormat) + Format(vNowCST, sOutputFormat) + Format(dVoxel_SI, sOutputFormat))
			End If
		End If
		'
		If vNowCST > dMax_CST_arbitr Then
			dMax_CST_arbitr = vNowCST
			dMax_X_CST_arbitr = uvw_CST(0)
			dMax_Y_CST_arbitr = uvw_CST(1)
			dMax_Z_CST_arbitr = uvw_CST(2)
		End If
		If vNowCST < dMin_CST_arbitr Then
			dMin_CST_arbitr = vNowCST
			dMin_X_CST_arbitr = uvw_CST(0)
			dMin_Y_CST_arbitr = uvw_CST(1)
			dMin_Z_CST_arbitr = uvw_CST(2)
		End If
		'
		If (bSolids_CST) Then
			'
			iNowSolid_CST_arbitr = iListItemSolidID(iyy)
			'
			If (iNowSolid_CST_arbitr = 0) Then MsgBox "Problem in Solid Handling"
			'
			' now writing solid-specific arrays  (ID: iNowSolid_CST_arbitr)
			'
			dSolid_Integral_CST_arbitr(iNowSolid_CST_arbitr-1) = dSolid_Integral_CST_arbitr(iNowSolid_CST_arbitr-1) + dVoxel_SI * vNowCST
			dSolid_Volume_CST_arbitr(iNowSolid_CST_arbitr-1) = dSolid_Volume_CST_arbitr(iNowSolid_CST_arbitr-1) + dVoxel_SI
			nSolid_Data_CST_arbitr(iNowSolid_CST_arbitr-1) = nSolid_Data_CST_arbitr(iNowSolid_CST_arbitr-1) + 1
			'
			If (vNowCST > dSolid_Maximum_CST_arbitr(iNowSolid_CST_arbitr-1)) Then dSolid_Maximum_CST_arbitr(iNowSolid_CST_arbitr-1) = vNowCST
			If (vNowCST < dSolid_Minimum_CST_arbitr(iNowSolid_CST_arbitr-1)) Then dSolid_Minimum_CST_arbitr(iNowSolid_CST_arbitr-1) = vNowCST
			'
		End If
		'
	SkipThisPoint:

		' Stop if user clicks on "Abort"; only check every 100th sample
		If iyy Mod 100 = 0 Then
			If (GetTemplateAborted) Then
				Exit Function
			End If
		End If

	Next iyy
	
	If bDebugOutput Then ReportInformationToWindow("done loop calcualting points: " + CSTr(Timer()-dStart) + "s")

	If (b1Dplot_CST_arbitr) Then
		EvaluateTmp.Save(GetProjectPath("Result")+"tmp-arbitr-coord_"+GetScriptSetting("sName", "")+GetScriptSetting("TemplateType", "")+".sig")
	End If

End Function

Function Evaluate0D() As Double

	bLogFileFirstEval = CBool(GetScriptSetting("GenerateLogfile", "1"))
	bDSTemplate = Left(GetApplicationName,2)="DS"

	sWarningTemplateName = sWarningTemplateName_Local

	Dim sFieldCST As String, sfrqtime As String, dvalue As Double, i As Long

	sFieldCST = GetScriptSetting("ResultTreeNameArbitr","")
	sFieldCST = ReplaceVariables(sFieldCST)

	If CInt(GetScriptSetting("CheckBoxFrqTimeActive",0))=0 Then
		sfrqtime = ""
		dvalue=0.0
	Else
		Select Case CInt(GetScriptSetting("GroupConstSweep",0))
		Case 0 ' single frequency
			sfrqtime = "frq"
			dvalue = Evaluate(GetScriptSetting("tflow","0.0"))
		Case 3
			sfrqtime = "time"
			dvalue = Evaluate(GetScriptSetting("tflow","0.0"))
		Case Else
			ReportError sWarningTemplateName + "Error with frq/time sweep in Evaluate0D"
			Evaluate0D = lib_rundef
		End Select
	End If

	If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dvalue) Then
		If PerformEvaluation() Then
			Evaluate0D = cst_value_arbitr
		Else
			Evaluate0D = lib_rundef
		End If
	Else
		Evaluate0D = lib_rundef
	End If

End Function

Function EvaluateMultiple0D(ncount As Long, sName As String, sTableName As String) As Double

	Select Case ncount
		Case 1 ' first run, initialize and calculate
			Evaluate0D()
			sTableName = sName + "\Maximum"
			EvaluateMultiple0D = GetScriptSetting("MaximumM0D", "0")
		Case 2
			sTableName = sName + "\Maximum X-Position"
			EvaluateMultiple0D = GetScriptSetting("MaximumXM0D", "0")
		Case 3
			sTableName = sName + "\Maximum Y-Position"
			EvaluateMultiple0D = GetScriptSetting("MaximumYM0D", "0")
		Case 4
			sTableName = sName + "\Maximum Z-Position"
			EvaluateMultiple0D = GetScriptSetting("MaximumZM0D", "0")
		Case 5
			sTableName = sName + "\Minimum"
			EvaluateMultiple0D = GetScriptSetting("MinimumM0D", "0")
		Case 6
			sTableName = sName + "\Minimum X-Position"
			EvaluateMultiple0D = GetScriptSetting("MinimumXM0D", "0")
		Case 7
			sTableName = sName + "\Minimum Y-Position"
			EvaluateMultiple0D = GetScriptSetting("MinimumYM0D", "0")
		Case 8
			sTableName = sName + "\Minimum Z-Position"
			EvaluateMultiple0D = GetScriptSetting("MinimumZM0D", "0")
		Case 9
			sTableName = sName + "\Mean"
			EvaluateMultiple0D = GetScriptSetting("MeanM0D", "0")
		Case 10
			sTableName = sName + "\Deviation"
			EvaluateMultiple0D = GetScriptSetting("DeviationM0D", "0")
		Case 11
			sTableName = sName + "\Volume"
			EvaluateMultiple0D = GetScriptSetting("VoxelSumM0D", "0")
			If (EvaluateMultiple0D = 0) Then ' volume and integral were not calculated, abort here
				sTableName = ""
				EvaluateMultiple0D = -lib_rundef
			End If
		Case 12
			sTableName = sName + "\Integral"
			EvaluateMultiple0D = GetScriptSetting("IntegralM0D", "0")
		Case Else
			sTableName = ""
			EvaluateMultiple0D = -lib_rundef
	End Select

End Function

Function Evaluate1D() As Object

	bLogFileFirstEval = CBool(GetScriptSetting("GenerateLogfile", "1"))
	bDSTemplate = Left(GetApplicationName,2)="DS"

	sWarningTemplateName = sWarningTemplateName_Local

	Dim sFieldCST As String, sfrqtime As String, dvalue As Double, bok As Boolean
	Dim iNumberOfSteps As Long, i As Long, j As Long
	Dim oRes1DC As Object
	Dim bSweepMultipleMonitors As Boolean

	Dim b_sweep_frq As Boolean
	Dim b_sweep_time As Boolean

	Dim dtf_low As Double, dtf_high As Double, dtf_step As Double, dtf(9999) As Double
	Dim iStart As Long, iEnd As Long

	bSweepMultipleMonitors = False

	sFieldCST = GetScriptSetting("ResultTreeNameArbitr","")
	sFieldCST = ReplaceVariables(sFieldCST)

	Set Evaluate1D = Result1D("")

	Mesh.ViewMeshMode  False
	bok = SelectTreeItem("2D/3D Results\"+ sFieldCST)
	If Not bok Then
		ReportError("Could not find result " & "'2D/3D Results\" & sFieldCST & "'. Please check your settings.")
	End If

	DetermineSweepSettings(b_sweep_frq, b_sweep_time, bSweepMultipleMonitors, dtf_low, dtf_high, dtf_step, dtf(), iNumberOfSteps, iStart, iEnd, sfrqtime, dvalue)

	If b_sweep_frq Or b_sweep_time Then
		If Not bSweepMultipleMonitors Then
			For i=1 To iNumberOfSteps
				If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dtf(i)) Then
					Wait 1e-5
					If PerformEvaluation() Then
						Evaluate1D.AppendXY dtf(i), cst_value_arbitr
						bLogFileFirstEval = False
					End If
				End If
				If GetTemplateAborted Then Exit Function
			Next i
		Else
			' Evaluate0D for all frequencies, in vba_globals_3d.lib
			Set Evaluate1D = Evaluate0D1DCForAllDiscreteMonitorFrequencies_LIB("ResultTreeNameArbitr", "Evaluate0D")
		End If
		If b_sweep_frq Then
			Evaluate1D.Xlabel "Frequency / " + Units.GetFrequencyUnit
		End If
		If b_sweep_time Then
			Evaluate1D.Xlabel "Time / " + Units.GetTimeUnit
		End If
	Else
		' now we are in 1dplot
		If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dvalue) Then
			If PerformEvaluation() Then
				Set oRes1DC = Result1DComplex(GetProjectPath("Result")+"tmp-arbitr-coord_"+GetScriptSetting("sName", "")+GetScriptSetting("TemplateType", "")+".sig")
				Set Evaluate1D = oRes1DC.Real ' fsr 12/24/2012: At this point, the real part may contain Re/Im/Mag/Ph, values are assigned above in "CalculatePointsInList"
				Evaluate1D.Title("") ' .Real adds title "Real Part", which has no meaning at this point and causes problems with sweep option. Remove title. FSR 6/20/2016

				bLogFileFirstEval = False

				' now create the XLabel for the Table-Plot

				Dim sXLabel_CST As String, sVaryingCoordinate As String
				sVaryingCoordinate = GetScriptSetting("VaryingCoordinate", "x")
				sXLabel_CST = ""

				If (iWCS_CST_arbitr > 0) Then
					sXLabel_CST = sXLabel_CST + aWCS(iWCS_CST_arbitr) + "    "
				End If

				sXLabel_CST = sXLabel_CST + sVaryingCoordinate + " / "

				If (sVaryingCoordinate="F" Or sVaryingCoordinate="Theta" Or sVaryingCoordinate="Phi" ) Then
					sXLabel_CST = sXLabel_CST + "Degree"
				Else
					sXLabel_CST = sXLabel_CST + Units.GetGeometryUnit
				End If

				Evaluate1D.Xlabel sXLabel_CST

			ElseIf GetTemplateAborted Then
				Exit Function
			Else
				Evaluate1D.AppendXY  0, lib_rundef
			End If
		End If
	End If

	If Evaluate1D.GetN < 1 Then
		Evaluate1D.AppendXY  0, lib_rundef
	End If

End Function

Function EvaluateMultiple1D(ncount As Long, sName As String, sTableName As String) As Object
	Set EvaluateMultiple1D = EvaluateMultiple(ncount, sName, sTableName)
End Function

Function Evaluate1DComplex() As Object

	bLogFileFirstEval = CBool(GetScriptSetting("GenerateLogfile", "1"))
	bDSTemplate = Left(GetApplicationName,2)="DS"

	sWarningTemplateName = sWarningTemplateName_Local

	Dim sFieldCST As String, sfrqtime As String, dvalue As Double, bok As Boolean
	Dim iNumberOfSteps As Long, i As Long, j As Long
	Dim bSweepMultipleMonitors As Boolean

	Dim b_sweep_frq As Boolean
	Dim b_sweep_time As Boolean

	Dim dtf_low As Double, dtf_high As Double, dtf_step As Double, dtf(9999) As Double
	Dim iStart As Long, iEnd As Long

	bSweepMultipleMonitors = False
	sFieldCST = GetScriptSetting("ResultTreeNameArbitr","")
	sFieldCST = ReplaceVariables(sFieldCST)

	Set Evaluate1DComplex = Result1DComplex("")
	Mesh.ViewMeshMode  False
	bok = SelectTreeItem("2D/3D Results\"+ sFieldCST)
	If Not bok Then
		ReportError("Could not find result " & "'2D/3D Results\" & sFieldCST & "'. Please check your settings.")
	End If

	DetermineSweepSettings(b_sweep_frq, b_sweep_time, bSweepMultipleMonitors, dtf_low, dtf_high, dtf_step, dtf(), iNumberOfSteps, iStart, iEnd, sfrqtime, dvalue)

	If b_sweep_frq Or b_sweep_time Then
		If Not bSweepMultipleMonitors Then
			For i=1 To iNumberOfSteps
				If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dtf(i)) Then
					Wait 1e-5
					If PerformEvaluation() Then
						' fsr 12/24/2012: The line below will fail. Currently, it is never triggered because sweep time/freq not allowed for complex results
						' Imaginary part needs to be added as third parameter to AppendXY once time/freq sweeps are allowed.
						Evaluate1DComplex.AppendXY(dtf(i), cst_value_arbitr)
						bLogFileFirstEval = False
					End If
				End If
				If GetTemplateAborted Then Exit Function
			Next i
		Else
			ReportError("Evaluate1DComplex: This option currently does not support frequency sweeps for discrete monitors.")
		End If
		If b_sweep_frq Then
			Evaluate1DComplex.Xlabel "Frequency / " + Units.GetFrequencyUnit
		End If
		If b_sweep_time Then
			Evaluate1DComplex.Xlabel "Time / " + Units.GetTimeUnit
		End If
	Else
		' now we are in 1dplot
		If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dvalue) Then
			If PerformEvaluation() Then
				Evaluate1DComplex.Load(GetProjectPath("Result")+"tmp-arbitr-coord_"+GetScriptSetting("sName", "")+GetScriptSetting("TemplateType", "")+".sig")

				bLogFileFirstEval = False

				' now create the XLabel for the Table-Plot

				Dim sXLabel_CST As String, sVaryingCoordinate As String
				sVaryingCoordinate = GetScriptSetting("VaryingCoordinate", "x")
				sXLabel_CST = ""

				If (iWCS_CST_arbitr > 0) Then
					sXLabel_CST = sXLabel_CST + aWCS(iWCS_CST_arbitr) + "    "
				End If

				sXLabel_CST = sXLabel_CST + sVaryingCoordinate + " / "

				If (sVaryingCoordinate="F" Or sVaryingCoordinate="Theta" Or sVaryingCoordinate="Phi" ) Then
					sXLabel_CST = sXLabel_CST + "Degree"
				Else
					sXLabel_CST = sXLabel_CST + Units.GetGeometryUnit
				End If

				Evaluate1DComplex.Xlabel sXLabel_CST
				
				'set logarithmic factor for correct dB scaling
				Dim logFactor As Double
				logFactor = VectorPlot3D.GetLogarithmicFactor()
				If (logFactor > 0.0) Then
					Evaluate1DComplex.SetLogarithmicFactor(logFactor)
				End If

			ElseIf GetTemplateAborted Then
				Exit Function
			Else
				Evaluate1DComplex.AppendXYDouble(0, lib_rundef, lib_rundef)
			End If
		End If
	End If

	If Evaluate1DComplex.GetN < 1 Then
		Evaluate1DComplex.AppendXYDouble(0, lib_rundef, lib_rundef)
	End If

End Function

Function EvaluateMultiple1DComplex(ncount As Long, sName As String, sTableName As String) As Object
	Set EvaluateMultiple1DComplex = EvaluateMultiple(ncount, sName, sTableName)
End Function

Function EvaluateMultiple(ncount As Long, sName As String, sTableName As String) As Object

	Dim sFieldSignatureFolder As String, dFrequency As Double
	Dim sFieldSignatureLeft As String, sFieldSignatureRight As String, sCurrentTreeEntry As String

	Dim sFieldCST As String, sfrqtime As String, dvalue As Double, bok As Boolean
	Dim iNumberOfSteps As Long, i As Long, j As Long, nM0Dcount As Long
	Dim bSweepMultipleMonitors As Boolean

	Dim b_sweep_frq As Boolean
	Dim b_sweep_time As Boolean

	Dim dtf_low As Double, dtf_high As Double, dtf_step As Double, dtf(9999) As Double
	Dim iStart As Long, iEnd As Long

	Dim EvaluateMultiple0DResults() As Object, sNameLocal As String, sTableNamesLocal() As String, nM0DResults As Long

	bSweepMultipleMonitors = False

	' Read in field from the template settings. Decide below if it should be actively selected.
	sFieldCST = GetScriptSetting("ResultTreeNameArbitr","")
	sFieldCST = ReplaceVariables(sFieldCST)

	' Select field only during first run if multiple monitors are evaluated. Selecting it again later interferes with 'all frequencies at discrete monitors'
	' It is the responsibility of this function to ensure that all selected monitors are consistent with the first selection
	If (ncount <= 1) Then
		Mesh.ViewMeshMode  False
		If (GetSelectedTreeItem = "2D/3D Results\"+ sFieldCST) Then
			bok = True
		Else
			bok = SelectTreeItem("2D/3D Results\"+ sFieldCST)
		End If
		If Not bok Then
			ReportError("Could not find result " & "'2D/3D Results\" & sFieldCST & "'. Please check your settings.")
		End If
	End If

	' DetermineSweepSettings needs to be called after a result has been selected
	DetermineSweepSettings(b_sweep_frq, b_sweep_time, bSweepMultipleMonitors, dtf_low, dtf_high, dtf_step, dtf(), iNumberOfSteps, iStart, iEnd, sfrqtime, dvalue)

	' On abort, restore original settings and leave
	If GetTemplateAborted Then
		sTableName = ""
		Exit Function
	End If

	If b_sweep_frq Or b_sweep_time Then ' sweep time or frequency
		If (Not bSweepMultipleMonitors And ((ncount <= iNumberOfSteps) Or (GetScriptSetting("M0DTableName_" & CStr(ncount), "") <> ""))) Then
			If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dtf(ncount)) Then
				dFrequency = dtf(ncount)
				' Temporarily disable sweep
				StoreScriptSetting("CheckBoxFrqTimeActive",0)
				Select Case GetScriptSetting("TemplateType", "M1D")
					Case "M1D"
						' Could be generated from either 1D or M0D
						If (GetScriptSetting("a0DValue","") = "Statistics-3D (Min/Max/Mean/Deviation...)") Then ' M0D -> M1D
							' Calculate all values for all time steps, then release curves one by one
							If ncount = 1 Then
								' Extract all M0D curves for all frequencies when ncount = 1; release the first M1D curve immediately, then one with each increment of ncount
								For i = 1 To iNumberOfSteps
									If Select2D3DFieldInTree(sFieldCST,bScalar_CST_arbitr,sfrqtime,dtf(i)) Then
										nM0DResults = 0
										Do
											nM0DResults = nM0DResults + 1
											If (i = 1) Then ' only build array during first time step
												ReDim Preserve EvaluateMultiple0DResults(nM0DResults)
												Set EvaluateMultiple0DResults(nM0DResults) = Result1D("")
												ReDim Preserve sTableNamesLocal(nM0DResults)
											End If
											EvaluateMultiple0DResults(nM0DResults).AppendXY(dtf(i), EvaluateMultiple0D(nM0DResults, sNameLocal, sTableNamesLocal(nM0DResults)))
										Loop Until (sTableNamesLocal(nM0DResults) = "")
									End If
									If GetTemplateAborted Then Exit Function
								Next
								For j = 1 To UBound(sTableNamesLocal)
									StoreScriptSetting("M0DTableName_" & CStr(j), sTableNamesLocal(j))
									EvaluateMultiple0DResults(j).Save(GetProjectPath("Result") & "EvaluateMultiple0D_" & CStr(j))
								Next
							End If
							' For ncount >=1, restore and return corresponding result curve
							If (GetScriptSetting("M0DTableName_" & CStr(ncount), "") <> "") Then
								Set EvaluateMultiple = Result1D(GetProjectPath("Result") & "EvaluateMultiple0D_" & CStr(ncount))
							Else
								Set EvaluateMultiple = Nothing ' Needs to be set explicitly to Nothing; otherwise, at least an empty object will exist
							End If
						Else ' 1D -> M1D
							Set EvaluateMultiple = Evaluate1D
						End If
					Case "M1DC"
						Set EvaluateMultiple = Evaluate1DComplex
					Case Else
						ReportError("EvaluateMultiple: Unknown type")
				End Select
				' Enable sweep
				StoreScriptSetting("CheckBoxFrqTimeActive",1)
			End If
			If GetTemplateAborted Then Exit Function
		ElseIf bSweepMultipleMonitors Then
			Select Case GetScriptSetting("TemplateType", "M1D")
				Case "M1DC"
					Set EvaluateMultiple = Evaluate0D1DCForAllDiscreteMonitorFrequencies_LIB("ResultTreeNameArbitr", "Evaluate1DComplex", ncount, dFrequency)
				Case "M1D"
					' Could be generated from either 1D or M0D
					If (GetScriptSetting("a0DValue","") = "Statistics-3D (Min/Max/Mean/Deviation...)") Then ' M0D -> M1D
						Set EvaluateMultiple = Evaluate0D1DCForAllDiscreteMonitorFrequencies_LIB("ResultTreeNameArbitr", "EvaluateMultiple0D", ncount, dFrequency)
					Else ' 1D -> M1D
						Set EvaluateMultiple = Evaluate0D1DCForAllDiscreteMonitorFrequencies_LIB("ResultTreeNameArbitr", "Evaluate1D", ncount, dFrequency)
					End If
				Case Else
					ReportError("EvaluateMultiple: Unknown type")
			End Select
		End If
	Else
		' The only "Multiple" result that is not a sweep right now is Statistics-3D (10/2/2015 FSR)
		' Since that function is handled separately in EvaluateMultiple0D, nothing to do here...
	End If
	If EvaluateMultiple Is Nothing Then
		sTableName = ""
	ElseIf (GetScriptSetting("a0DValue","") = "Statistics-3D (Min/Max/Mean/Deviation...)") Then
		sTableName = sName & GetScriptSetting("M0DTableName_" & CStr(ncount), "")
		If b_sweep_frq Then
			EvaluateMultiple.XLabel("Frequency / " & Units.GetFrequencyUnit)
		ElseIf b_sweep_time Then
			EvaluateMultiple.XLabel("Time / " & Units.GetTimeUnit)
		End If
	ElseIf b_sweep_frq Then
		sTableName = sName & "\" & IIf(EvaluateMultiple.GetTitle() <> "", EvaluateMultiple.GetTitle(), "f=" + CStr(dFrequency))
	ElseIf b_sweep_time Then
		sTableName = sName & "\" & IIf(EvaluateMultiple.GetTitle() <> "", EvaluateMultiple.GetTitle(), "t=" + CStr(dtf(ncount)))
	End If

End Function

Sub DetermineSweepSettings(b_sweep_frq As Boolean, b_sweep_time As Boolean, bSweepMultipleMonitors As Boolean, _
							dtf_low As Double, dtf_high As Double, dtf_step As Double, dtf() As Double, _
							iNumberOfSteps As Long, iStart As Long, iEnd As Long, sfrqtime As String, dvalue As Double)

	Dim i As Long, d2 As Double, j As Long, imin As Long
	Dim sTempString As String

	If CInt(GetScriptSetting("CheckBoxFrqTimeActive",0))=0 Then
		b_sweep_frq = False
		b_sweep_time = False
		sfrqtime = ""
		dvalue=0.0
	Else
		Select Case CInt(GetScriptSetting("GroupConstSweep",0))
		Case 0
			b_sweep_frq = False
			b_sweep_time = False
			sfrqtime = "frq"
			dvalue = Evaluate(GetScriptSetting("tflow","0.0"))
		Case 3
			b_sweep_frq = False
			b_sweep_time = False
			sfrqtime = "time"
			dvalue = Evaluate(GetScriptSetting("tflow","0.0"))
		Case 1 ' all frq
			b_sweep_frq = True
			b_sweep_time = False
			sfrqtime = "frq"
			With ResultMap("")
				If .IsValid Then
					iNumberOfSteps = .GetItemCount()
					For i = 1 To iNumberOfSteps
						iStart = InStr(.GetItemParameters(i),"Frequency=")
						sTempString = Mid(.GetItemParameters(i),iStart+10)
						iEnd = InStr(sTempString,";")
						If (iEnd > 0) Then
							sTempString = Left(sTempString, iEnd-1)
						End If
						dtf(i) = Evaluate(sTempString)*Units.GetFrequencySIToUnit
					Next i
					' sort values, since frq-values might be mixed
					For i = 1 To iNumberOfSteps-1
						imin = i
						For j = i+1 To iNumberOfSteps
							If dtf(j) < dtf(imin) Then
								imin = j
							End If
						Next j
						If (imin <> i) Then
							d2 = dtf(i)
							dtf(i) = dtf(imin)
							dtf(imin) = d2
						End If
					Next i
				Else
					bSweepMultipleMonitors = True
				End If
			End With
		Case 4 ' all time
			b_sweep_frq = False
			b_sweep_time = True
			sfrqtime = "time"
			With ResultMap("")
				If .IsValid Then
					iNumberOfSteps = .GetItemCount()
					For i = 1 To iNumberOfSteps
						iStart = InStr(.GetItemParameters(i),"Time=")
						sTempString = Mid(.GetItemParameters(i),iStart+5)
						iEnd = InStr(sTempString,";")
						If (iEnd > 0) Then
							sTempString = Left(sTempString, iEnd-1)
						End If
						dtf(i) = Evaluate(sTempString)*Units.GetTimeSIToUnit
					Next i
				Else
					With VectorPlot3D
						dtf_low  = .GetTStart
						dtf_high = (.GetTStart+.GetTStep*(.GetNumberOfSamples-0.9))
						dtf_step = .GetTStep
					End With
					'
					' MsgBox Cstr(dtf_low) + vbCrLf + Cstr(dtf_high) + vbCrLf + Cstr(dtf_step)
					iNumberOfSteps=0
					For dvalue=dtf_low To dtf_high STEP dtf_step
						iNumberOfSteps = iNumberOfSteps+1
						dtf(iNumberOfSteps) = dvalue
					Next
					'MsgBox Cstr(count) + vbCrLf + Cstr(dtf(count))
				End If
			End With
		Case 2 ' userdef frq
			If Not ResultMap("").IsValid() Then ReportError(sWarningTemplateName & "User-defined frequency sweeps are not supported for discrete monitor frequencies.")
			b_sweep_frq = True
			b_sweep_time = False
			sfrqtime = "frq"
			dtf_low  = Evaluate(GetScriptSetting("tflow","0.0"))
			dtf_high = Evaluate(GetScriptSetting("tfhigh","0.0"))
			dtf_step = Evaluate(GetScriptSetting("tfstepsize","0.0"))
			'
			iNumberOfSteps=0
			For dvalue=dtf_low To dtf_high STEP dtf_step
				iNumberOfSteps = iNumberOfSteps+1
				dtf(iNumberOfSteps) = dvalue
			Next
		Case 5 ' userdef time
			b_sweep_frq  = False
			b_sweep_time = True
			sfrqtime = "time"
			dtf_low  = Evaluate(GetScriptSetting("tflow","0.0"))
			dtf_high = Evaluate(GetScriptSetting("tfhigh","0.0"))
			dtf_step = Evaluate(GetScriptSetting("tfstepsize","0.0"))
			'
			iNumberOfSteps=0
			For dvalue=dtf_low To dtf_high STEP dtf_step
				iNumberOfSteps = iNumberOfSteps+1
				dtf(iNumberOfSteps) = dvalue
			Next
		End Select
	End If

End Sub

Function PerformEvaluation(Optional dAutostepReduction As Double, Optional bPreviewOnly As Boolean) As Boolean

	Dim dMeshMin As Double, dMeshMax As Double, dAbsMeshMax As Double
	Dim nNumberOfPoints As Long, nMinimumNumberOfPoints As Long, dRefinementFactor As Double
	Dim iDim_CST As Integer
	Dim dstepsize_CST As Double, bmaxrange_CST As Boolean
	Dim bGenerateLogFile As Boolean
	Dim ddStep As Double, nIndex As Long
	Dim sActionCST As String
	Dim xyzbox(3,2) As Double
	Dim bAngle(3) As Boolean
	Dim dstptmp As Double, nstpstmp As Long, dminmaxtmp As Double, dRmeanTmp As Double
	Dim dStepLength(3) As Double
	Dim im1Dir_CST As Integer
	Dim dSumVoxel_SI As Double
	Dim dMeanCST As Double
	Dim dDeviCST As Double
	Dim dNorming As Double

	Dim dStartTime As Double
	dStartTime = Timer()
	If bDebugOutput Then ReportInformationToWindow("--------------------------------------------------------------")

	nCurrentLocale = GetLocale ' save current locale, temporarily switch to US locale for numerical file output using period; this is faster than manually replacing , with . in each string

	iDim_CST 				= CInt(GetScriptSetting("edim","0"))
	icoordsystem_CST_arbitr	= CInt(GetScriptSetting("coordsystem","0"))
	iWCS_CST_arbitr 		= CInt(GetScriptSetting("wcs","0"))
	iDir_CST_arbitr			= 1+CInt(GetScriptSetting("coordinates","0"))
	dstepsize_CST			= Evaluate(GetScriptSetting("stepsize","0.0"))
	bmaxrange_CST			= 1=CInt(GetScriptSetting("maxrange","1"))
	bWriteFile_CST_arbitr	= 1=CInt(GetScriptSetting("WriteFile","1"))
	bGenerateLogFile 		= CBool(GetScriptSetting("GenerateLogfile", "1"))

	sComponent_CST_arbitr	= GetScriptSetting("aComponent","")
	sComplex_CST_arbitr		= GetScriptSetting("aComplex","")
	sActionCST 		= GetScriptSetting("a0DValue","")

	' dAutostepReduction can be used to recursively reduce the step width if it is set to "auto"; if it is missing, it should be "1"
	If dAutostepReduction = 0 Then dAutostepReduction = 1
	nMinimumNumberOfPoints = 10*10^iDim_CST

	PerformEvaluation = True

	FillWCSArray

	Boundary.GetCalculationBox x1box_CST_arbitr, x2box_CST_arbitr, y1box_CST_arbitr, y2box_CST_arbitr, z1box_CST_arbitr, z2box_CST_arbitr

	' . get min and max meshstep
	dAbsMeshMax = 0

	With Mesh

		On Error GoTo NoMeshExists

		dMeshMin = .GetMinimumEdgeLength
		dMeshMax = .GetMaximumEdgeLength

		On Error GoTo 0
		GoTo MeshExists

	NoMeshExists:
		PerformEvaluation = False
		Exit Function

	MeshExists:

		If Abs(x1box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(x1box_CST_arbitr)
		If Abs(y1box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(y1box_CST_arbitr)
		If Abs(z1box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(z1box_CST_arbitr)

		If Abs(x2box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(x2box_CST_arbitr)
		If Abs(y2box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(y2box_CST_arbitr)
		If Abs(z2box_CST_arbitr) > dAbsMeshMax Then dAbsMeshMax = Abs(z2box_CST_arbitr)

	End With

	' now dMeshMin, dMeshMax  contain the smallest and biggest meshstep
	' dAbsMeshMax is the biggest absolute dimension from origin (useful for maxrange guess)
	b1Dplot_CST_arbitr = (sActionCST = "1D Plot of Field Values")

	' read solid information

	bSolids_CST = CInt(GetScriptSetting("bSolids","0"))
	nSolids_CST = CInt(GetScriptSetting("nSolids","0"))

	If ( iDim_CST < 1 ) Then
		bSolids_CST = False
		bWriteFile_CST_arbitr = True
	End If

	If (bSolids_CST) Then
		If (nSolids_CST > 0) Then
			ReDim aSolidArray_CST(nSolids_CST-1)

			ReDim dSolid_Integral_CST_arbitr(nSolids_CST-1)
			ReDim dSolid_Volume_CST_arbitr(nSolids_CST-1)
			ReDim nSolid_Data_CST_arbitr(nSolids_CST-1)
			ReDim dSolid_Maximum_CST_arbitr(nSolids_CST-1)
			ReDim dSolid_Minimum_CST_arbitr(nSolids_CST-1)

			For iSolid_CST = 1 To nSolids_CST
				aSolidArray_CST(iSolid_CST-1) = GetScriptSetting("Solid" + CStr(iSolid_CST),"")

				dSolid_Integral_CST_arbitr(iSolid_CST-1) = 0.0
				dSolid_Volume_CST_arbitr(iSolid_CST-1) = 0.0
				nSolid_Data_CST_arbitr(iSolid_CST-1) = 0
				dSolid_Maximum_CST_arbitr(iSolid_CST-1) = lib_rundef ' -1.23456e27
				dSolid_Minimum_CST_arbitr(iSolid_CST-1) = - lib_rundef ' +1.23456e27
			Next
		End If
	End If

	dUVWvalue(1,1) = Evaluate (GetScriptSetting("u1","0.0"))
	dUVWvalue(2,1) = Evaluate (GetScriptSetting("v1","0.0"))
	dUVWvalue(3,1) = Evaluate (GetScriptSetting("w1","0.0"))
	dUVWvalue(1,2) = Evaluate (GetScriptSetting("u2","0.0"))
	dUVWvalue(2,2) = Evaluate (GetScriptSetting("v2","0.0"))
	dUVWvalue(3,2) = Evaluate (GetScriptSetting("w2","0.0"))

	' set stepsize ddStep (equidistant) dep on dimension and user choice
	If iDim_CST = 0 Then
		' 0D
		ddStep = 0
	Else
		' 1D / 2D / 3D
		If dstepsize_CST = 0.0 Then
			' automatic choice
			Select Case Mesh.GetMeshType
			Case "Surface", "Tetrahedral"
				If 50.0*dMeshMin > dMeshMax Then
					If iDim_CST = 3 Then
						ddStep = 2.0 * dMeshMin ' for 3D a coarser sampling is used than for 1D and 2D
					Else
						ddStep = 0.5 * dMeshMin
					End If
				Else
					ddStep = 0.1 * dMeshMax
				End If
			Case Else
				' hex
				If iDim_CST = 3 Then
					ddStep = 0.5 * (dMeshMin+dMeshMax)' for 3D  mean meshstep is used
				Else
					If iDim_CST = 2 Then
						ddStep = 0.25 * (dMeshMin+dMeshMax) ' for 2D half of mean meshstep is used
					Else
						ddStep = 0.5 * dMeshMin ' for 1D half of min meshstep is used
					End If
				End If
			End Select
			ddStep = ddStep/dAutostepReduction ' used for automatic refinement
		Else
			ddStep = dstepsize_CST
		End If
	End If

	xyzbox(1,1) = x1box_CST_arbitr
	xyzbox(1,2) = x2box_CST_arbitr
	xyzbox(2,1) = y1box_CST_arbitr
	xyzbox(2,2) = y2box_CST_arbitr
	xyzbox(3,1) = z1box_CST_arbitr
	xyzbox(3,2) = z2box_CST_arbitr

	If (bSolids_CST) Then
		'
		' --- if solids are selected, adjust max xyzbox bounding box according to the loose smallest box, containing all selected solids
		'
		Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, z1 As Double, z2 As Double, bFirstSolid As Boolean, ss2 As String
		bFirstSolid = True
		For iSolid_CST = 1 To nSolids_CST
			If Solid.GetLooseBoundingBoxOfShape("solid$"+aSolidArray_CST(iSolid_CST-1), x1,x2,y1,y2,z1,z2) Then
				' loose bounding box might be larger than actual bounding box -> never go beyond global box!
				If bFirstSolid Then
					If (x1 > x1box_CST_arbitr) Then xyzbox(1,1)=x1
					If (x2 < x2box_CST_arbitr) Then xyzbox(1,2)=x2
					If (y1 > y1box_CST_arbitr) Then xyzbox(2,1)=y1
					If (y2 < y2box_CST_arbitr) Then xyzbox(2,2)=y2
					If (z1 > z1box_CST_arbitr) Then xyzbox(3,1)=z1
					If (z2 < z2box_CST_arbitr) Then xyzbox(3,2)=z2
				Else
					If (xyzbox(1,1) > x1 And x1 >= x1box_CST_arbitr) Then xyzbox(1,1)=x1
					If (xyzbox(1,2) < x2 And x2 <= x2box_CST_arbitr) Then xyzbox(1,2)=x2
					If (xyzbox(2,1) > y1 And y1 >= y1box_CST_arbitr) Then xyzbox(2,1)=y1
					If (xyzbox(2,2) < y2 And y2 <= y2box_CST_arbitr) Then xyzbox(2,2)=y2
					If (xyzbox(3,1) > z1 And z1 >= z1box_CST_arbitr) Then xyzbox(3,1)=z1
					If (xyzbox(3,2) < z2 And z2 <= z2box_CST_arbitr) Then xyzbox(3,2)=z2
				End If
				bFirstSolid = False
			Else
				ReportWarning("Evaluate field in arbitrary coordinates: Solidname not found: "+aSolidArray_CST(iSolid_CST-1))
			End If
		Next

	End If

	' if max-range, then set minmax values dep. on dim and coord.system

	If (bmaxrange_CST) Then

		Select Case iDim_CST

			Case 0 ' 0D

			Case 1 ' 1D

				Select Case icoordsystem_CST_arbitr
					Case 0 ' cartesian
						If (iWCS_CST_arbitr = 0) Then
							' global xyz -> take known box dimensions
							dUVWvalue(iDir_CST_arbitr,1) = xyzbox(iDir_CST_arbitr,1)
							dUVWvalue(iDir_CST_arbitr,2) = xyzbox(iDir_CST_arbitr,2)
						Else
							' local wcs uvw -> minmax not easily predictable - make it safe...
							dUVWvalue(iDir_CST_arbitr,1) = -2*dAbsMeshMax
							dUVWvalue(iDir_CST_arbitr,2) = 2*dAbsMeshMax
						End If
					Case 1 ' cylindrical
						If iDir_CST_arbitr = 1 Then ' r
							dUVWvalue(iDir_CST_arbitr,1) = 0.0
							dUVWvalue(iDir_CST_arbitr,2) = 2*dAbsMeshMax
						ElseIf iDir_CST_arbitr = 2 Then ' f
							dUVWvalue(iDir_CST_arbitr,1) = 0.0
							dUVWvalue(iDir_CST_arbitr,2) = 360.0
						ElseIf iDir_CST_arbitr = 3 Then ' z
							dUVWvalue(iDir_CST_arbitr,1) = -2*dAbsMeshMax
							dUVWvalue(iDir_CST_arbitr,2) = 2*dAbsMeshMax
						End If
					Case 2 ' spherical
						If iDir_CST_arbitr = 1 Then ' r
							dUVWvalue(iDir_CST_arbitr,1) = 0.0
							dUVWvalue(iDir_CST_arbitr,2) = 2*dAbsMeshMax
						ElseIf iDir_CST_arbitr = 2 Then ' t
							dUVWvalue(iDir_CST_arbitr,1) = 0.0
							dUVWvalue(iDir_CST_arbitr,2) = 180.0
						ElseIf iDir_CST_arbitr = 3 Then ' f
							dUVWvalue(iDir_CST_arbitr,1) = 0.0
							dUVWvalue(iDir_CST_arbitr,2) = 360.0
						End If
				End Select

			Case 2 ' 2D

				Select Case icoordsystem_CST_arbitr
					Case 0 ' cartesian
						If (iWCS_CST_arbitr = 0) Then
							' global xyz -> take known box dimensions
							For nIndex = 1 To 3
								If (nIndex <> iDir_CST_arbitr) Then
									dUVWvalue(nIndex,1) = xyzbox(nIndex,1)
									dUVWvalue(nIndex,2) = xyzbox(nIndex,2)
								End If
							Next
						Else
							' local wcs uvw -> minmax not easily predictable - make it safe...
							For nIndex = 1 To 3
								If (nIndex <> iDir_CST_arbitr) Then
									dUVWvalue(nIndex,1) = -2*dAbsMeshMax
									dUVWvalue(nIndex,2) = 2*dAbsMeshMax
								End If
							Next
						End If
					Case 1 ' cylindrical
						For nIndex = 1 To 3
							If (nIndex <> iDir_CST_arbitr) Then
								If nIndex = 1 Then ' r
									dUVWvalue(nIndex,1) = 0.0
									dUVWvalue(nIndex,2) = 2*dAbsMeshMax
								ElseIf nIndex = 2 Then ' f
									dUVWvalue(nIndex,1) = 0.0
									dUVWvalue(nIndex,2) = 360.0
								ElseIf nIndex = 3 Then ' z
									dUVWvalue(nIndex,1) = -2*dAbsMeshMax
									dUVWvalue(nIndex,2) = 2*dAbsMeshMax
								End If
							End If
						Next
					Case 2 ' spherical
						For nIndex = 1 To 3
							If (nIndex <> iDir_CST_arbitr) Then
								If nIndex = 1 Then ' r
									dUVWvalue(nIndex,1) = 0.0
									dUVWvalue(nIndex,2) = 2*dAbsMeshMax
								ElseIf nIndex = 2 Then ' t
									dUVWvalue(nIndex,1) = 0.0
									dUVWvalue(nIndex,2) = 180.0
								ElseIf nIndex = 3 Then ' f
									dUVWvalue(nIndex,1) = 0.0
									dUVWvalue(nIndex,2) = 360.0
								End If
							End If
						Next
				End Select

			Case 3 ' 3D

				Select Case icoordsystem_CST_arbitr
					Case 0 ' cartesian
						If (iWCS_CST_arbitr = 0) Then
							' global xyz -> take known box dimensions
							For nIndex = 1 To 3
								dUVWvalue(nIndex,1) = xyzbox(nIndex,1)
								dUVWvalue(nIndex,2) = xyzbox(nIndex,2)
							Next
						Else
							' local wcs uvw -> minmax not easily predictable - make it safe...
							For nIndex = 1 To 3
								dUVWvalue(nIndex,1) = -2*dAbsMeshMax
								dUVWvalue(nIndex,2) = 2*dAbsMeshMax
							Next
						End If
					Case 1 ' cylindrical
						dUVWvalue(1,1) = 0.0
						dUVWvalue(1,2) = 2*dAbsMeshMax
						dUVWvalue(2,1) = 0.0
						dUVWvalue(2,2) = 360.0
						dUVWvalue(3,1) = -2*dAbsMeshMax
						dUVWvalue(3,2) = 2*dAbsMeshMax
					Case 2 ' spherical
						dUVWvalue(1,1) = 0.0
						dUVWvalue(1,2) = 2*dAbsMeshMax
						dUVWvalue(2,1) = 0.0
						dUVWvalue(2,2) = 180.0
						dUVWvalue(3,1) = 0.0
						dUVWvalue(3,2) = 360.0
				End Select

		End Select
	End If

	' now maxrange is set, all dUVWvalues() are set now
	' ReportInformationToWindow("Max range set after "+CSTr(dStartTime-Timer())+" secs.")

	' adjust final min, max values and afterwards adjust the stepsize, fitting to it
	' also, recalculate angle-stepwidth from from ddStep
	bAngle(1) = False
	bAngle(2) = False
	bAngle(3) = False

	Select Case icoordsystem_CST_arbitr
		Case 0 ' cartesian   xyz
		Case 1 ' cylindrical  rFz
			bAngle(2) = True
		Case 2 ' spherical   rTP
			bAngle(2) = True
			bAngle(3) = True
	End Select

	Select Case iDim_CST

		Case 0 ' 0D

			' max=min, step=1, for all dimensions and coordinate systems

			For nIndex = 1 To 3
				dUVWvalue(nIndex,2) = dUVWvalue(nIndex,1)
				dUVWvalue(nIndex,3) = 1
			Next 

		Case 1 ' 1D

			' max=min, step = 1 for all directions unequal idir direction

			For nIndex = 1 To 3
				If nIndex <> iDir_CST_arbitr Then
					dUVWvalue(nIndex,2) = dUVWvalue(nIndex,1)
					dUVWvalue(nIndex,3) = 1
				End If
			Next

			If bAngle(iDir_CST_arbitr) Then
				dRmeanTmp = 0.5 * ( dUVWvalue(1,1) + dUVWvalue(1,2) )
				If (dRmeanTmp > 0) Then
					dstptmp = ( ddStep / dRmeanTmp ) * lib_rad2deg
				Else
					PerformEvaluation = False
					Exit Function
				End If
			Else
				dstptmp = ddStep
			End If

			dminmaxtmp = dUVWvalue(iDir_CST_arbitr,2)-dUVWvalue(iDir_CST_arbitr,1)

			nstpstmp = CLng ( dminmaxtmp / dstptmp )
			If nstpstmp < 1 Then nstpstmp = 1

			dUVWvalue(iDir_CST_arbitr,3) = dminmaxtmp / nstpstmp

		Case 2 ' 2D

			' max=min, step = 1 for idir direction

			dUVWvalue(iDir_CST_arbitr,2) = dUVWvalue(iDir_CST_arbitr,1)
			dUVWvalue(iDir_CST_arbitr,3) = 1

			' now handle transversal directions
			' adjust stepsize to minmax range and set min startpoints (2d => NOT for IDIR) at half stepsize

			For nIndex = 1 To 3
				If nIndex <> iDir_CST_arbitr Then

					If bAngle(nIndex) Then
						dRmeanTmp = 0.5 * ( dUVWvalue(1,1) + dUVWvalue(1,2) )
						If (dRmeanTmp > 0) Then
							dstptmp = ( ddStep / dRmeanTmp ) * lib_rad2deg
						Else
							PerformEvaluation = False
							Exit Function
						End If
					Else
						dstptmp = ddStep
					End If

					dminmaxtmp = dUVWvalue(nIndex,2)-dUVWvalue(nIndex,1)

					nstpstmp = CLng ( dminmaxtmp / dstptmp )
					If nstpstmp < 1 Then nstpstmp = 1

					dUVWvalue(nIndex,3) = dminmaxtmp / nstpstmp

				End If
			Next

		Case 3 ' 3D

			' handle all 3 dimensions
			' adjust stepsize to minmax range and set min startpoints (2d => NOT for IDIR) at half stepsize

			For nIndex = 1 To 3

				If bAngle(nIndex) Then
					dRmeanTmp = 0.5 * ( dUVWvalue(1,1) + dUVWvalue(1,2) )
					If (dRmeanTmp > 0) Then dstptmp = ( ddStep / dRmeanTmp ) * lib_rad2deg
				Else
					dstptmp = ddStep
				End If

				dminmaxtmp = dUVWvalue(nIndex,2)-dUVWvalue(nIndex,1)

				nstpstmp = CLng ( dminmaxtmp / dstptmp )
				If nstpstmp < 1 Then nstpstmp = 1

				If dminmaxtmp = 0.0 Then
					dUVWvalue(nIndex,3) = 1.0
				Else
					dUVWvalue(nIndex,3) = dminmaxtmp / nstpstmp
				End If
			Next

	End Select

	If (bSolids_CST) Then
		' Make sure there is an intersection between bounding box and selected solids; if not, this could lead to endless loop
		' Need to wait until this point, where final dUVWvalues are known
		If ((dUVWvalue(1, 1) > xyzbox(1, 2)) Or (dUVWvalue(1, 2)<xyzbox(1, 1))) _
			Or ((dUVWvalue(2, 1) > xyzbox(2, 2)) Or (dUVWvalue(2, 2)<xyzbox(2, 1))) _
			Or ((dUVWvalue(3, 1) > xyzbox(3, 2)) Or (dUVWvalue(3, 2)<xyzbox(3, 1))) Then
			ReportError("Evaluate fields in arbitrary coordinates: Intersection between selected evaluation range and selected solid(s) is zero. Please check your settings.")
		End If
	End If

	If bLogFileFirstEval Then

		Dim ntoday As Long, ntime As Long
		Dim sLogFile_CST As String
		Dim sDataFile_CST As String

		' always write new log-file name for every evaluate (otherwise overwriting result files by duplicated templates)
		ntoday = CLng(Day(Date)) + 100 * CLng(Month(Date)) + 10000 * CLng(Year(Date))
		ntime = CLng(Second(Time)) + 100 * CLng(Minute(Time)) + 10000 * CLng(Hour(Time))
		StoreScriptSetting("sLogFilename_CST",NoForbiddenFilenameCharacters(CStr(ntoday) + CStr (ntime) + Cstr(Cint(100*Rnd()))))

		sLogFilename_CST = GetScriptSetting("sLogFilename_CST","")

		If (sLogFilename_CST = "") Then
			PerformEvaluation = False
			Exit Function
		End If

		sLogFile_CST = GetProjectPath("Result") + sLogFilename_CST + ".log"
		sDataFile_CST = GetProjectPath("Result") + sLogFilename_CST + ".dat"

		SetLocale(&H409) ' temporarily change to US locale for file output
		Open sLogFile_CST For Output As #1

		If bWriteFile_CST_arbitr Then
			iDataFileID = OpenBufferedFile_LIB(sDataFile_CST, "Output")

			If (b1Dplot_CST_arbitr) Then
				BufferedFileWriteLine_LIB(iDataFileID, PP12(GetScriptSetting("coord"+CStr(iDir_CST_arbitr),"")) + PP12("Fieldvalue"))
				BufferedFileWriteLine_LIB(iDataFileID, PP12("--------------------") + PP12("--------------------"))
			Else
				BufferedFileWriteLine_LIB(iDataFileID, PP12(GetScriptSetting("coord1","")) + PP12(GetScriptSetting("coord2","")) + PP12(GetScriptSetting("coord3","")) + PP12("Fieldvalue") + PP12("VoxelSize"))
				BufferedFileWriteLine_LIB(iDataFileID, PP12("--------------------") + PP12("--------------------") + PP12("--------------------") + PP12("--------------------") + PP12("--------------------"))
			End If
		End If

		Print #1, vbCrLf + _
				"              Logfile of Field Evaluation" + vbCrLf + _
				"              ===========================" + vbCrLf + _
				vbCrLf + _
				"Calculation window:" + vbCrLf + _
				"==================="+ vbCrLf + _
				vbCrLf + _
				PP25L("Coordinate System") + ": " + aWCS(iWCS_CST_arbitr) + "   " + GetScriptSetting("aCoordsystem","") + vbCrLf + _
				PP25L("Dimension") + ": " + CStr(iDim_CST)+"D"
		If (iDim_CST = 1) Then	Print #1, PP25L("Direction (tangential)") + ": " + GetScriptSetting("coord"+CStr(iDir_CST_arbitr),"")
		If (iDim_CST = 2) Then	Print #1, PP25L("Face Normal")            + ": " + GetScriptSetting("coord"+CStr(iDir_CST_arbitr),"")

		If (iDim_CST > 0) Then	Print #1, PP25L("Stepsize")               + ": " + CStr(ddStep)

		Print #1, _
				vbCrLf + _
				PP8("") + PP12("low") + PP12("high") + PP12("stepsize") + vbCrLf + _
				PP8("") + PP12("----------------") + PP12("----------------") + PP12("----------------") + vbCrLf + _
				PP8(GetScriptSetting("coord1","")) + PP12(dUVWvalue(1,1))+PP12(dUVWvalue(1,2))+PP12(dUVWvalue(1,3)) + vbCrLf + _
		       	PP8(GetScriptSetting("coord2","")) + PP12(dUVWvalue(2,1))+PP12(dUVWvalue(2,2))+PP12(dUVWvalue(2,3)) + vbCrLf + _
		       	PP8(GetScriptSetting("coord3","")) + PP12(dUVWvalue(3,1))+PP12(dUVWvalue(3,2))+PP12(dUVWvalue(3,3)) + vbCrLf

		Print #1, _
				"Evaluated field, component and result value:" + vbCrLf + _
				"============================================"+ vbCrLf + _
				vbCrLf + _
				PP25L("Evaluated field") + ": "+ GetScriptSetting("ResultTreeNameArbitr","") + vbCrLf + _
				PP25L("Field component") + ": "+ GetScriptSetting("aComponent","") + vbCrLf + _
				PP25L("Complex")         + ": "+ GetScriptSetting("aComplex","") + vbCrLf + _
				PP25L("Result Value")    + ": "+ GetScriptSetting("a0DValue","Integral") + _
	        	IIf (Left((GetScriptSetting("a0DValue","Integral")),14) ="Integral f(x)-",":   f(x)="+ GetScriptSetting("Integ_fx","x"),"")+ vbCrLf+vbCrLf

		Print #1, _
				"Result Data:" + vbCrLf + _
				"============"+ vbCrLf

	End If	' bLogFileFirstEval

	' finally set min startpoint (1d => ONLY in IDIR, 2d=> ALL, but not for IDIR) at half stepsize

	Select Case iDim_CST
		Case 0 ' 0D
		Case 1 ' 1D
			If (Not b1Dplot_CST_arbitr) Then
				' NOT for 1d-plot
				dUVWvalue(iDir_CST_arbitr,1) = dUVWvalue(iDir_CST_arbitr,1) + 0.5 * dUVWvalue(iDir_CST_arbitr,3)
			End If

		Case 2 ' 2D
			For nIndex = 1 To 3
				If nIndex <> iDir_CST_arbitr Then
					dUVWvalue(nIndex,1) = dUVWvalue(nIndex,1) + 0.5 * dUVWvalue(nIndex,3)
				End If
			Next

		Case 3 ' 3D
			For nIndex = 1 To 3
				dUVWvalue(nIndex,1) = dUVWvalue(nIndex,1) + 0.5 * dUVWvalue(nIndex,3)
			Next
	End Select

	bMultiplyRadiusLater = False
	bMultiplyRsinThetaLater = False
	dVoxel_Unit = 1

	' angle-stepwidth needs to be taken in radian, new array dStepLength(1-3)
	For nIndex = 1 To 3
		dStepLength(nIndex) = dUVWvalue(nIndex,3)
		If bAngle(nIndex) Then
			dStepLength(nIndex) = dStepLength(nIndex) * lib_deg2rad
		End If
	Next

	Select Case iDim_CST
		Case 0 ' 0D
		Case 1 ' 1D
			dVoxel_Unit = dVoxel_Unit * dStepLength(iDir_CST_arbitr)

			If (icoordsystem_CST_arbitr <> 0) And (iDir_CST_arbitr = 2) Then ' rFz or rTf
				bMultiplyRadiusLater = True
			End If

			If (icoordsystem_CST_arbitr = 2) And (iDir_CST_arbitr = 3) Then ' rtF
				bMultiplyRsinThetaLater = True
			End If

		Case 2 ' 2D
			For nIndex = 1 To 3
				If nIndex <> iDir_CST_arbitr Then
					dVoxel_Unit = dVoxel_Unit * dStepLength(nIndex)
				End If
			Next

			If (icoordsystem_CST_arbitr = 1) And (iDir_CST_arbitr <> 2) Then ' rFz with normal=r or z
				bMultiplyRadiusLater = True
			End If

			If (icoordsystem_CST_arbitr = 2) Then ' rTF
				Select Case iDir_CST_arbitr
					Case 1 ' normal=r
						bMultiplyRadiusLater = True
						bMultiplyRsinThetaLater = True
					Case 2 ' normal=t
						bMultiplyRsinThetaLater = True
					Case 3 ' normal=f
						bMultiplyRadiusLater = True
				End Select
			End If
		Case 3 ' 3D
			For nIndex = 1 To 3
				dVoxel_Unit = dVoxel_Unit * dStepLength(nIndex)
			Next

			If (icoordsystem_CST_arbitr = 1) Then ' rFz
				bMultiplyRadiusLater = True
			End If

			If (icoordsystem_CST_arbitr = 2) Then ' rTF
				bMultiplyRadiusLater = True
				bMultiplyRsinThetaLater = True
			End If
	End Select

	dVoxel_SI = dVoxel_Unit * (Units.GetGeometryUnitToSI()) ^ iDim_CST

	If (bMultiplyRadiusLater Or bMultiplyRsinThetaLater) Then
		' save reference values
		dRef_Voxel_Unit = dVoxel_Unit
		dRef_Voxel_SI = dVoxel_SI
	End If

	If bDebugOutput Then ReportInformationToWindow("Voxels calculated after "+CSTr(Timer()-dStartTime)+" secs.")
	dStartTime = Timer()

	bLokVA_CST_arbitr = False		' rfw, rfz or rtf
	bLokVB_CST_arbitr = False		' uvw
	bLokVC_CST_arbitr = False		' xyz

	For nIndex = 0 To 2
		d_va_CST_arbitr(nIndex) = 0.0 	' component in rfw, rfz or rtf coordinates
		d_vb_CST_arbitr(nIndex) = 0.0 	' component in uvw coordinates
		d_vc_CST_arbitr(nIndex) = 0.0 	' component in xyz coordinates
	Next

	im1Dir_CST = iDir_CST_arbitr-1
	Select Case sComponent_CST_arbitr
		Case "Scalar"
			' vector transformation not required for scalar
		Case "Tangential"
			If (icoordsystem_CST_arbitr = 0 ) Then
				If (iWCS_CST_arbitr = 0) Then
					bLokVC_CST_arbitr = True
					d_vc_CST_arbitr(im1Dir_CST) = 1.0
				Else
					bLokVB_CST_arbitr = True
					d_vb_CST_arbitr(im1Dir_CST) = 1.0
				End If
			Else
				bLokVA_CST_arbitr = True
				d_va_CST_arbitr(im1Dir_CST) = 1.0
			End If
		Case "Normal"
			If (icoordsystem_CST_arbitr = 0 ) Then
				If (iWCS_CST_arbitr = 0) Then
					bLokVC_CST_arbitr = True
					d_vc_CST_arbitr(im1Dir_CST) = 1.0
				Else
					bLokVB_CST_arbitr = True
					d_vb_CST_arbitr(im1Dir_CST) = 1.0
				End If
			Else
				bLokVA_CST_arbitr = True
				d_va_CST_arbitr(im1Dir_CST) = 1.0
			End If
		Case "X"
			bLokVC_CST_arbitr = True
			d_vc_CST_arbitr(0) = 1.0
		Case "Y"
			bLokVC_CST_arbitr = True
			d_vc_CST_arbitr(1) = 1.0
		Case "Z"
			bLokVC_CST_arbitr = True
			d_vc_CST_arbitr(2) = 1.0
		Case "U"
			bLokVB_CST_arbitr = True
			d_vb_CST_arbitr(0) = 1.0
		Case "V"
			bLokVB_CST_arbitr = True
			d_vb_CST_arbitr(1) = 1.0
		Case "W"
			bLokVB_CST_arbitr = True
			d_vb_CST_arbitr(2) = 1.0
		Case "R"
			bLokVA_CST_arbitr = True
			d_va_CST_arbitr(0) = 1.0
		Case "F"
			bLokVA_CST_arbitr = True
			d_va_CST_arbitr(1) = 1.0
		Case "Theta"
			bLokVA_CST_arbitr = True
			d_va_CST_arbitr(1) = 1.0
		Case "Phi"
			bLokVA_CST_arbitr = True
			d_va_CST_arbitr(2) = 1.0

	End Select

	If (iWCS_CST_arbitr > 0 ) Then
		InitWCS(iWCS_CST_arbitr) ' sets numbers and martices of aWCS(iwcs)
		If (bLokVB_CST_arbitr) Then
			bLokVC_CST_arbitr = True
			Convert_Vector2Global d_vb_CST_arbitr, d_vc_CST_arbitr
		End If
	End If

	dSumVoxel_Unit_CST_arbitr = 0.0
	dSumIntegral_CST_arbitr = 0.0
	dMax_CST_arbitr = lib_rundef ' -1.23456e27
	dMin_CST_arbitr = - lib_rundef ' 1.23456e27

	If bDebugOutput Then ReportInformationToWindow("List prepared after " + CSTr(Timer()-dStartTime) + " secs")
	nNumberOfPoints = FillPointsInList()
	If bDebugOutput Then ReportInformationToWindow("List filled after " + CSTr(Timer()-dStartTime) + " secs")
	' If log file name exists, write data file with corresponding name.
	If (sLogFilename_CST <> "") And Not bPreviewOnly Then
		WriteXYZFile(GetProjectPath("Result") + sLogFilename_CST + ".xyz", nNumberOfPoints)
		If CInt(GetScriptSetting("CheckUseFixedPointlist","0"))=1 Then
			Dim sExportFolder As String, sFixedFileName As String
			sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d\"
			CST_MkDir sExportFolder
			sFixedFileName = sExportFolder + GetScriptSetting("FilePointlist","xyz-pointlist.txt")
			FileCopy GetProjectPath("Result") + sLogFilename_CST + ".xyz", sFixedFileName
			If Not bInfoAlreadyShown Then
				ReportInformation "Pointlist written: " + sFixedFileName
				bInfoAlreadyShown = True
			End If
		End If
	ElseIf bPreviewOnly Then
		WriteXYZFile(GetProjectPath("Result") + "points_preview.xyz", nNumberOfPoints)
		Exit Function
	End If
	If bDebugOutput Then ReportInformationToWindow("XYZ file written after " + CSTr(Timer()-dStartTime) + " secs")
	If (nNumberOfPoints < 0) Then
		ReportError(sWarningTemplateName + "Number of samples is negative, aborting.")
		PerformEvaluation = False
		Exit Function
	ElseIf ((dstepsize_CST=0.0) And (nNumberOfPoints < nMinimumNumberOfPoints) And (dAutostepReduction^iDim_CST < nMinimumNumberOfPoints) And (iDim_CST>0)) Then ' if auto step size and less than no. of min samples, re-run; limit dAutostepReduction to prevent endless loops
		If nNumberOfPoints = 0 Then
			dRefinementFactor = 2*dAutostepReduction
		Else
			dRefinementFactor = dAutostepReduction*((nMinimumNumberOfPoints*1.1)/(nNumberOfPoints))^(1/iDim_CST)
		End If
		If bWriteFile_CST_arbitr Then
			CloseBufferedFile_LIB(iDataFileID) ' Close buffered stream and reset related variables; will be opened during each recursion and start the files anew
		End If
		Close ' Close all remaining streams; they will be opened during each recursion and start the files anew
		SetLocale(nCurrentLocale) ' switch back to original locale
		PerformEvaluation = PerformEvaluation(dRefinementFactor)
		Exit Function
	ElseIf nNumberOfPoints = 0 Then
		If bLogFileFirstEval Then
			ReportError sWarningTemplateName + "No data point found. Reason might be too big a stepsize ("+Cstr(ddStep)+") to find any points in selected range/solids. Please manually set stepsize to a smaller value. In addition, choosing a smaller subvolume instead of max. range might speed up evaluation."
		End If
		cst_value_arbitr = lib_rundef
		Exit Function
	End If
	dStartTime = Timer()

	If Not GetTemplateAborted Then
		VectorPlot3D.CalculateList
		If bDebugOutput Then ReportInformationToWindow("List calculated after " + CSTr(Timer()-dStartTime) + " secs")
		dStartTime = Timer()
	Else
		PerformEvaluation = False
		Exit Function
	End If

	CalculatePointsInList(nNumberOfPoints)
	If bDebugOutput Then ReportInformationToWindow("Results calculated after " + CSTr(Timer()-dStartTime) + " secs")
	If bDebugOutput Then ReportInformationToWindow("--------------------------------------------------------------")
	dSumVoxel_SI = dSumVoxel_Unit_CST_arbitr * (Units.GetGeometryUnitToSI()) ^ iDim_CST
	dMeanCST = lib_rundef
	If (dSumVoxel_SI <> 0) Then
		dMeanCST = dSumIntegral_CST_arbitr/dSumVoxel_SI
	End If

	dNorming = 0.5*( Abs(dMax_CST_arbitr) + Abs(dMin_CST_arbitr) )
	dDeviCST = dMax_CST_arbitr - dMin_CST_arbitr
	If (dNorming <> 0) Then
		dDeviCST = dDeviCST / dNorming
	End If

	If bLogFileFirstEval Then

		Print #1, " NData Points = "+CStr(nNumberOfPoints)
		Print #1, ""

		If bSkippedPointsDueToClamping Then
			Print #1, " *** Due to clamping field values not all "+CStr(nNumberOfPoints)+" data points were considered for evaluation."
			Print #1, " *** Statistics below only considers points and values, fitting in the clamped ranged"
			Print #1, ""
			ReportInformation("Evaluate Field in arbitrary Coordinates: Due to clamping field values not all data points were considered for evaluation.")
		End If

		If (nNumberOfPoints = 1) Then
			Print #1, " Field Value  = "+PP12(dSumIntegral_CST_arbitr)
		Else
			Print #1, " Max. value   = "+PP12(dMax_CST_arbitr)+" at ("+PP12(dMax_X_CST_arbitr)+"/"+PP12(dMax_Y_CST_arbitr)+"/"+PP12(dMax_Z_CST_arbitr)+")"
			Print #1, " Min. value   = "+PP12(dMin_CST_arbitr)+" at ("+PP12(dMin_X_CST_arbitr)+"/"+PP12(dMin_Y_CST_arbitr)+"/"+PP12(dMin_Z_CST_arbitr)+")"
			Print #1, " Mean value   = "+PP12(dMeanCST)
			Print #1, " Deviation    = "+PP12(dDeviCST)
			Print #1, ""
			Print #1, " Integral     = "+PP12(dSumIntegral_CST_arbitr)
			Print #1, " Volume (Unit)= "+PP12(dSumVoxel_Unit_CST_arbitr) + " " + Units.GetGeometryUnit + "^" + CStr(iDim_CST)
			Print #1, " Volume (SI)  = "+PP12(dSumVoxel_SI) + " m^" + CStr(iDim_CST)+vbCrLf
		End If

		If (bSolids_CST) Then
			' now printout seperate solid results into print file
			Print #1, PP20L("------------------------------------") + _
					PP10("--------------") + PP12("-----------------------") + PP12("-----------------------") + _
					PP12("-----------------------") + PP12("-----------------------") + PP12("-----------------------")
			Print #1, PP20L("Solid Name") + PP10("NData") + PP12("min") + PP12("max") + PP12("mean") + PP12("Integral") + PP12("Volume")
			Print #1, PP20L("------------------------------------") + _
					PP10("--------------") + PP12("-----------------------") + PP12("-----------------------") + _
					PP12("-----------------------") + PP12("-----------------------") + PP12("-----------------------")
			'
			Dim dmean_cst As Double, dVol_tmp As Double
			For iSolid_CST = 1 To nSolids_CST

				dVol_tmp = dSolid_Volume_CST_arbitr(iSolid_CST-1)

				If (dVol_tmp > 0) Then
					dmean_cst = dSolid_Integral_CST_arbitr(iSolid_CST-1) / dVol_tmp
				Else
					dmean_cst = lib_rundef
				End If

				Print #1, PP20L(aSolidArray_CST(iSolid_CST-1)) + PP10(nSolid_Data_CST_arbitr(iSolid_CST-1)) + _
				PP12(dSolid_Minimum_CST_arbitr(iSolid_CST-1)) + PP12(dSolid_Maximum_CST_arbitr(iSolid_CST-1)) + _
				PP12(dmean_cst) + PP12(dSolid_Integral_CST_arbitr(iSolid_CST-1)) + PP12(dVol_tmp)
			Next
			Print #1, PP20L("------------------------------------") + _
					PP10("--------------") + PP12("-----------------------") + PP12("-----------------------") + _
					PP12("-----------------------") + PP12("-----------------------") + PP12("-----------------------")
			'
			Print #1, PP20L("Total") + PP10(nNumberOfPoints) + PP12(dMin_CST_arbitr) + PP12(dMax_CST_arbitr) + _
					PP12(dMeanCST) + PP12(dSumIntegral_CST_arbitr) + PP12(dSumVoxel_SI)
			'
			Print #1, PP20L("------------------------------------") + _
					PP10("--------------") + PP12("-----------------------") + PP12("-----------------------") + _
					PP12("-----------------------") + PP12("-----------------------") + PP12("-----------------------")
		End If

	End If	' bLogFileFirstEval

	cst_value_arbitr = lib_rundef

	' Store all values for further use in EvaluateMultiple0D
	StoreScriptSetting("IntegralM0D", dSumIntegral_CST_arbitr)
	StoreScriptSetting("MaximumM0D", dMax_CST_arbitr)
	StoreScriptSetting("MaximumXM0D", dMax_X_CST_arbitr)
	StoreScriptSetting("MaximumYM0D", dMax_Y_CST_arbitr)
	StoreScriptSetting("MaximumZM0D", dMax_Z_CST_arbitr)
	StoreScriptSetting("MinimumM0D", dMin_CST_arbitr)
	StoreScriptSetting("MinimumXM0D", dMin_X_CST_arbitr)
	StoreScriptSetting("MinimumYM0D", dMin_Y_CST_arbitr)
	StoreScriptSetting("MinimumZM0D", dMin_Z_CST_arbitr)
	StoreScriptSetting("MeanM0D", dMeanCST)
	StoreScriptSetting("DeviationM0D", dDeviCST)
	StoreScriptSetting("VoxelSumM0D", dSumVoxel_SI)

	Select Case Left(sActionCST,4)

		Case 	"Fiel" ' 	0d  "Field Value"
			cst_value_arbitr = dSumIntegral_CST_arbitr

		Case 	"Inte" ' 	"Integral-1D","Integral-2D","Integral-3D"
			cst_value_arbitr = dSumIntegral_CST_arbitr

		Case 	"Maxi" ' 	"Maximum-1D","Maximum-2D","Maximum-3D"
			cst_value_arbitr = dMax_CST_arbitr
		Case 	"Mini" ' 	"Minimum-1D","Minimum-2D","Minimum-3D"
			cst_value_arbitr = dMin_CST_arbitr
		Case 	"Mean" ' 	"Mean Value-1D","Mean Value-2D","Mean Value-3D"
			cst_value_arbitr = dMeanCST

		Case 	"Devi" ' 	"Deviation-1D","Deviation-2D","Deviation-3D"
			cst_value_arbitr = dDeviCST

		Case 	"Leng" ' 	"Length-1D"
			cst_value_arbitr = dSumVoxel_SI

		Case 	"Area" ' 	"Area-2D"
			cst_value_arbitr = dSumVoxel_SI

		Case 	"Volu" ' 	"Volume-3D"
			cst_value_arbitr = dSumVoxel_SI

	End Select

'yyx
	If bLogFileFirstEval Then
		If bWriteFile_CST_arbitr Then
			CloseBufferedFile_LIB(iDataFileID)
		End If
		Close #1 ' close log file, hardwired stream number #1
		SetLocale(nCurrentLocale) ' switch back to original locale
	End If

End Function ' PerformEvaluation

