'*********************************************************************************
' Project: Pontis Composite Tools
' Module: Material Export
' Description: Module exporting created material objects from Femap to
' new Excel spreadsheet.
'
' Authors:
'   - Darren Ellam <del@pontis-engineering.com>
'
' Copyright © 2024 Pontis Engineering. All rights reserved.
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' This code is provided "as is" without warranty of any kind, either expressed or
' implied, including but not limited to the implied warranties of merchantability
' and/or fitness for a particular purpose.
'
' You may modify this code for your own personal or commercial use, but you may not
' distribute or publish it without prior written permission from Pontis Engineering.
'*********************************************************************************

Option Explicit

Dim FemapApp As femap.model
Dim oLayup As Object

Dim ExcelFileName As String
Dim ExcelApp As Object		' Excel.Application
Dim ExcelWorkbook As Object 	' Excel.Workbook
Dim ExcelWorksheet As Object	' Excel.Worksheet
Dim ExcelDataArray As Variant

Dim i As Integer, j As Integer


Sub Main()

'---set objects

	Set FemapApp = feFemap()
	Set oLayup = FemapApp.feLayup
	Dim a2_data As Variant

	Call GetData(a2_data, "layup")
	Call femap_oLayup_create_(a2_data)

end Sub

Sub femap_oLayup_create_(a2_data As Variant)

'------ set variables
	Dim i As Integer, j As Long, k As Long, s_txt As String, a_temp As Variant
	Dim rc As Integer
	Dim nNumPly As Long, nMatlID As Variant, dThickness As Variant, dAngle As Variant, nGlobalPly As Variant
	Dim n_id As Long, s_title As String

'------ start main script

	Call a2_tidy_(a2_data, a2_look_("use", a2_data, -1), "")
	Call a2_to_a3_(a2_data, 1, "text")

	For i = 1 To UBound(a2_data)
		a_temp = a2_data(i)
		n_id = a_temp(2, a2_look_("layup id", a_temp, -1))
		s_title = a_temp(2, a2_look_("layup name", a_temp, -1))
		nNumPly = UBound(a_temp) - 1
		ReDim a_temp2(0 To nNumPly - 1) As Variant
		nMatlID = a_temp2
		dThickness = a_temp2
		dAngle = a_temp2
		nGlobalPly = a_temp2

		For j = 1 To UBound(a_temp) - 1
			nMatlID(j - 1) = a_temp(j + 1, a2_look_("mtrl id", a_temp, -1))
			dThickness(j - 1) = a_temp(j + 1, a2_look_("ply t", a_temp, -1))
			dAngle(j - 1) = a_temp(j + 1, a2_look_("deg", a_temp, -1))
			nGlobalPly(j - 1) = a_temp(j + 1, a2_look_("gply#", a_temp, -1))
		Next

		rc = oLayup.Get(n_id)
		oLayup.Title = s_title

		rc = oLayup.SetAllPly(nNumPly, nMatlID, dThickness, dAngle, nGlobalPly)
		rc = oLayup.Put(n_id)

	Next
'------ end main script
End Sub


Function GetData(a2_data As Variant, ws_name As String)

	Dim rc As Integer

' Prompt user to select Excel file
	rc = FemapApp.feFileGetName("Select Excel File", "Excel Files", "*.xl*", True, ExcelFileName)
	Call HandleReturnCode(rc)

' Create an instance of Excel
	Set ExcelApp = CreateObject("Excel.Application")

' Open the selected Excel file
	Set ExcelWorkbook = ExcelApp.Workbooks.Open(ExcelFileName)

' Set the active worksheet
	ExcelWorkbook.Worksheets(ws_name).Activate
	Set ExcelWorksheet = ExcelWorkbook.Activesheet

' Retrieve data from the used range of the worksheet
	ExcelDataArray = ExcelWorksheet.UsedRange.Value

	a2_data = ExcelDataArray

' Close Excel
	ExcelWorkbook.Close False
	ExcelApp.Quit

' Release Excel objects
	Set ExcelWorksheet = Nothing
	Set ExcelWorkbook = Nothing
	Set ExcelApp = Nothing

End Function

Sub HandleReturnCode(rc)
' Get reference to Femap application
	Set FemapApp = GetObject(, "femap.model")

' Handle return code
	Select Case rc
	Case 0, 2
' User pressed Cancel, exit the program
		FemapApp.feAppMessage(FCM_NORMAL, "User pressed Cancel, exiting...")
		End
	End Select
End Sub


Sub a2_tidy_(a2_data As Variant, i_num As Integer, s_misc As String)

	Dim a_temp As Variant
	a_temp = a2_data

	Dim i As Long, j As Long, count As Long, k As Long

	If i_num = -1 Then k = UBound(a2_data, 2)
	If i_num > -1 Then k = i_num

	For i = LBound(a2_data, 1) To UBound(a2_data, 1)
		If a2_data(i, k) <> 0 And a2_data(i, k) <> "" Then count = count + 1
	Next

	ReDim a_temp2(LBound(a2_data, 1) To count, LBound(a2_data, 2) To UBound(a2_data, 2)) As Variant

	count = LBound(a2_data, 1)

	For i = LBound(a2_data, 1) To UBound(a2_data, 1)

		If a2_data(i, k) <> 0 And a2_data(i, k) <> "" Then
			For j = LBound(a2_data, 2) To UBound(a2_data, 2)
				a_temp2(count, j) = a_temp(i, j)
			Next

			count = count + 1

		End If
	Next

	a2_data = a_temp2

End Sub

Function a2_look_(s_keyword As Variant, a2_data As Variant, num As Long) '24p1

'------ set variables
	Dim i As Long, j As Long
	a2_look_ = -1 'default

	Dim a_row1 As Variant, a_col1 As Variant
	Dim i_row As Integer, i_col As Integer
	Dim i_a1D As Integer

'--- checks

	If num > UBound(a2_data) - LBound(a2_data) + 1 Then
		Exit Function
	End If

	On Error Resume Next

	i = UBound(a2_data, 2)

	If Err.Number > 0 Then ' check if a 1D array
		Debug.Print "Not a 2D array"
		Exit Function
	End If

'--- search

	i = LBound(a2_data, 1)

	For j = LBound(a2_data, 2) To UBound(a2_data, 2)
		If a2_data(i, j) = s_keyword Then
			a2_look_ = j

			If num <> 0 And num <> -1 Then a2_look_ = a2_data(num, j)
			If num = 0 And IsEmpty(a2_look_) = False Then a2_look_ = "col" & a2_look_
			If num = -1 And IsEmpty(a2_look_) = False Then a2_look_ = a2_look_
			If IsEmpty(a2_look_) = False Then Exit Function

		End If
	Next

End Function

Sub a2_to_a3_(a_temp As Variant, col As Long, s_meth As String)

	Dim n_col As Integer, i As Long, j As Long
	Dim a_temp2 As Variant, a_col As Variant

	n_col = col
	a_temp2 = a_temp

	For i = 1 To UBound(a_temp2)
		If IsNumeric(a_temp2(i, n_col)) = False Then a_temp2(i, n_col) = ""
	Next

	For i = 2 To UBound(a_temp2)
		If a_temp2(i, n_col) = a_temp2(i - 1, n_col) Then a_temp2(i - 1, n_col) = ""
	Next

	Call a2_tidy_(a_temp2, n_col, "")

	ReDim a_col(LBound(a_temp2, 1) To UBound(a_temp2, 1)) As Variant

	For i = LBound(a_temp2, 1) To UBound(a_temp2, 1)
		a_col(i) = a_temp2(i, n_col)
	Next

	ReDim a_out(0 To UBound(a_col)) As Variant

	For i = 1 To UBound(a_out)
		a_out(i) = a_temp
	Next

	For i = 1 To UBound(a_out)
		a_temp2 = a_out(i)

		If s_meth <> "text" Then
			If a_temp2(1, n_col) <> a_col(i) Then a_temp2(1, 1) = ""
		End If

		For j = 2 To UBound(a_temp2)
			If a_temp2(j, n_col) <> a_col(i) Then a_temp2(j, 1) = ""
		Next

		Call a2_tidy_(a_temp2, 1, "")

		a_out(i) = a_temp2
	Next

	a_out(0) = a_col
	a_temp = a_out

End Sub
