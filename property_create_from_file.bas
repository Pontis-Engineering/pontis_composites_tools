'*********************************************************************************
' Project: Pontis Composite Tools
' Module: Property Create from File

' Description: Module importing created Property objects from Excel spreadsheet to

' Femap model.
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
Dim oMatl As Object
Dim oLayup As Object
Dim oProp As Object

Dim ExcelFileName As String
Dim ExcelApp As Object			  ' Excel.Application
Dim ExcelWorkbook As Object 	' Excel.Workbook
Dim ExcelWorksheet As Object 	' Excel.Worksheet
Dim ExcelDataArray As Variant

Dim i As Integer, j As Integer

Sub Main()

'---set objects

	Set FemapApp = feFemap()
	Set oMatl = FemapApp.feMatl
	Set oLayup = FemapApp.feLayup
	Set oProp = FemapApp.feProp

	Dim a2_data As Variant

	Call GetData(a2_data, "property")
	Call femap_oProp_create_(a2_data)

end Sub

Sub femap_oProp_create_(a2_data As Variant)

	Dim i As Long, j As Long, k As Long, s_txt As String, a_temp As Variant
	Dim rc As Integer
	Dim n_id As Long, n_layup As Long, n_prop As Long, vpval As Variant

'------ main script

'--- tidy array

	Call a2_tidy_(a2_data, a2_look_("use", a2_data, -1), "")

'--- loop and create properties

	n_prop = UBound(a2_data) - LBound(a2_data)
	Dim n_layer As Integer

	For i = LBound(a2_data) + 1 To UBound(a2_data)

		n_id = a2_data(i, 1)
		rc = oProp.Get(n_id)
		n_layer = a2_look_("layer", a2_data, i)

		If n_layer = 0 Then n_layer = 1

		vpval = oProp.vpval
		oProp.Title = a2_look_("prop name", a2_data, i)
		oProp.matlID = a2_look_("matl id", a2_data, i)
		oProp.Color = 110
		oProp.layer = n_layer
		oProp.Type = a2_look_("type", a2_data, i)
		oProp.layupID = a2_look_("layup id", a2_data, i)
		rc = oProp.Put(n_id)

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

	Debug.Print(ExcelWorksheet.Name)
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


Sub a2_tidy_(a2_data As Variant, i_num As Integer, s_misc As String) '24p1

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
		Debug.Print "Provided index exceeds maximum array size"
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

			If num <> 0 And num <> -1 Then a2_look_ = a2_data(num, j)    'And IsEmpty(a2_look_) = False
			If num = 0 And IsEmpty(a2_look_) = False Then a2_look_ = "col" & a2_look_
			If num = -1 And IsEmpty(a2_look_) = False Then a2_look_ = a2_look_
			If IsEmpty(a2_look_) = False Then Exit Function

		End If

	Next

End Function
