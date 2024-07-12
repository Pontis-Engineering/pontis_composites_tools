'*********************************************************************************
' Project: Pontis Composite Tools
' Module: Material Create from File


' Description: Module importing created Material objects from Excel spreadsheet to
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

Dim ExcelFileName As String
Dim ExcelApp As Object
Dim ExcelWorkbook As Object
Dim ExcelWorksheet As Object
Dim ExcelDataArray As Variant

Dim i As Integer, j As Integer

Sub Main()

'---set objects

	Set FemapApp = feFemap()
	Set oMatl = FemapApp.feMatl

	Dim a2_data As Variant


	Call GetData(a2_data,"material")
	Call femap_oMatl_create_(a2_data)

end Sub

sub femap_oMatl_create_(a2_data as Variant)

'------ set variables
	Dim i As Integer, j As Long, k As Long, s_txt As String, a_temp As Variant
	Dim vmval(199) As Variant

'------ start main script

'--- tidy array

	Call a2_tidy_(a2_data, a2_look_("use", a2_data, -1), "")

'--- find col num
	Dim n_id As Long, n_title As Long, n_type As Long, n_den As Long, n_use As Long, _
	n_E11 As Long, n_E22 As Long, n_E33 As Long, _
	n_G12 As Long, n_G23 As Long, n_G31 As Long, _
	n_nu12 As Long, n_nu23 As Long, n_nu13 As Long, _
	n_s11t As Long, n_s22t As Long, n_s11c As Long, n_s22c As Long, n_s12 As Long

	n_id = a2_look_("mtrl id", a2_data, -1)
	n_title = a2_look_("mtrl name", a2_data, -1)
	n_type = a2_look_("type id", a2_data, -1)

	n_den = a2_look_("den", a2_data, -1)
	n_use = a2_look_("use", a2_data, -1)

	n_E11 = a2_look_("E11", a2_data, -1)
	n_E22 = a2_look_("E22", a2_data, -1)
	n_E33 = a2_look_("E33", a2_data, -1)

	n_G12 = a2_look_("G12", a2_data, -1)
	n_G23 = a2_look_("G23", a2_data, -1)
	n_G31 = a2_look_("G31", a2_data, -1)

	n_nu12 = a2_look_("nu12", a2_data, -1)
	n_nu23 = a2_look_("nu23", a2_data, -1)
	n_nu13 = a2_look_("nu13", a2_data, -1)

	n_s11t = a2_look_("s11t", a2_data, -1)
	n_s22t = a2_look_("s22t", a2_data, -1)
	n_s11c = a2_look_("s11c", a2_data, -1)
	n_s22c = a2_look_("s22c", a2_data, -1)
	n_s12 = a2_look_("s12", a2_data, -1)

'--- apply data into femap

	For i = LBound(a2_data) + 1 To UBound(a2_data)

		j = a2_data(i, n_id)

		oMatl.Get (j)
		oMatl.vmval = vmval
		oMatl.layer = 1  'layer
		oMatl.Color = 55  'default

		If n_type > 0 Then oMatl.Type = a2_data(i, n_type)   'Type
		If n_title > 0 Then oMatl.Title = a2_data(i, n_title)   'Title
		If n_E11 > 0 Then oMatl.Ex = a2_data(i, n_E11)   'Ex
		If n_E22 > 0 Then oMatl.Ey = a2_data(i, n_E22)   'Ey
		If n_E33 > 0 Then oMatl.Ez = a2_data(i, n_E33)   'Ez
		If n_G12 > 0 Then oMatl.Gx = a2_data(i, n_G12)   'Gx
		If n_G23 > 0 Then oMatl.Gy = a2_data(i, n_G23)   'Gy
		If n_G31 > 0 Then oMatl.Gz = a2_data(i, n_G31)   'Gz
		If n_nu12 > 0 Then oMatl.NUxy = a2_data(i, n_nu12)   'NUxy
		If n_nu23 > 0 Then oMatl.NUyz = a2_data(i, n_nu23)   'NUyz
		If n_nu13 > 0 Then oMatl.NUxz = a2_data(i, n_nu13)   'NUxz
		If n_den > 0 Then oMatl.density = a2_data(i, n_den) * 0.000000001 'density assume input is kg/m^3
		If n_s11t > 0 Then oMatl.TensionLimit1 = a2_data(i, n_s11t) 'TensionLimit1
		If n_s22t > 0 Then oMatl.TensionLimit2 = a2_data(i, n_s22t) 'TensionLimit2
		If n_s11c > 0 Then oMatl.CompressionLimit1 = a2_data(i, n_s11c) 'CompressionLimit1
		If n_s22c > 0 Then oMatl.CompressionLimit2 = a2_data(i, n_s22c) 'CompressionLimit2
		If n_s12 > 0 Then oMatl.ShearLimit = a2_data(i, n_s12) 'ShearLimit

		oMatl.Put (j)
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

			If num <> 0 And num <> -1 Then a2_look_ = a2_data(num, j)    ''And IsEmpty(a2_look_) = False

			If num = 0 And IsEmpty(a2_look_) = False Then a2_look_ = "col" & a2_look_

			If num = -1 And IsEmpty(a2_look_) = False Then a2_look_ = a2_look_

			If IsEmpty(a2_look_) = False Then Exit Function

		End If

	Next

End Function


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
