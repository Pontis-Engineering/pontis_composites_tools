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

Dim femap as femap.model
Dim oMatl As Object
Dim oLayup As Object
Dim oProp As Object

Dim appExcel As Object
Dim wb As Object
Dim ws As Object

Sub Main

'---set objects

	Set femap = feFemap()
	Set oMatl = femap.feMatl
	Set oLayup = femap.feLayup
	Set oProp = femap.feProp

	Dim a2_data As Variant

	Call femap_oProp_extract_(a2_data)

'--- open xls book

	Set appExcel = CreateObject("Excel.Application") 'New Excel.Application
	Set wb = appExcel.Workbooks.Add
	Set ws = wb.ActiveSheet
	ws.Name = "property"

	a2_data = appExcel.WorksheetFunction.Transpose(a2_data)

	Dim n_row As Long, n_col As Long
	n_row =UBound(a2_data,1)-LBound(a2_data,1)+1
	n_col = UBound(a2_data,2)-LBound(a2_data,2)+1

'--- write out data

	ws.Range("A1").Offset(0, 0).Resize(n_row, n_col) = a2_data
	appExcel.Visible = True
	AppActivate appExcel.ActiveWindow.Caption

end Sub

Sub femap_oProp_extract_(a2_data As Variant)

'------ set variables
	Dim i As Long, j As Long, k As Long, s_txt As String, a_temp As Variant, rc As Integer

'------ main script

	Dim n_prop As Integer
	n_prop = oProp.countset

	ReDim a_out(0 To n_prop) As Variant
	Dim a_prop(1 To 9) As Variant

	a_prop(1) = "prop id"
	a_prop(2) = "prop name"
	a_prop(3) = "matl id"
	a_prop(4) = "color"
	a_prop(5) = "layer"
	a_prop(6) = "type"
	a_prop(7) = "type name"
	a_prop(8) = "layup id"
	a_prop(9) = "use"

	a_out(0) = a_prop

	rc = oProp.Reset

	Dim vpval As Variant

	For i = 1 To n_prop

		rc = oProp.Next

		a_prop(1) = oProp.ID
		a_prop(2) = oProp.Title
		a_prop(3) = oProp.matlID
		a_prop(4) = oProp.Color
		a_prop(5) = oProp.layer

		a_prop(6) = oProp.Type

		If a_prop(6) = 5 Then a_prop(7) = "beam (5)"
		If a_prop(6) = 29 Then a_prop(7) = "rigid (29)"
		If a_prop(6) = 21 Then a_prop(7) = "laminate (21)"
		If a_prop(6) = 22 Then a_prop(7) = "laminate (22)"
		If a_prop(6) = 27 Then a_prop(7) = "mass (27)"
		If a_prop(6) = 17 Then a_prop(7) = "plate (17)"
		If a_prop(6) = 25 Then a_prop(7) = "solid (25)"

		a_prop(8) = oProp.layupID
		a_prop(9) = 1
		a_out(i) = a_prop
	Next

	a2_data = a3_to_a2(a_out, "col1")

'------ end main script

End Sub

Function a3_to_a2(a3 As Variant, sMeth As Variant)

	Dim a_temp As Variant, a_2D As Variant, s_txt As String
	Dim i As Long, j As Long, k As Long

	If Left(sMeth, 3) = "col" Or Left(sMeth, 3) = "COL" Then s_txt = "col" 'stack data to the right
	If Left(sMeth, 3) = "row" Or Left(sMeth, 3) = "ROW" Then s_txt = "row" 'stack data down

	Dim rowS As Integer, rowF As Integer         'new 2D array size
	Dim colS As Integer, colF As Integer         'new 2D array size
	Dim n_row As Long, n_col As Long
	Dim i0 As Integer, j0 As Long
	Dim n_3D_start As Long, n_3D_finish As Long
	n_3D_start = LBound(a3)
	n_3D_finish = UBound(a3)

'check if first array is an array
	a_temp = a3(LBound(a3))
	If IsArray(a_temp) = False Then
		Dim a_temp0(0) As Variant
		a_temp0(0) = a3(0)
		a3(0) = a_temp0
	End If

	Dim a_output As Variant

	For k = n_3D_start To n_3D_finish
		a_temp = a3(k)

		On Error Resume Next
		i = UBound(a_temp, 2)
		If Err.Number > 0 Then

			ReDim array_2D(LBound(a_temp, 1) To UBound(a_temp, 1), 1 To 1) As Variant

			For i = LBound(a_temp, 1) To UBound(a_temp, 1)
				array_2D(i, 1) = a_temp(i)
			Next

			a_temp = array_2D

		End If
		On Error GoTo 0

		a3(k) = a_temp

		i0 = UBound(a_temp) - LBound(a_temp) + 1
		j0 = UBound(a_temp, 2) - LBound(a_temp, 2) + 1

		n_row = i0 + n_row
		n_col = j0 + n_col
	Next

	If Mid(sMeth, 1, 3) = "col" Then n_row = i0
	If Mid(sMeth, 1, 3) = "row" Then n_col = j0
	ReDim a_output(1 To n_row, 1 To n_col) As Variant

	Dim count As Long
	count = 0

	Dim i1 As Long, j1 As Long

	For k = n_3D_start To n_3D_finish
		a_temp = a3(k)
		i1 = 0

		For i = LBound(a_temp, 1) To UBound(a_temp, 1)
			i1 = i1 + 1
			j1 = 0

			For j = LBound(a_temp, 2) To UBound(a_temp, 2)
				j1 = j1 + 1

				If Mid(sMeth, 1, 3) = "col" Then a_output(i1, count + j1) = a_temp(i, j)
				If Mid(sMeth, 1, 3) = "row" Then a_output(count + i1, j1) = a_temp(i, j)
			Next
		Next

		If Mid(sMeth, 1, 3) = "col" Then count = count + UBound(a_temp, 2) - LBound(a_temp, 2) + 1
		If Mid(sMeth, 1, 3) = "row" Then count = count + UBound(a_temp, 1) - LBound(a_temp, 1) + 1
	Next

	a3_to_a2 = a_output

End Function

