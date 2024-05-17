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

	Dim a2_data As Variant

	Call femap_oLayup_read_(a2_data)

'--- open xls book

	Set appExcel = CreateObject("Excel.Application")
	Set wb = appExcel.Workbooks.Add
	Set ws = wb.ActiveSheet
	ws.Name = "layup"

	Dim n_row As Long, n_col As Long

	n_row =UBound(a2_data,1)-LBound(a2_data,1)+1
	n_col = UBound(a2_data,2)-LBound(a2_data,2)+1

'--- write out data

	ws.Range("A1").Offset(0, 0).Resize(n_row, n_col) = a2_data
	appExcel.Visible = True

	AppActivate appExcel.ActiveWindow.Caption

end Sub

Sub femap_oLayup_read_(a2_data As Variant)

'------ set variables
	Dim i As Long, j As Long, k As Long, s_txt As String, a_temp As Variant, rc As Integer

'------ main script

	Dim n_layup As Long
	n_layup = oLayup.countset

	ReDim a_out(0 To n_layup) As Variant

	rc = oLayup.Reset

	ReDim a_layup(0 To 0, 1 To 9) As Variant

	k = k + 1
	a_layup(0, k) = "layup id"
	k = k + 1
	a_layup(0, k) = "layup name"
	k = k + 1
	a_layup(0, k) = "gply#"
	k = k + 1
	a_layup(0, k) = "ply#"
	k = k + 1
	a_layup(0, k) = "matl id"
	k = k + 1
	a_layup(0, k) = "matl name"
	k = k + 1
	a_layup(0, k) = "ply t"
	k = k + 1
	a_layup(0, k) = "deg"
	k = k + 1
	a_layup(0, k) = "use"

	a_out(0) = a_layup


	Dim vpval As Variant
	Dim n_id As Long, s_title As String, n_plys As Integer
	Dim vmatlID As Variant, vglobalply As Variant, vthickness  As Variant, vangle As Variant
	Dim count As Long

	For i = 1 To n_layup
		rc = oLayup.Next
		n_id = oLayup.ID
		s_title = oLayup.Title
		n_plys = oLayup.NumberOfPlys
		count = count + n_plys
		vmatlID = oLayup.vmatlID
		vglobalply = oLayup.vglobalply
		vthickness = oLayup.vthickness
		vangle = oLayup.vangle

		ReDim a_temp(1 To n_plys, 1 To UBound(a_layup, 2)) As Variant

		For j = 1 To n_plys
			k = 0
			k = k + 1
			a_temp(j, k) = n_id
			k = k + 1
			a_temp(j, k) = s_title
			k = k + 1
			a_temp(j, k) = vglobalply(j - 1)
			k = k + 1
			a_temp(j, k) = j
			k = k + 1
			a_temp(j, k) = vmatlID(j - 1)
			rc = oMatl.Get(a_temp(j, k))
			k = k + 1
			a_temp(j, k) = oMatl.Title
			k = k + 1
			a_temp(j, k) = vthickness(j - 1)
			k = k + 1
			a_temp(j, k) = vangle(j - 1)
			k = k + 1
			a_temp(j, k) = 1
		Next

		a_out(i) = a_temp
	Next

	a2_data = a3_to_a2(a_out, "row")

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
