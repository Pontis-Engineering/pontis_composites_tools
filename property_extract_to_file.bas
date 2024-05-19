'*********************************************************************************
' Project: Pontis Composite Tools
' Module: Property Extract to File

' Description: Module exporting created Property objects from Femap to
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
	a_prop(3) = "mtrl id"
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

	a2_data = a3_to_a2_(a_out, "")

'------ end main script

End Sub

Function a3_to_a2_(a3 As Variant, sMeth As Variant)

  Dim a_temp As Variant
  Dim i As Long, j As Long, k As Long, count As Long

  Dim n_a2_row_start As Long, n_a2_row_end As Long
  Dim n_a2_col_start As Long, n_a2_col_end As Long

'--- # fill array assuming 2d arrays

    On Error GoTo ErrHandler

    n_a2_row_start = 0

    For i = LBound(a3) To UBound(a3)

        a_temp = a3(i)

        n_a2_row_end = n_a2_row_end + UBound(a_temp) - LBound(a_temp) + 1

    Next

    a_temp = a3(n_a2_row_start)

    ReDim a2(n_a2_row_start To n_a2_row_end, LBound(a_temp, 2) To UBound(a_temp, 2)) As Variant

    count = n_a2_row_start

    For k = LBound(a3) To UBound(a3)

        a_temp = a3(k)

        j = LBound(a_temp, 2)

        For i = LBound(a_temp) To UBound(a_temp)

            For j = LBound(a_temp, 2) To UBound(a_temp, 2)

                a2(count, j) = a_temp(i, j)

            Next

            count = count + 1

        Next

    Next

    a3_to_a2_ = a2

    Exit Function


ErrHandler:

'--- # fill array assuming 1d arrays

    a_temp = a3(LBound(a3, 1))

    n_a2_row_start = LBound(a3)

    n_a2_row_end = UBound(a3)

    ReDim a2(n_a2_row_start To n_a2_row_end, LBound(a_temp, 1) To UBound(a_temp, 1)) As Variant

    For i = LBound(a2) To UBound(a2)

        a_temp = a3(i)

        For j = LBound(a_temp) To UBound(a_temp)

            a2(i, j) = a_temp(j)

        Next

        count = count + 1

    Next

    a3_to_a2_ = a2


End Function
