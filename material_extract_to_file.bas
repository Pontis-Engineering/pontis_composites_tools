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

Dim appExcel As Object
Dim wb As Object
Dim ws As Object

Sub Main

'---set objects

	Set femap = feFemap()
	Set oMatl = femap.feMatl

	Dim a2_data As Variant

	Call femap_oMatl_extract_(a2_data)

'--- open xls book

	Set appExcel = CreateObject("Excel.Application") 'New Excel.Application
	Set wb = appExcel.Workbooks.Add
	Set ws = wb.ActiveSheet
	ws.Name = "material"

	Dim n_row As Long, n_col As Long

	n_row =UBound(a2_data,1)-LBound(a2_data,1)+1

	n_col = UBound(a2_data,2)

'--- write out data

	ws.Range("A1").Offset(0, 0).Resize(n_row, n_col) = a2_data
	appExcel.Visible = True
	AppActivate appExcel.ActiveWindow.Caption

end Sub

Sub femap_oMatl_extract_(a_matl As Variant)

'------ req. functions

'------ set variables
	Dim i As Long, j As Long, s_txt As String

	Dim values As Variant, element As Variant

	Dim appExcel As Object
	Dim wb As Object
	Dim ws As Object

'------ main script

'--- define 1st row an # of parameters (n_par)

	Dim n_par As Long
	Dim a0 As Variant

	a0 = Array("mtrl id", "mtrl name", "type id", "den", "E11", "E22", "G12", "nu12", "s11t", "s22t", "s11c", "s22c", "s12", "use")

	n_par = UBound(a0) + 1

'--- define array to store data

	Dim n_mat As Long

	n_mat = oMatl.countset

	ReDim a_matl(0 To n_mat, 1 To n_par)

'--- define title row

	For i = LBound(a0) To UBound(a0)

		a_matl(0, i + 1) = a0(i)

	Next

'--- fill array with data

	j = 0

	While oMatl.Next

		j = j + 1

		For i = LBound(a0) To UBound(a0)

			s_txt = a0(i)

			If s_txt = "mtrl id" Then a_matl(j, i + 1) = oMatl.ID
			If s_txt = "mtrl name" Then a_matl(j, i + 1) = oMatl.Title
			If s_txt = "type id" Then a_matl(j, i + 1) = oMatl.Type
			If s_txt = "den" Then a_matl(j, i + 1) = oMatl.density * 1000 * 1000 * 1000
			If s_txt = "E11" Then a_matl(j, i + 1) = oMatl.Ex
			If s_txt = "E22" Then a_matl(j, i + 1) = oMatl.Ey
			If s_txt = "E33" Then a_matl(j, i + 1) = oMatl.Ez
			If s_txt = "G12" Then a_matl(j, i + 1) = oMatl.Gx
			If s_txt = "G23" Then a_matl(j, i + 1) = oMatl.Gy
			If s_txt = "G31" Then a_matl(j, i + 1) = oMatl.Gz
			If s_txt = "nu12" Then a_matl(j, i + 1) = oMatl.NUxy
			If s_txt = "nu23" Then a_matl(j, i + 1) = oMatl.NUyz
			If s_txt = "nu13" Then a_matl(j, i + 1) = oMatl.NUxz
			If s_txt = "s11t" Then a_matl(j, i + 1) = oMatl.TensionLimit1
			If s_txt = "s22t" Then a_matl(j, i + 1) = oMatl.TensionLimit2
			If s_txt = "s11c" Then a_matl(j, i + 1) = oMatl.CompressionLimit1
			If s_txt = "s22c" Then a_matl(j, i + 1) = oMatl.CompressionLimit2
			If s_txt = "s12" Then a_matl(j, i + 1) = oMatl.ShearLimit

			If s_txt = "use" Then a_matl(j, i + 1) = 1

		Next

	Wend

'------ end main script

end Sub
