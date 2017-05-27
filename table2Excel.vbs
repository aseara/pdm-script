'******************************************************************************
'* File:     pdm2excel.txt
'* Title:    pdm export to excel
'* Purpose:  To export the tables and columns to Excel
'* Model:    Physical Data Model
'* Objects:  Table, Column, View
'* Author:   ziyan
'* Created:  2012-05-03
'* Version:  1.0
'******************************************************************************
Option Explicit
   Dim rowsNum
   rowsNum = 0
'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------
' Get the current active model
Dim Model
Set Model = ActiveModel
If (Model Is Nothing) Or (Not Model.IsKindOf(PdPDM.cls_Model)) Then
  MsgBox "The current model is not an PDM model."
Else
 ' Get the tables collection
 '创建EXCEL APP
 dim beginrow
 DIM EXCEL, SHEET
 set EXCEL = CREATEOBJECT("Excel.Application")
 EXCEL.workbooks.add(-4167)'添加工作表
 EXCEL.workbooks(1).sheets(1).name ="test"
 set sheet = EXCEL.workbooks(1).sheets("test")

 ShowProperties Model, SHEET
 EXCEL.visible = true
 '设置列宽和自动换行
 sheet.Columns(1).ColumnWidth = 24 
 sheet.Columns(2).ColumnWidth = 40 
 sheet.Columns(1).WrapText =true
 sheet.Columns(2).WrapText =true
End If

'-----------------------------------------------------------------------------
' Show properties of tables
'-----------------------------------------------------------------------------
Sub ShowProperties(mdl, sheet)
   ' Show tables of the current model/package
   rowsNum=0
   beginrow = rowsNum+1
   ' For each table
   output "begin"
   
   Dim tab
   For Each tab In mdl.tables
      ShowTable tab,sheet
   Next
   
   if mdl.tables.count > 0 then
        sheet.Range("A" & beginrow + 1 & ":A" & rowsNum).Rows.Group
   end if
   
   output "end"
End Sub

'-----------------------------------------------------------------------------

' Show table properties

'-----------------------------------------------------------------------------

Sub ShowTable(tab, sheet)
   If IsObject(tab) Then
     Dim rangFlag
     rowsNum = rowsNum + 1
      ' Show properties

      Output "================================"
      
      sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 2)).Borders.LineStyle = "1"
      
      sheet.cells(rowsNum, 1) = tab.code
      sheet.cells(rowsNum, 2) = tab.name
      
      sheet.cells(rowsNum, 1).HorizontalAlignment = -4108
      
      Output "FullDescription: "       + tab.Name
   End If
End Sub
