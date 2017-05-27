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
 sheet.Columns(1).ColumnWidth = 20 
 sheet.Columns(2).ColumnWidth = 20 
 sheet.Columns(3).ColumnWidth = 15 
 sheet.Columns(4).ColumnWidth = 40 
 sheet.Columns(5).ColumnWidth = 30 
 sheet.Columns(1).WrapText =true
 sheet.Columns(2).WrapText =true
 sheet.Columns(4).WrapText =true
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
      
      sheet.cells(rowsNum, 1) = tab.name
      
      sheet.cells(rowsNum, 1).Font.Bold = true
      sheet.cells(rowsNum, 1).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 1).Interior.Color= RGB(205,69,0)
      
      sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 5)).Merge
      rowsNum = rowsNum + 1
      sheet.Range(sheet.cells(rowsNum-1, 1),sheet.cells(rowsNum, 5)).Borders.LineStyle = "1"
      
      sheet.cells(rowsNum, 1) = "表名"
      sheet.cells(rowsNum, 2) = tab.code
      
      sheet.cells(rowsNum, 1).Font.Bold = true
      sheet.cells(rowsNum, 2).Font.Bold = true
      sheet.cells(rowsNum, 2).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 1).HorizontalAlignment = -4108
      
      
      sheet.Range(sheet.cells(rowsNum, 2),sheet.cells(rowsNum, 5)).Merge
      rowsNum = rowsNum + 1
      sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 5)).Borders.LineStyle = "1"
      
      sheet.cells(rowsNum, 1) = "字段中文名"
      sheet.cells(rowsNum, 2) = "字段名"
      sheet.cells(rowsNum, 3) = "字段类型"
      sheet.cells(rowsNum, 4) = "说明"
      sheet.cells(rowsNum, 5) = "备注"
      sheet.cells(rowsNum, 1).Font.Bold = true
      sheet.cells(rowsNum, 2).Font.Bold = true
      sheet.cells(rowsNum, 3).Font.Bold = true
      sheet.cells(rowsNum, 4).Font.Bold = true
      sheet.cells(rowsNum, 5).Font.Bold = true
      sheet.cells(rowsNum, 1).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 2).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 3).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 4).HorizontalAlignment = -4108
      sheet.cells(rowsNum, 5).HorizontalAlignment = -4108
      
      sheet.Range(sheet.cells(rowsNum-1,1),sheet.cells(rowsNum,5)).Interior.Color= RGB(238,154,0)
      
   
      '设置边框
      sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 5)).Borders.LineStyle = "1"

      Dim col ' running column
      Dim colsNum
      colsNum = 0

      for each col in tab.columns
        rowsNum = rowsNum + 1
        colsNum = colsNum + 1
        sheet.cells(rowsNum, 1) = col.name
        sheet.cells(rowsNum, 2) = col.code
        sheet.cells(rowsNum, 3) = col.datatype
        sheet.cells(rowsNum, 4) = col.comment
        sheet.cells(rowsNum, 5) = ""
      next
      
      sheet.Range(sheet.cells(rowsNum-colsNum+1,1),sheet.cells(rowsNum,5)).Borders.LineStyle = "2"
      
      sheet.Range(sheet.cells(rowsNum-colsNum+1,1),sheet.cells(rowsNum,5)).Interior.Color= RGB(108,166,205)
      
      rowsNum = rowsNum + 1
      Output "FullDescription: "       + tab.Name
   End If
End Sub
