'******************************************************************************
'* File:     fieldCheck.vbs
'* Title:    check every field is end with '_'
'* Purpose:  check every field if it is not end with '_' ,add it then print the field name 
'* Model:    Physical Data Model
'* Objects:  Table, Column, View
'* Author:   qiujingde
'* Created:  2016-11-25
'* Version:  1.0
'******************************************************************************
'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------
' Get the current active model
Dim Model
Set Model = ActiveModel
If (Model Is Nothing) Or (Not Model.IsKindOf(PdPDM.cls_Model)) Then
  MsgBox "The current model is not an PDM model."
Else
 '开始处理
 ScanModel Model
End If

'-----------------------------------------------------------------------------
' Scan tables
'-----------------------------------------------------------------------------
Sub ScanModel(mdl)
   ' For each table
   output "begin"
   
   Dim tab
   For Each tab In mdl.tables
      ScanTable tab
   Next
   
   output "end"
End Sub

'-----------------------------------------------------------------------------

' Show table properties

'-----------------------------------------------------------------------------

Sub ScanTable(tab)
   If IsObject(tab) Then
      ' For each column
	  
      Dim col
      for each col in tab.columns
	    if col.name = "IS_DELETE_" and col.defaultValue <> "0" then
         col.defaultValue = "0"
		   output tab.name + "----------->" + col.name
		 elseif col.name = "IS_FINAL_" and col.defaultValue <> "1" then
         col.defaultValue = "1"
		   output tab.name + "----------->" + col.name
		 elseif col.name = "CRT_TIME_" and col.defaultValue <> "CURRENT_TIMESTAMP" then
         col.defaultValue = "CURRENT_TIMESTAMP"
		   output tab.name + "----------->" + col.name
		 elseif col.name = "UPD_TIME_" and col.defaultValue <> "CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP" then
         col.defaultValue = "CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
		   output tab.name + "----------->" + col.name  
       end if
      next

   End If
End Sub
