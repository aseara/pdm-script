'******************************************************************************
'* File:     defaultNull.vbs
'* Title:    default value check
'* Purpose:  Delete column default value if it is null
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
   
   Dim first
   Dim tab
   For Each tab In mdl.tables
      if first = 0 then
        first = 1
        output tab.name
        ScanTable tab
      end if
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
      Dim colCheck
      for each col in tab.columns
        if colCheck = 0 then
          colCheck = 1
	       output col.name
          
        end if
      next

   End If
End Sub
