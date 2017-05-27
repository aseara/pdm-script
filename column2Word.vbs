'******************************************************************************
'* File:     pdm2word.txt
'* Title:    pdm export to word
'* Purpose:  To export the tables and columns to Word
'* Model:    Physical Data Model
'* Objects:  Table, Column, View
'* Author:   qiujingde
'* Created:  2017-03-07
'* Version:  1.0
'******************************************************************************
Option Explicit
  Dim pkCol
  pkCol = ""
'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------
' Get the current active model
Dim model
Set model = ActiveModel
If (model Is Nothing) Or (Not model.IsKindOf(PdPDM.cls_Model)) Then
  MsgBox "The current model is not an PDM model."
Else
  ' Get the tables collection
  '����Word APP
  DIM myWord, myDocument
  Set myWord = CREATEOBJECT("Word.Application")
  Set myDocument = myWord.documents.add  '����ĵ�
 
  ExportModel model, myDocument
 
  myDocument.Activate        '���ĵ����ڻ״̬
  myWord.visible = true      '��ʾ�ĵ�
End If

'-----------------------------------------------------------------------------
' Show properties of tables
'-----------------------------------------------------------------------------
Sub ExportModel(model, document)
  ' Show tables of the current model/package
  ' For each table
  output "begin"
  
  Dim tab
  Dim myRange
  Set myRange = document.Range(0, 0)
   
  For Each tab In model.Tables
    pkCol = ""
    
	With myRange
	  ' �������
	  .InsertAfter(tab.code & Chr(13))
	  .Font.Name = "Times New Roman"
	  .Font.Bold = True
	  .Font.Size = 10
	  .Start = .End
	  
	  ' ���������
      .InsertAfter("�������ƣ�" & tab.name & Chr(13) & "�ṹ��" & Chr(13))
      .Font.Name = "����"
      .Font.Bold = True
      .Font.Size = 10
      .Start = .End
    End With
	
	' ����ֶα��
    ExportColumnsTable tab, document, myRange
	myRange.Start = myRange.End
    
	With myRange
	  ' �������
      .InsertAfter("������")
      .Font.Name = "����"
      .Font.Bold = True
      .Font.Size = 10  
      .Start = .End

      .InsertAfter(pkCol & Chr(13) & Chr(13))
      .Font.Name = "Times New Roman"
      .Font.Bold = False
      .Font.Size = 10
      .Start = .End
	end With
   
   Next
   output "end"
End Sub

'-----------------------------------------------------------------------------
' Show table properties
'-----------------------------------------------------------------------------

Sub ExportColumnsTable(tab, document, range)
  Dim myTable
  Dim rowNum
 
  Set myTable = document.Tables.Add(range, 1, 7, 1)
  rowNum = 1
  
  With myTable
    ' ���ñ����
    .PreferredWidthType = 2
    .PreferredWidth = 100
    ' ȡ������Զ�����
    .AllowAutoFit = False
  
    .Range.Font.Size = 9
    .Range.Font.Bold = False
    .Range.Font.Name = "����"
  
    Dim withArr
	withArr = Array(18, 15, 7, 22, 8, 7, 23)
	
	Dim titleArr
	titleArr = Array("�ֶ���", "�ֶ�����", "����", "���", "�Ƿ�Ϊ��", "ȱʡֵ", "˵��")
	
    Dim i
    For i = 1 to 7
      ' ���ñ�ͷ�о��ж���
      .Cell(1, i).Range.Paragraphs.Alignment = 1
	  .Columns.Item(i).PreferredWidthType = 2
	  .Columns.Item(i).PreferredWidth = withArr(i - 1)
	  .Cell(1, i).Range.Text = titleArr(i - 1)
    next
  
    Dim col
    for each col in tab.columns
      if col.Primary then
        pkCol = col.code
      end if
  
      .Rows.Add
      rowNum = rowNum + 1
      .Cell(rowNum, 1).Range.Text = col.code
      .Cell(rowNum, 2).Range.Text = col.datatype
      if col.length <> 0 then
        .Cell(rowNum, 3).Range.Text = col.length
      end if  
      if col.foreignKey then
        .Cell(rowNum, 4).Range.Text = "��"
      end if
      if col.Mandatory then
        .Cell(rowNum, 5).Range.Text = "��"
      else 
        .Cell(rowNum, 5).Range.Text = "��"
      end if
      if col.code <> "CRT_TIME_" and col.code <> "UPD_TIME_" then
        .Cell(rowNum, 6).Range.Text = col.defaultValue
      end if
      .Cell(rowNum, 7).Range.Text = col.comment
	
	  For i = 1 to 7
	    ' ���������п������
        .Cell(rowNum, i).Range.Paragraphs.Alignment = 0
      next
    next
  Set range = .Range
  End With
End Sub
