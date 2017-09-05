'****************************************************************************** 
'* Title:    PDM2ExcelVB 
'* Purpose:  Export Physical Data Model table list and table columns to Excel 
'* Author:   Lucas 
'* Created:  2017-09-05 
'* Version:  0.1 
'****************************************************************************** 
Option Explicit 
    Dim rowsNum 
    rowsNum = 0 
    Dim Fldr
    Set Fldr = ActiveDiagram.Parent
    Dim vExcel, vBook, vSheet
    Set vExcel = CreateObject("Excel.Application")
    vExcel.Visible = True
    Set vBook = vExcel.Workbooks.Add(-4167)
    vBook.Sheets(1).Name = "���ݿ��ṹ"
    Set vSheet = vExcel.workbooks(1).sheets("���ݿ��ṹ")
    ShowProperties Fldr, vSheet
    vExcel.visible = true 
    vSheet.Columns(1).ColumnWidth = 10   
    vSheet.Columns(2).ColumnWidth = 40   
    vSheet.Columns(3).ColumnWidth = 50   
    vSheet.Columns(1).WrapText =true 
    vSheet.Columns(2).WrapText =true 
    vSheet.Columns(3).WrapText =true 
    vSheet.Activate
    vSheet.Cells.EntireColumn.AutoFit
    vSheet.Cells.EntireRow.AutoFit
Sub ShowProperties(fldrs, sheets) 
    output "begin" 
    ListObjects fldrs,sheets
    output "end" 
End Sub

Sub ListObjects(fldrn,sheetn)
    Dim obj ' running object
    For Each obj In fldrn.children
        if obj.ClassName ="Table" then 
            ShowTable obj,sheetn
        end if
    Next
    Dim f ' running folder
    For Each f In fldrn.Packages
        ListObjects f,sheetn
    Next
End Sub

Sub ShowTable(tab, sheet)   
    If IsObject(tab) Then 
        Dim rangFlag
        sheet.cells(1, 1) = "���" 
        sheet.cells(1, 2) = "����"
        sheet.cells(1, 3) = "ʵ����"
        '���ñ߿� 
        sheet.Range(sheet.cells(1, 1),sheet.cells(1, 3)).Borders.LineStyle = "1"
        '���ñ�����ɫ
        sheet.Range(sheet.cells(1, 1),sheet.cells(1, 3)).Interior.ColorIndex = "19"

        rowsNum = rowsNum + 1
        sheet.cells(rowsNum+1, 1) = rowsNum 
        sheet.cells(rowsNum+1, 2) = tab.code
        sheet.cells(rowsNum+1, 3) = tab.name
        sheet.Hyperlinks.Add sheet.cells(rowsNum+1, 2), "", (tab.code+"!A1"), tab.code
        '���ñ߿�
        sheet.Range(sheet.cells(rowsNum+1,1),sheet.cells(rowsNum+1,3)).Borders.LineStyle = "2"
        '����Sheet
        vBook.Sheets.Add , vBook.Sheets(vBook.Sheets.count)
        vBook.Sheets(rowsNum+1).Name = tab.code 

        Dim shtn
        Set shtn = vExcel.workbooks(1).sheets(tab.code)
        shtn.Cells(1, 4).FormulaR1C1 = "�����ܱ�"
        shtn.Hyperlinks.Add shtn.Cells(1, 4), "", "���ݿ��ṹ!A1", "�����ܱ�"

        '�����п�ͻ���
        shtn.Columns(1).ColumnWidth = 30   
        shtn.Columns(2).ColumnWidth = 20   
        shtn.Columns(3).ColumnWidth = 20
        shtn.Columns(5).ColumnWidth = 30   
        shtn.Columns(6).ColumnWidth = 20   
        shtn.Columns(1).WrapText =true 
        shtn.Columns(2).WrapText =true 
        shtn.Columns(3).WrapText =true
        shtn.Columns(5).WrapText =true 
        shtn.Columns(6).WrapText =true

        '�����б���
        shtn.cells(1, 1) = "�ֶ�������" 
        shtn.cells(1, 2) = "�ֶ���"
        shtn.cells(1, 3) = "�ֶ�����"
        shtn.cells(1, 5) = tab.code
        shtn.cells(1, 6) = tab.Name
        '���ñ߿� 
        shtn.Range(shtn.cells(1, 1),shtn.cells(1, 3)).Borders.LineStyle = "1"
        shtn.Range(shtn.cells(1, 4),shtn.cells(1, 4)).Borders.LineStyle = "1"
        shtn.Range(shtn.cells(1, 5),shtn.cells(1, 6)).Borders.LineStyle = "1"
        '���ñ�����ɫ
        shtn.Range(shtn.cells(1, 1),shtn.cells(1, 3)).Interior.ColorIndex = "19"
        shtn.Range(shtn.cells(1, 4),shtn.cells(1, 4)).Interior.ColorIndex = "8"
        shtn.Range(shtn.cells(1, 5),shtn.cells(1, 6)).Interior.ColorIndex = "19"

        Dim col ' running column 
        Dim colsNum
        Dim rNum 
        colsNum = 0
        rNum = 0 
        for each col in tab.columns 
            rNum = rNum + 1 
            colsNum = colsNum + 1 
            shtn.cells(rNum+1, 1) = col.name 
            shtn.cells(rNum+1, 2) = col.code 
            shtn.cells(rNum+1, 3) = col.datatype 
        next 
        shtn.Range(shtn.cells(rNum-colsNum+2,1),shtn.cells(rNum+1,3)).Borders.LineStyle = "2"         
        rNum = rNum + 1 
        Output "FullDescription: " + tab.Name
    End If   
End Sub
