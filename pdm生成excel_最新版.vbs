Option Explicit
 ValidationMode = True
 InteractiveMode = im_Abort
 
 
 Dim mdl ' ���嵱ǰ��ģ��
 
 'ͨ��ȫ�ֲ�����õ�ǰ��ģ��
 Set mdl = ActiveModel
 If (mdl Is Nothing) Then
    MsgBox "û��ѡ��ģ�ͣ���ѡ��һ��ģ�Ͳ���."
 ElseIf Not mdl.IsKindOf(PdPDM.cls_Model) Then
    MsgBox "��ǰѡ��Ĳ���һ������ģ�ͣ�PDM��."
 Else
dim num
dim oexcel,nowSheet,firstSheet
num = 1
Set oexcel=CreateObject("excel.application")
oexcel.Workbooks.Add()
'Set worksheet = oexcel.Workbook.Sheets.Add , oexcel.workbook.Sheets(oexcel.workbook.Sheets.Count)
'Sheet1
oExcel.WorkSheets(1).Activate 
Set firstSheet=oexcel.ActiveWorkbook.Sheets("Sheet1")
 
    ProcessFolder mdl,oexcel,num,firstSheet
    
    
    
oexcel.ActiveWorkbook.SaveAs("d:\pgm.xlsx")
oexcel.Quit
set nowSheet=Nothing
Set oexcel=Nothing
    
 End If
 
 
 '--------------------------------------------------------------------------------
 '���ܺ���
 '--------------------------------------------------------------------------------
 Private Sub ProcessFolder(folder,oexcel,num,firstSheet)
    Dim tab '�������ݱ����
	Dim sheetName,rowNum,col
	dim strarr,i
	if InStr(folder.name,"_") > 0  then
		strarr=split(folder.name,"_")
		for i=0 to ubound(strarr)
			sheetName = strarr(i)
		next
	else
		sheetName = folder.name
	end if
    	
	for each tab in folder.tables
		on error resume next
		oexcel.ActiveWorkbook.Sheets.Add.Name = sheetName & "." & tab.name
		if err.number<>0 then  
		'catch ������
		oexcel.ActiveWorkbook.Sheets.Add.Name = sheetName & "." & tab.name & "jyo"
		end if
		'MsgBox sheetName & "." & tab.name
    	Set nowSheet=oexcel.ActiveWorkbook.Sheets(sheetName & "." & tab.name)
    	'Sheet1
    	firstSheet.Cells(num, 1) = sheetName
    	firstSheet.Cells(num, 2) = tab.name
    	firstSheet.Cells(num, 3) = tab.code
		num = num + 1
		'������sheet
    	'��������
		nowSheet.Cells(1, 1) = tab.name
		'��Ӣ����
		nowSheet.Cells(1, 2) = tab.code
		rowNum = 1
		for each col in tab.columns
			rowNum = rowNum + 1
			nowSheet.Cells(rowNum, 1) = col.Name
			nowSheet.Cells(rowNum, 2) = col.Code
			nowSheet.Cells(rowNum, 3) = col.Comment
			nowSheet.Cells(rowNum, 4) = col.DefaultValue
			nowSheet.Cells(rowNum, 5) = col.DataType
			nowSheet.Cells(rowNum, 6) = col.Length
			nowSheet.Cells(rowNum, 7) = col.Precision
			nowSheet.Cells(rowNum, 8) = col.Primary
			nowSheet.Cells(rowNum, 9) = col.ForeignKey
			nowSheet.Cells(rowNum, 10) = col.Mandatory
		next
    next
   
    '���Ӱ����еݹ飬�����ʹ�õݹ�ֻ��ȡ����һ��ģ��ͼ�ڵı�
    
    dim subfolder,sonfolder
    'һ��Ŀ¼
    for each subfolder in folder.Packages
    	if subfolder.tables.count > 0 then
			ProcessFolder subfolder,oexcel,num,firstSheet
		end if
	'����Ŀ¼
		for each sonfolder in subfolder.Packages
    		if sonfolder.tables.count > 0 then
				ProcessFolder sonfolder,oexcel,num,firstSheet
			end if
    	next
    next
 
End Sub



