Option Explicit
 ValidationMode = True
 InteractiveMode = im_Abort
 
 
 Dim mdl ' 定义当前的模型
 
 '通过全局参数获得当前的模型
 Set mdl = ActiveModel
 If (mdl Is Nothing) Then
    MsgBox "没有选择模型，请选择一个模型并打开."
 ElseIf Not mdl.IsKindOf(PdPDM.cls_Model) Then
    MsgBox "当前选择的不是一个物理模型（PDM）."
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
    
    
    
oexcel.ActiveWorkbook.SaveAs("d:\立生数据库设计二期_徐敏整理.xlsx")
oexcel.Quit
set nowSheet=Nothing
Set oexcel=Nothing
    
 End If
 
 
 '--------------------------------------------------------------------------------
 '功能函数
 '--------------------------------------------------------------------------------
 Private Sub ProcessFolder(folder,oexcel,num,firstSheet)
    Dim tab '定义数据表对象
    Dim sheetName,rowNum,col
    if InStr(folder.name,"_") > 0  then
		sheetName = Mid(folder.name,4,10)
	else
		sheetName = folder.name
	end if
    	
    for each tab in folder.tables
		oexcel.ActiveWorkbook.Sheets.Add.Name = sheetName & "." & tab.name
    	Set nowSheet=oexcel.ActiveWorkbook.Sheets(sheetName & "." & tab.name)
    	'Sheet1
    	firstSheet.Cells(num, 1) = sheetName
    	firstSheet.Cells(num, 2) = tab.name
    	firstSheet.Cells(num, 3) = tab.code
		num = num + 1
		'新增的sheet
    	'表中文名
		nowSheet.Cells(1, 1) = tab.name
		'表英文名
		nowSheet.Cells(1, 2) = tab.code
		rowNum = 1
		for each col in tab.columns
			rowNum = rowNum + 1
			'字段中文名
			nowSheet.Cells(rowNum, 1) = col.name
			'字段英文名
			nowSheet.Cells(rowNum, 2) = col.code
			nowSheet.Cells(rowNum, 3) = col.comment
			nowSheet.Cells(rowNum, 4) = col.datatype
		next
    next
   
    '对子包进行递归，如果不使用递归只能取到第一个模型图内的表
    
    dim subfolder,sonfolder
    '一级目录
    for each subfolder in folder.Packages
    	if subfolder.tables.count > 0 then
			ProcessFolder subfolder,oexcel,num,firstSheet
		end if
	'二级目录
		for each sonfolder in subfolder.Packages
    		if sonfolder.tables.count > 0 then
				ProcessFolder sonfolder,oexcel,num,firstSheet
			end if
    	next
    next
 
End Sub



