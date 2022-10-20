sheet1-source
序号	任务描述	添加时间	完成时间	创建人	级别	备注	文件名	sheet页	[责任人]列号	[VALUE]列号	[TASK]行号	仅统计SE

sheet2-汇总表
任务	添加时间	完成周期	重要性	完成状态	百分比	反馈状态

sheet3-报表
SE	任务数	完成数	PL
张X	35	5	冯X
宋X	40	0	冯X

sheet3-SE名单
SE	业务PL	业务描述
张X	冯X	协同

VBA
Dim inputWb As Workbook
Dim ws As Worksheet
Dim inputWs As Worksheet

Dim destWb As Workbook
Dim destWs As Worksheet



Dim seArrayName  As Variant

Dim seArrayTotal(100) As Integer
Dim seArrayCnt(100) As Integer
Dim seCount As Integer

'以下为表格中提取的输入信息，需要用户确认
Dim inputFileName As String
Dim inputSheetName As String
Dim inputSeCol As String
Dim inputValueCol As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''一些重要的限制条件'''''''''''''''''''''''''''''''''''''

'SE名单表从第二行开始，最多100行


Sub 按钮2_Click()

Dim metError As Integer
metError = False

ret = Application.InputBox("请输入待统计的行号（序号列）", "荣耀任务统计", , , , , , 1)

If ret = False Then
Exit Sub
End If

inputFileName = ThisWorkbook.Sheets("source").Range("H" & ret).value
inputSheetName = ThisWorkbook.Sheets("source").Range("I" & ret).value
inputSeCol = ThisWorkbook.Sheets("source").Range("J" & ret).value
inputValueCol = ThisWorkbook.Sheets("source").Range("K" & ret).value



On Error GoTo Err_Handle_File
Workbooks.Open inputFileName

On Error GoTo Err_Handle_Sheet
Set inputWs = ActiveWorkbook.Sheets(inputSheetName)
On Error GoTo 0 '恢复正常错误处理

''------------------检查SE列是否包含有效的名称
'读取所有合法的SE到内存中
Set ws = ThisWorkbook.Sheets("SE名单")
seArrayName = ws.Range("A2:B50").value
seCount = getArraySize(seArrayName)


For iNew = 1 To 1000 Step 1  '只检查前1000项
    
    If inputWs.Range(inputSeCol & iNew).value = "" Then
        GoTo NextTag001
    End If

    Dim cellValue As String
    cellValue = inputWs.Range(inputSeCol & iNew).value

    matchRow = contains(seArrayName, cellValue) '找到匹配的SE
    If matchRow > 0 Then
        Exit For
    End If
        
NextTag001:
 Next

If iNew = 1001 Then
    inputSeCol = "列号错误，没有包含有效名称，请检查！"
    metError = True
End If

ThisWorkbook.Sheets("source").Activate '重新激活当前页


ContinueMsgbox:
ret = MsgBox("请确认待统计的输入信息：" & vbCrLf & "文件名：" & inputFileName & vbCrLf & "表名：" & inputSheetName & vbCrLf & "SE列号：" & inputSeCol & vbCrLf & "数值列号：" & inputValueCol, vbOKCancel, "请确认信息无误后，再点击确定开始统计")


If ret = vbOK And metError = False Then
    processWorkBook
End If

ThisWorkbook.Sheets("报表").Activate '处理完毕后，默认激活报表页

Exit Sub

Err_Handle_File:
    inputFileName = "文件名不合法，请检查！"
    metError = True
    GoTo ContinueMsgbox

Err_Handle_Sheet:
    inputSheetName = "Sheet页不合法，请检查！"
    metError = True
    GoTo ContinueMsgbox
End Sub




'入参任务编号-行号
'excel文件名
Function processWorkBook()


'之前已经读取所有合法的SE到内存中
'MsgBox (seCount) 'for test

'开始处理源数据表,之前已经打开过了
Workbooks.Open inputFileName
Set inputWs = ActiveWorkbook.Sheets(inputSheetName)



'遍历input表 的所有行，把数据写入内存
'inputSeCol = "J" 'test
'inputValueCol = "G" 'test

For iNew = 2 To 1000 Step 1
    Application.StatusBar = "正在查询页对应数组（" & iNew & "):" & inputWs.Range(inputSeCol & iNew).value
    
    If inputWs.Range(inputSeCol & iNew).value = "" Then
        GoTo NextTag
    End If
    
    Dim cellValue As String
    
    cellValue = inputWs.Range(inputSeCol & iNew).value
    
    Dim matchRow As Integer
    matchRow = contains(seArrayName, cellValue) '找到匹配的SE
    If matchRow > 0 Then
        seArrayTotal(matchRow) = seArrayTotal(matchRow) + 1
        If isWorked(inputWs.Range(inputValueCol & iNew)) Then
            seArrayCnt(matchRow) = seArrayCnt(matchRow) + 1
        End If
    End If
                
       
        
NextTag:
 Next

'把统计数据写入报表页
Dim REPORT_START_ROW As Integer
REPORT_START_ROW = 1

Set destWs = ThisWorkbook.Sheets("报表")
destWs.Range("A2:Z100").Clear

For iNew = 1 To seCount Step 1
    Application.StatusBar = "正在写入报表 ------" & iNew
    
    If seArrayName(iNew, 1) = "" Then
        GoTo NextTag
    End If
    
    seName = seArrayName(iNew, 1) '
    
    destWs.Cells(REPORT_START_ROW + iNew, 1) = seName
    destWs.Cells(REPORT_START_ROW + iNew, 2) = seArrayTotal(iNew)
    destWs.Cells(REPORT_START_ROW + iNew, 3) = seArrayCnt(iNew)
                

 Next
 
 

'报表页后处理，包括填充PL名字
For iNew = 1 To seCount Step 1
    Application.StatusBar = "正在写入报表 ------" & iNew
    
    If seArrayName(iNew, 1) = "" Then
        GoTo NextTag
    End If
    
    plName = seArrayName(iNew, 2) 'PL名字
    destWs.Cells(REPORT_START_ROW + iNew, 4) = plName
 Next
Application.StatusBar = "正在写入报表 ------" & "已经完成后处理，请打开报表页和汇总页查看"
End Function


'TODO:改成名字模糊匹配
Function contains(arr As Variant, value As String) As Integer
contains = -1
For i = 1 To seCount
    If arr(i, 1) = value Then
        contains = i
        Exit For
    End If
Next
End Function

Function isWorked(value As String) As Boolean
If value <> "" And value <> "null" And value <> "NULL" Then
    isWorked = True
Else
    isWorked = False
End If
End Function


'TODO:
'判定参数合法性，弹出提示框，用户确认后下一步
'包括输入的行号，输出的行号
'可选，对应的列号是否包含有效的SE名称
Function prepare()
End Function


Function getArraySize(arr As Variant) As Integer
getArraySize = 0
For i = LBound(arr) To UBound(arr)
    If arr(i, 1) = "" Then
        Exit For
    End If
    getArraySize = getArraySize + 1
Next
End Function
