'TODO:   按照以下步骤启用功能区(XML)项: 

'1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
'   操作(例如单击按钮)。注意: 如果已经从功能区设计器中导出此功能区，
'   请将代码从事件处理程序移动到回调方法，并
'   修改该代码以使用功能区扩展性(RibbonX)编程模型。

'3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。

'有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI


    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("FunWithExcel.Ribbon1.xml")
    End Function

#Region "功能区回调"

    '在此创建回调方法。有关添加回调方法的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub onAutoGen(ByVal control As Office.IRibbonControl)
        Dim ash As Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)
        If LTrim(RTrim(ash.Range("A1").Value)) <> "名称" And LTrim(RTrim(ash.Range("B1").Value)) <> "偏差" Then
            MsgBox("不适应于该Sheet工作表!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "错误")
            Exit Sub
        End If
        Erase ArrayStr
        Erase tmpNumber
        Erase tmpStrText
        Erase tmpT
        Erase tmpUW
        Dim countT As Integer = 0
        Dim countUW As Integer = 0
        Do Until ash.Range("T" & countT + 2).Text = ""
            countT = countT + 1
        Loop
        Do Until ash.Range("U" & countUW + 2).Text = ""
            countUW = countUW + 1
        Loop
        If countT > 0 And countUW > 0 Then
            ReDim tmpT(countT - 1)
            ReDim tmpUW(countUW - 1, 2)
            countT = 0
            Do Until ash.Range("T" & countT + 2).Text = ""
                tmpT(countT) = ash.Range("T" & countT + 2).Value
                countT = countT + 1
            Loop
            countUW = 0
            Do Until ash.Range("U" & countUW + 2).Text = ""
                tmpUW(countUW, 0) = ash.Range("U" & countUW + 2).Value
                tmpUW(countUW, 1) = ash.Range("V" & countUW + 2).Value
                tmpUW(countUW, 2) = ash.Range("W" & countUW + 2).Value
                countUW = countUW + 1
            Loop
        End If
        '////////////////////////////////////////////////////////////////////////////////
        Dim a As New Form1
        a.ShowDialog()
        If UBound(tmpNumber) < 0 Or UBound(ArrayStr) < 0 Then
            MsgBox("数组为空！", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "错误")
            Exit Sub
        End If
        ash.Range("T:AI").Delete()
        With ash.Range("AA:AI")
            .NumberFormatLocal = "#0.0"
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = -4108
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With
        Dim tmpmax As Integer = 0
        Dim i As Integer
        For i = 0 To UBound(ArrayStr)
            If ArrayStr(i, 2) > tmpmax Then
                tmpmax = ArrayStr(i, 2)
            End If
        Next
        With ash.Range("T:W").Font
            .Name = "宋体"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .TintAndShade = 0
        End With
        ash.Range("T1").Value = "横轴线编号"
        ash.Range("U1").Value = "起始轴线"
        ash.Range("V1").Value = "终止轴线"
        ash.Range("W1").Value = "该区间轴线排数"
        For i = 0 To UBound(tmpNumber)
            ash.Range("T" & i + 2).Value = tmpNumber(i)
        Next
        For i = 0 To UBound(ArrayStr)
            ash.Range("U" & i + 2).Value = ArrayStr(i, 0)
            ash.Range("V" & i + 2).Value = ArrayStr(i, 1)
            ash.Range("W" & i + 2).Value = ArrayStr(i, 2)
        Next
        If tmpmax > 2 Then
            ash.Range("AA1").Value = "柱号"
            ash.Range("AB1").Value = "北偏/cm"
            ash.Range("AC1").Value = "东偏/cm"
            ash.Range("AD1").Value = "柱号"
            ash.Range("AE1").Value = "北偏/cm"
            ash.Range("AF1").Value = "东偏/cm"
            ash.Range("AG1").Value = "柱号"
            ash.Range("AH1").Value = "北偏/cm"
            ash.Range("AI1").Value = "东偏/cm"
        ElseIf tmpmax > 1 Then
            ash.Range("AA1").Value = "柱号"
            ash.Range("AB1").Value = "北偏/cm"
            ash.Range("AC1").Value = "东偏/cm"
            ash.Range("AD1").Value = "柱号"
            ash.Range("AE1").Value = "北偏/cm"
            ash.Range("AF1").Value = "东偏/cm"
        Else
            ash.Range("AA1").Value = "柱号"
            ash.Range("AB1").Value = "北偏/cm"
            ash.Range("AC1").Value = "东偏/cm"
        End If
        Dim Num As Integer = 1
        Dim j As Integer
        Dim tmpCN, tmpCE, tmpBN, tmpBE, tmpDN, tmpDE As Double
        For i = 0 To UBound(ArrayStr)
            If ArrayStr(i, 2) = 3 Then
                For j = tmpStrText(i, 0) To tmpStrText(i, 1) Step 6
                    ash.Range("AA" & Num + 1).Value = tmpNumber(0) & Num
                    ash.Range("AD" & Num + 1).Value = tmpNumber(1) & Num
                    ash.Range("AG" & Num + 1).Value = tmpNumber(2) & Num
                    tmpCN = ash.Range("K" & j).Value * 100 * 2 / 3
                    tmpCE = ash.Range("J" & j + 1).Value * 100 * 2 / 3
                    tmpBN = ash.Range("K" & j + 2).Value * 100 * 2 / 3
                    tmpBE = ash.Range("J" & j + 3).Value * 100 * 2 / 3
                    tmpDN = ash.Range("K" & j + 4).Value * 100 * 2 / 3
                    tmpDE = ash.Range("J" & j + 5).Value * 100 * 2 / 3
                    ash.Range("AB" & Num + 1).Value = Format(Math.Round(tmpCN, 2), "0.0")
                    ash.Range("AC" & Num + 1).Value = Format(Math.Round(tmpCE, 2), "0.0")
                    ash.Range("AE" & Num + 1).Value = Format(Math.Round(tmpBN, 2), "0.0")
                    ash.Range("AF" & Num + 1).Value = Format(Math.Round(tmpBE, 2), "0.0")
                    ash.Range("AH" & Num + 1).Value = Format(Math.Round(tmpDN, 2), "0.0")
                    ash.Range("AI" & Num + 1).Value = Format(Math.Round(tmpDE, 2), "0.0")
                    If Math.Abs(tmpCN) > 2 Then
                        With ash.Range("AB" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpCE) > 2 Then
                        With ash.Range("AC" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpBN) > 2 Then
                        With ash.Range("AE" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpBE) > 2 Then
                        With ash.Range("AF" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpDN) > 2 Then
                        With ash.Range("AH" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpDE) > 2 Then
                        With ash.Range("AI" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    Num = Num + 1
                Next
            ElseIf ArrayStr(i, 2) = 2 Then
                For j = tmpStrText(i, 0) To tmpStrText(i, 1) Step 4
                    ash.Range("AA" & Num + 1).Value = tmpNumber(0) & Num
                    ash.Range("AD" & Num + 1).Value = tmpNumber(1) & Num
                    tmpCN = ash.Range("K" & j).Value * 100 * 2 / 3
                    tmpCE = ash.Range("J" & j + 1).Value * 100 * 2 / 3
                    tmpBN = ash.Range("K" & j + 2).Value * 100 * 2 / 3
                    tmpBE = ash.Range("J" & j + 3).Value * 100 * 2 / 3
                    ash.Range("AB" & Num + 1).Value = Format(Math.Round(tmpCN, 2), "0.0")
                    ash.Range("AC" & Num + 1).Value = Format(Math.Round(tmpCE, 2), "0.0")
                    ash.Range("AE" & Num + 1).Value = Format(Math.Round(tmpBN, 2), "0.0")
                    ash.Range("AF" & Num + 1).Value = Format(Math.Round(tmpBE, 2), "0.0")
                    If Math.Abs(tmpCN) > 2 Then
                        With ash.Range("AB" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpCE) > 2 Then
                        With ash.Range("AC" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpBN) > 2 Then
                        With ash.Range("AE" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpBE) > 2 Then
                        With ash.Range("AF" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    Num = Num + 1
                Next
            ElseIf ArrayStr(i, 2) = 1 Then
                For j = tmpStrText(i, 0) To tmpStrText(i, 1) Step 2
                    ash.Range("AA" & Num + 1).Value = tmpNumber(0) & Num
                    tmpCN = ash.Range("K" & j).Value * 100 * 2 / 3
                    tmpCE = ash.Range("J" & j + 1).Value * 100 * 2 / 3
                    ash.Range("AB" & Num + 1).Value = Format(Math.Round(tmpCN, 2), "#0.0")
                    ash.Range("AC" & Num + 1).Value = Format(Math.Round(tmpCE, 2), "#0.0")
                    If Math.Abs(tmpCN) > 2 Then
                        With ash.Range("AB" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    If Math.Abs(tmpCE) > 2 Then
                        With ash.Range("AC" & Num + 1).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If
                    Num = Num + 1
                Next
            End If
        Next
        MsgBox("完成", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
    End Sub
    Public Sub onAbout1(ByVal control As Office.IRibbonControl)
        MsgBox("Excel插件 by VSTO VB.Net", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "提示")
    End Sub
    Public Sub onCalPrism(ByVal control As Office.IRibbonControl)
        Dim ash As Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)
        If LTrim(RTrim(ash.Range("A1").Value)) <> "周期" And LTrim(RTrim(ash.Range("B1").Value)) <> "变形点" Then
            MsgBox("不适应于该Sheet工作表!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "错误")
            Exit Sub
        End If
        Dim Row As Integer = 2
        Do Until ash.Range("A" & Row).Value < 1
            If LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1A" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1B" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1C" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJA" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJC" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJD" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJE" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJF" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1A" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1B" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1C" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJA" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJC" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJD" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJE" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJF" Then
                ash.Range("G" & Row).Value = ash.Range("G" & Row).Value - 0.03
            ElseIf LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXZA" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXZB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXZC" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXZA" _
                Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXZB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXZC" Then
                ash.Range("G" & Row).Value = ash.Range("G" & Row).Value - 0.007
            Else
                ash.Range("G" & Row).Value = ash.Range("G" & Row).Value - 0.025
            End If
            Row = Row + 1
        Loop
        ash.Range("A1").Select()
        'ash.Range("A" & "1").Value = "12"
        MsgBox("完成", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
    End Sub
    Const PI = 3.1415926535898
    Public Sub onCalBase(ByVal control As Office.IRibbonControl)
        Dim ash As Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)
        If LTrim(RTrim(ash.Range("A1").Value)) <> "周期" And LTrim(RTrim(ash.Range("B1").Value)) <> "变形点" Then
            MsgBox("不适应于该Sheet工作表!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "错误")
            Exit Sub
        End If
        Dim Row As Integer = 2
        Do Until ash.Range("B" & Row).Text = ""
            If LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1A" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1B" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJ1C" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJA" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJC" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJD" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJE" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "DXJF" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1A" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1B" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJ1C" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJA" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJB" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJC" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJD" _
            Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJE" Or LTrim(RTrim(ash.Range("B" & Row).Text)) = "XXJF" Then
                Dim Hz, V As String
                Dim hdd, hmm, hss, vdd, vmm, vss, sdist, X, Y, Z As Double
                Hz = ash.Range("E" & Row).Text
                V = ash.Range("F" & Row).Text
                hdd = Mid(Hz, 1, InStr(1, Hz, "°") - 1)
                hmm = Mid(Hz, InStr(1, Hz, "°") + 1, InStr(1, Hz, "′") - InStr(1, Hz, "°") - 1)
                hss = Mid(Hz, InStr(1, Hz, "′") + 1, Len(Hz) - InStr(1, Hz, "′") - 1)
                vdd = Mid(V, 1, InStr(1, V, "°") - 1)
                vmm = Mid(V, InStr(1, V, "°") + 1, InStr(1, V, "′") - InStr(1, V, "°") - 1)
                vss = Mid(V, InStr(1, V, "′") + 1, Len(V) - InStr(1, V, "′") - 1)
                sdist = ash.Range("G" & Row).Value
                X = sdist * Math.Cos(PI / 2 - (vdd + vmm / 60 + vss / 3600) * PI / 180) * Math.Cos((hdd + hmm / 60 + hss / 3600) * PI / 180)
                Y = sdist * Math.Cos(PI / 2 - (vdd + vmm / 60 + vss / 3600) * PI / 180) * Math.Sin((hdd + hmm / 60 + hss / 3600) * PI / 180)
                Z = sdist * Math.Sin(PI / 2 - (vdd + vmm / 60 + vss / 3600) * PI / 180)
                ash.Range("I" & Row).Value = X
                ash.Range("J" & Row).Value = Y
                ash.Range("K" & Row).Value = Z
            End If
            Row = Row + 1
        Loop
        MsgBox("完成", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
    End Sub
    Public Sub onReplace(ByVal control As Office.IRibbonControl)
        'Dim ash1 As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("原始数据"), Excel.Worksheet)
        'If CType(Globals.ThisAddIn.Application.Worksheets("自动化监测"), Excel.Worksheet) Is Nothing Or
        '        CType(Globals.ThisAddIn.Application.Worksheets("隧道收敛"), Excel.Worksheet) Is Nothing Then
        '    MsgBox("指定Sheet工作表不存在", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
        'End If
        Dim ash2 As Excel.Worksheet
        Dim ash3 As Excel.Worksheet
        Try
            ash2 = CType(Globals.ThisAddIn.Application.Worksheets("自动化监测"), Excel.Worksheet)
            ash3 = CType(Globals.ThisAddIn.Application.Worksheets("隧道收敛"), Excel.Worksheet)
        Catch ex As Exception
            MsgBox("指定Sheet工作表不存在", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
            Exit Sub
        End Try

        Dim nRows2, nRows3, tablerow As Integer
        nRows2 = ash2.UsedRange.Rows.Count
        nRows3 = ash3.UsedRange.Rows.Count
        Dim tabletext As String = ""
        For tablerow = 1 To nRows2
            tabletext = ash2.Range("A" & tablerow).Text
            If tabletext = "东线隧道沉降（自动）监测报表" Or tabletext = "西线隧道沉降（自动）监测报表" Then
                Dim initrow1, initrow11 As Integer
                initrow1 = tablerow + 4   '点号起始行
                initrow11 = tablerow + 4   '点号起始行
                Do Until ash2.Range("A" & initrow1).Text = ""
                    Dim str1 As Double
                    str1 = ash2.Range("B" & initrow1).Value + ash2.Range("C" & initrow1).Value
                    ash2.Range("B" & initrow1).Value = Format(str1, "0.0")
                    initrow1 = initrow1 + 1
                Loop
                Do Until ash2.Range("F" & initrow11).Text = ""
                    Dim str2 As Double
                    str2 = ash2.Range("G" & initrow11).Value + ash2.Range("H" & initrow11).Value
                    ash2.Range("G" & initrow11).Value = Format(str2, "0.0")
                    initrow11 = initrow11 + 1
                Loop
            ElseIf tabletext = "东线隧道水平位移监测报表" Or tabletext = "西线隧道水平位移监测报表" Then
                Dim initrow1, initrow11 As Integer
                initrow1 = tablerow + 4   '点号起始行
                initrow11 = tablerow + 4   '点号起始行
                Do Until ash2.Range("A" & initrow1).Text = ""
                    Dim str1 As Double
                    str1 = ash2.Range("B" & initrow1).Value + ash2.Range("C" & initrow1).Value
                    ash2.Range("B" & initrow1).Value = Format(str1, "0.0")
                    initrow1 = initrow1 + 1
                Loop
                Do Until ash2.Range("F" & initrow11).Text = ""
                    Dim str2 As Double
                    str2 = ash2.Range("G" & initrow11).Value + ash2.Range("H" & initrow11).Value
                    ash2.Range("G" & initrow11).Value = Format(str2, "0.0")
                    initrow11 = initrow11 + 1
                Loop
            ElseIf tabletext = "东线隧道倾斜监测报表" Or tabletext = "西线隧道倾斜监测报表" Then

            End If
            tabletext = ""
        Next   'sheet自动化监测替换完成
        '//////////////////////////////////////////////////////////////////////////////////
        For tablerow = 1 To nRows3
            tabletext = ash3.Range("A" & tablerow).Text
            If tabletext = "东线隧道收敛监测报表" Or tabletext = "西线隧道收敛监测报表" Then
                Dim initrow As Integer
                initrow = tablerow + 3
                Dim tmptext As String
                tmptext = ash3.Range("A" & initrow).Text
                Do Until tmptext = ""
                    ash3.Range("D" & initrow).Value = ash3.Range("E" & initrow).Value
                    ash3.Range("D" & initrow + 1).Value = ash3.Range("E" & initrow + 1).Value
                    ash3.Range("D" & initrow + 2).Value = ash3.Range("E" & initrow + 2).Value
                    ash3.Range("D" & initrow + 3).Value = ash3.Range("E" & initrow + 3).Value
                    initrow = initrow + 4
                    tmptext = ash3.Range("A" & initrow).Text
                Loop
            End If
            tabletext = ""
        Next
        MsgBox("替换完成", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "结果")

    End Sub
    Public Sub onGenerate(ByVal control As Office.IRibbonControl)
        'If CType(Globals.ThisAddIn.Application.Worksheets("原始数据"), Excel.Worksheet) Is Nothing Or
        '    CType(Globals.ThisAddIn.Application.Worksheets("自动化监测"), Excel.Worksheet) Is Nothing Or
        '        CType(Globals.ThisAddIn.Application.Worksheets("隧道收敛"), Excel.Worksheet) Is Nothing Then
        '    MsgBox("指定Sheet工作表不存在", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
        'End If
        Dim ash1 As Excel.Worksheet
        Dim ash2 As Excel.Worksheet
        Dim ash3 As Excel.Worksheet
        Try
            ash1 = CType(Globals.ThisAddIn.Application.Worksheets("原始数据"), Excel.Worksheet)
            ash2 = CType(Globals.ThisAddIn.Application.Worksheets("自动化监测"), Excel.Worksheet)
            ash3 = CType(Globals.ThisAddIn.Application.Worksheets("隧道收敛"), Excel.Worksheet)
        Catch ex As Exception
            MsgBox("指定Sheet工作表不存在", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "提示")
            Exit Sub
        End Try
        'Dim ash1 As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("原始数据"), Excel.Worksheet)
        'Dim ash2 As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("自动化监测"), Excel.Worksheet)
        'Dim ash3 As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("隧道收敛"), Excel.Worksheet)
        Dim nRows1, nRows2, nRows3, tablerow As Integer
        nRows1 = ash1.UsedRange.Rows.Count
        nRows2 = ash2.UsedRange.Rows.Count
        nRows3 = ash3.UsedRange.Rows.Count
        Dim tabletext As String = ""
        For tablerow = 1 To nRows2
            tabletext = ash2.Range("A" & tablerow).Text
            If tabletext = "东线隧道沉降（自动）监测报表" Or tabletext = "西线隧道沉降（自动）监测报表" Then
                Dim initrow1, initrow11 As Integer
                initrow1 = tablerow + 4   '点号起始行
                initrow11 = tablerow + 4   '点号起始行
                Do Until ash2.Range("A" & initrow1).Text = ""
                    Dim str1 As String
                    str1 = ash2.Range("A" & initrow1).Text
                    Dim tmp1 As Integer
                    Dim tmpz1, tmpz2 As Double
                    For tmp1 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp1).Text)) = LTrim(RTrim(str1)) Then
                            tmpz1 = ash1.Range("D" & tmp1).Value
                            tmpz2 = ash1.Range("K" & tmp1).Value
                            Exit For
                        End If
                    Next
                    ash2.Range("C" & initrow1).Value = Format((tmpz1 - tmpz2) * 1000, "0.0")
                    initrow1 = initrow1 + 1
                Loop
                Do Until ash2.Range("F" & initrow11).Text = ""
                    Dim str2 As String
                    str2 = ash2.Range("F" & initrow11).Text
                    If LTrim(RTrim(str2)) = "DX19B" Then
                        str2 = "DXZA"
                    End If
                    If LTrim(RTrim(str2)) = "DX21B" Then
                        str2 = "DXZB"
                    End If
                    Dim tmp1 As Integer
                    Dim tmpz1, tmpz2 As Double
                    For tmp1 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp1).Text)) = LTrim(RTrim(str2)) Then
                            tmpz1 = ash1.Range("D" & tmp1).Value
                            tmpz2 = ash1.Range("K" & tmp1).Value
                            Exit For
                        End If
                    Next
                    ash2.Range("H" & initrow11).Value = Format((tmpz1 - tmpz2) * 1000, "0.0")
                    initrow11 = initrow11 + 1
                Loop
            ElseIf tabletext = "东线隧道水平位移监测报表" Or tabletext = "西线隧道水平位移监测报表" Then
                Dim initrow1, initrow11 As Integer
                initrow1 = tablerow + 4   '点号起始行
                initrow11 = tablerow + 4   '点号起始行
                Do Until ash2.Range("A" & initrow1).Text = ""
                    Dim str1 As String
                    str1 = ash2.Range("A" & initrow1).Text
                    Dim tmp1 As Integer
                    Dim tmpz1, tmpz2 As Double
                    For tmp1 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp1).Text)) = LTrim(RTrim(str1)) Then
                            tmpz1 = ash1.Range("C" & tmp1).Value
                            tmpz2 = ash1.Range("J" & tmp1).Value
                            Exit For
                        End If

                    Next
                    ash2.Range("C" & initrow1).Value = Format((tmpz1 - tmpz2) * 1000, "0.0")
                    initrow1 = initrow1 + 1
                Loop
                Do Until ash2.Range("F" & initrow11).Text = ""
                    Dim str2 As String
                    str2 = ash2.Range("F" & initrow11).Text
                    If LTrim(RTrim(str2)) = "DX19B" Then
                        str2 = "DXZA"
                    End If
                    If LTrim(RTrim(str2)) = "DX21B" Then
                        str2 = "DXZB"
                    End If
                    Dim tmp1 As Integer
                    Dim tmpz1, tmpz2 As Double
                    For tmp1 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp1).Text)) = LTrim(RTrim(str2)) Then
                            tmpz1 = ash1.Range("C" & tmp1).Value
                            tmpz2 = ash1.Range("J" & tmp1).Value
                            Exit For
                        End If
                    Next
                    ash2.Range("H" & initrow11).Value = Format((tmpz1 - tmpz2) * 1000, "0.0")
                    initrow11 = initrow11 + 1
                Loop
            ElseIf tabletext = "东线隧道倾斜监测报表" Or tabletext = "西线隧道倾斜监测报表" Then
                Dim initrow2 = tablerow + 4
                Do Until ash2.Range("A" & initrow2).Text = ""
                    Dim sloperate As Double
                    Dim tmpstr, str1, str2 As String
                    tmpstr = ash2.Range("A" & initrow2).Text
                    str1 = RTrim(tmpstr) & "B"
                    str2 = RTrim(tmpstr) & "D"
                    If LTrim(str2) = "DX20D" Then
                        str2 = "DXZC"
                    End If
                    Dim tmp2, tmp3 As Integer
                    Dim BX1, BY1, BZ1, BX2, BY2, BZ2, DX1, DY1, DZ1, DX2, DY2, DZ2 As Double
                    For tmp2 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp2).Text)) = LTrim(str1) Then
                            BX1 = ash1.Range("B" & tmp2).Value
                            BY1 = ash1.Range("C" & tmp2).Value
                            BZ1 = ash1.Range("D" & tmp2).Value
                            BX2 = ash1.Range("I" & tmp2).Value
                            BY2 = ash1.Range("J" & tmp2).Value
                            BZ2 = ash1.Range("K" & tmp2).Value
                            Exit For
                        End If
                    Next
                    For tmp3 = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & tmp3).Text)) = LTrim(str2) Then
                            DX1 = ash1.Range("B" & tmp3).Value
                            DY1 = ash1.Range("C" & tmp3).Value
                            DZ1 = ash1.Range("D" & tmp3).Value
                            DX2 = ash1.Range("I" & tmp3).Value
                            DY2 = ash1.Range("J" & tmp3).Value
                            DZ2 = ash1.Range("K" & tmp3).Value
                            Exit For
                        End If
                    Next
                    sloperate = ((BZ1 - BZ2) - (DZ1 - DZ2)) /
                    ((Math.Sqrt((BX1 - DX1) * (BX1 - DX1) + (BY1 - DY1) * (BY1 - DY1) + (BZ1 - DZ1) * (BZ1 - DZ1)) +
                    Math.Sqrt((BX2 - DX2) * (BX2 - DX2) + (BY2 - DY2) * (BY2 - DY2) + (BZ2 - DZ2) * (BZ2 - DZ2))) / 2)
                    ash2.Range("D" & initrow2).Value = Format(sloperate, "0.00") & "%"
                    If tabletext = "东线隧道倾斜监测报表" Then
                        If sloperate > 0 Then
                            ash2.Range("F" & initrow2).Value = "向A10基坑方向倾斜"
                        ElseIf sloperate < 0 Then
                            ash2.Range("F" & initrow2).Value = "向A03基坑方向倾斜"
                        Else
                            ash2.Range("F" & initrow2).Value = "未倾斜"
                        End If
                    ElseIf tabletext = "西线隧道倾斜监测报表" Then
                        If sloperate > 0 Then
                            ash2.Range("F" & initrow2).Value = "向A03基坑方向倾斜"
                        ElseIf sloperate < 0 Then
                            ash2.Range("F" & initrow2).Value = "向A10基坑方向倾斜"
                        Else
                            ash2.Range("F" & initrow2).Value = "未倾斜"
                        End If
                    End If
                    initrow2 = initrow2 + 1
                Loop
            End If
            tabletext = ""
        Next  'sheet自动化监测计算完成
        '///////////////////////////////////////////////////////////////////////////////////
        For tablerow = 1 To nRows3
            tabletext = ash3.Range("A" & tablerow).Text
            If tabletext = "东线隧道收敛监测报表" Or tabletext = "西线隧道收敛监测报表" Then
                Dim initrow As Integer
                Dim tmptext As String
                initrow = tablerow + 3
                tmptext = ash3.Range("A" & initrow).Text
                Do Until tmptext = ""
                    Dim strA, strB, strC, strD, strE As String
                    If LTrim(RTrim(tmptext)) = "XX19" Then
                        strA = RTrim(tmptext) & "A"
                        strB = "XXZA"
                        strC = "XXZB"
                        strD = "XXZC"
                        strE = "XX19E"
                    ElseIf LTrim(RTrim(tmptext)) = "DX20" Then
                        strA = "DX20A"
                        strB = "DX20B"
                        strC = "DX20C"
                        strD = "DXZC"
                        strE = "DX20E"
                    Else
                        strA = RTrim(tmptext) & "A"
                        strB = RTrim(tmptext) & "B"
                        strC = RTrim(tmptext) & "C"
                        strD = RTrim(tmptext) & "D"
                        strE = RTrim(tmptext) & "E"
                    End If
                    Dim ii As Integer
                    Dim AX1, AY1, AZ1, AX2, AY2, AZ2, BX1, BY1, BZ1, BX2, BY2, BZ2, CX1, CY1, CZ1, CX2, CY2, CZ2, DX1, DY1, DZ1, DX2, DY2, DZ2, EX1, EY1, EZ1, EX2, EY2, EZ2 As Double
                    For ii = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & ii).Text)) = LTrim(strA) Then
                            AX1 = ash1.Range("B" & ii).Value
                            AY1 = ash1.Range("C" & ii).Value
                            AZ1 = ash1.Range("D" & ii).Value
                            AX2 = ash1.Range("I" & ii).Value
                            AY2 = ash1.Range("J" & ii).Value
                            AZ2 = ash1.Range("K" & ii).Value
                            Exit For
                        End If
                    Next
                    For ii = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & ii).Text)) = LTrim(strB) Then
                            BX1 = ash1.Range("B" & ii).Value
                            BY1 = ash1.Range("C" & ii).Value
                            BZ1 = ash1.Range("D" & ii).Value
                            BX2 = ash1.Range("I" & ii).Value
                            BY2 = ash1.Range("J" & ii).Value
                            BZ2 = ash1.Range("K" & ii).Value
                            Exit For
                        End If
                    Next
                    For ii = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & ii).Text)) = LTrim(strC) Then
                            CX1 = ash1.Range("B" & ii).Value
                            CY1 = ash1.Range("C" & ii).Value
                            CZ1 = ash1.Range("D" & ii).Value
                            CX2 = ash1.Range("I" & ii).Value
                            CY2 = ash1.Range("J" & ii).Value
                            CZ2 = ash1.Range("K" & ii).Value
                            Exit For
                        End If
                    Next
                    For ii = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & ii).Text)) = LTrim(strD) Then
                            DX1 = ash1.Range("B" & ii).Value
                            DY1 = ash1.Range("C" & ii).Value
                            DZ1 = ash1.Range("D" & ii).Value
                            DX2 = ash1.Range("I" & ii).Value
                            DY2 = ash1.Range("J" & ii).Value
                            DZ2 = ash1.Range("K" & ii).Value
                            Exit For
                        End If
                    Next
                    For ii = 1 To nRows1
                        If LTrim(RTrim(ash1.Range("A" & ii).Text)) = LTrim(strE) Then
                            EX1 = ash1.Range("B" & ii).Value
                            EY1 = ash1.Range("C" & ii).Value
                            EZ1 = ash1.Range("D" & ii).Value
                            EX2 = ash1.Range("I" & ii).Value
                            EY2 = ash1.Range("J" & ii).Value
                            EZ2 = ash1.Range("K" & ii).Value
                            Exit For
                        End If
                    Next
                    Dim preBD, nowBD, preAE, nowAE, preAC, nowAC, preEC, nowEC As Double
                    preBD = Math.Sqrt((BX1 - DX1) ^ 2 + (BY1 - DY1) ^ 2 + (BZ1 - DZ1) ^ 2)
                    nowBD = Math.Sqrt((BX2 - DX2) ^ 2 + (BY2 - DY2) ^ 2 + (BZ2 - DZ2) ^ 2)
                    preAE = Math.Sqrt((AX1 - EX1) ^ 2 + (AY1 - EY1) ^ 2 + (AZ1 - EZ1) ^ 2)
                    nowAE = Math.Sqrt((AX2 - EX2) ^ 2 + (AY2 - EY2) ^ 2 + (AZ2 - EZ2) ^ 2)
                    preAC = Math.Sqrt((AX1 - CX1) ^ 2 + (AY1 - CY1) ^ 2 + (AZ1 - CZ1) ^ 2)
                    nowAC = Math.Sqrt((AX2 - CX2) ^ 2 + (AY2 - CY2) ^ 2 + (AZ2 - CZ2) ^ 2)
                    preEC = Math.Sqrt((EX1 - CX1) ^ 2 + (EY1 - CY1) ^ 2 + (EZ1 - CZ1) ^ 2)
                    nowEC = Math.Sqrt((EX2 - CX2) ^ 2 + (EY2 - CY2) ^ 2 + (EZ2 - CZ2) ^ 2)
                    ash3.Range("E" & initrow).Value = Format(nowBD, "0.0000")
                    ash3.Range("E" & initrow + 1).Value = Format(nowAE, "0.0000")
                    ash3.Range("E" & initrow + 2).Value = Format(nowAC, "0.0000")
                    ash3.Range("E" & initrow + 3).Value = Format(nowEC, "0.0000")
                    initrow = initrow + 4
                    tmptext = ash3.Range("A" & initrow).Text
                Loop
            End If
            tabletext = ""
        Next
        MsgBox("报表已生成", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "结果")
    End Sub

#End Region

#Region "帮助器"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
