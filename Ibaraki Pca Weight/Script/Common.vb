﻿Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading.Thread
Imports System.Windows.Forms

Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Check internet connection.
    ''' </summary>
    ''' <returns>Connection state.</returns>
    Private Function IsNetAvail()
        Dim objResp As WebResponse
        Try
            objResp = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse
            objResp.Close()
            objResp = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check update.
    ''' </summary>
    Private Sub ChkUpd()
        HdrSty("アップデートの確認...")
        If IsNetAvail() AndAlso Not (New WebClient).DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
            MsgBox($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", 262144, Title:="更新")
            Dim frmUpd = New FrmUpdate
            frmUpd.ShowDialog()
        End If
    End Sub

    ''' <summary>
    ''' Update valid license.
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Fade in form.
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity += 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Fade out form.
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity -= 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' End process.
    ''' </summary>
    ''' <param name="name">Process name.</param>
    Friend Sub KillPrcs(name As String)
        If GetProcessesByName(name).Count > 0 Then
            For Each item In GetProcessesByName(name)
                item.Kill()
            Next
        End If
    End Sub

    ''' <summary>
    ''' Kill excel.
    ''' </summary>
    Private Sub KillXl()
        Clear()
        HdrSty("警告：このアプリケーションを使用する前に、すべての「エクセル」を閉じてください。「エンター」キーを押して続行します...")
        ReadLine()
        KillPrcs(XL_NAME)
    End Sub

    ''' <summary>
    ''' Run application.
    ''' </summary>
    Friend Sub RunApp()
        ChkUpd()
        KillXl()
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = "「エクセル」ドキュメントを開く",
            .Filter = "「エクセル」ドキュメント|*.xlsx;*.xls"
        }
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim filePath = ofd.FileName
            xlApp.Workbooks.Open(filePath)
            WtIbarakiPca(xlApp)
            xlApp.ActiveWorkbook.Close(SaveChanges:=True)
            Process.Start(filePath)
        End If
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Create directory advanced.
    ''' </summary>
    ''' <param name="path">Folder path.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Delete file advanced.
    ''' </summary>
    ''' <param name="path">File path.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Header Yes/No question (1/0).
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Answer value.</returns>
    Friend Function HdrYNQ(caption As String)
        Dim value = HdrDWrng(caption)
        If value <> 0 Or value <> 1 Then
            Do Until value = 0 Or value = 1
                value = HdrDErr(caption)
            Loop
        End If
        Return value
    End Function

    ''' <summary>
    ''' Direct value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Friend Sub DctVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = value
    End Sub

    ''' <summary>
    ''' Mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Private Sub ModVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = value
        xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
    End Sub

    ''' <summary>
    ''' Direct value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub ClrVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.MergeArea.ClearContents()
    End Sub

    ''' <summary>
    ''' Publish string value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="caption">Caption.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub PubSVal(xlApp As Microsoft.Office.Interop.Excel.Application, caption As String, cell As String)
        DctVal(xlApp, cell, DtlSInp(caption))
    End Sub

    ''' <summary>
    ''' Publish double value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Friend Sub PubDVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Double)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Publish double mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Publish double mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="title">Title rebar.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Publish double mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="d">Diameter.</param>
    ''' <param name="title">Title rebar.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, d As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"S{row}", d)
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Publish double mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="d">Diameter.</param>
    ''' <param name="title">Title rebar.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="price">Price rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubDModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, d As String, title As String, name As String, weight As Double, price As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"S{row}", d)
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            ModVal(xlApp, $"CQ{row}", price)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr.Enabled Then
            tmr.Stop()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' Header style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub HdrSty(caption As String)
        ForegroundColor = DarkYellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' Intro style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub IntroSty(caption As String)
        ForegroundColor = Blue
        Write(caption)
    End Sub

    ''' <summary>
    ''' Title style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub TitSty(caption As String)
        ForegroundColor = Green
        Write(caption)
    End Sub

    ''' <summary>
    ''' Input style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub InpSty(caption As String)
        ForegroundColor = Cyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' Description style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub DescSty(caption As String)
        ForegroundColor = Magenta
        Write(caption)
    End Sub

    ''' <summary>
    ''' Warning style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub WrngSty(caption As String)
        ForegroundColor = Yellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' Error style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub ErrSty(caption As String)
        ForegroundColor = Red
        Write(caption)
    End Sub

    ''' <summary>
    ''' Prefix input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub PrefInp(caption As String)
        InpSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Prefix select.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub PrefSel(caption As String)
        WrngSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Prefix warning.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub PrefWrng(caption As String)
        WrngSty(caption)
        ForegroundColor = Red
    End Sub

    ''' <summary>
    ''' Suffix description.
    ''' </summary>
    ''' <param name="description">Description.</param>
    Private Sub SfxDesc(description As String)
        DescSty(description)
        PrefInp(": ")
    End Sub

    ''' <summary>
    ''' Intro.
    ''' </summary>
    Private Sub Intro()
        Clear()
        IntroSty(My.Resources.gr_name & vbCrLf)
        IntroSty(My.Resources.cc_text & vbCrLf)
        TitSty(vbCrLf & My.Resources.app_true_name & vbCrLf & vbCrLf)
    End Sub

    ''' <summary>
    ''' Header double input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function HdrDInp(caption As String)
        Intro()
        Return DtlDInp(caption)
    End Function

    ''' <summary>
    ''' Header string warning.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub HdrWrng(caption As String)
        Intro()
        WrngSty(caption)
    End Sub

    ''' <summary>
    ''' Header double warning.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Function HdrDWrng(caption As String)
        Intro()
        PrefSel(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Header double error.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Function HdrDErr(caption As String)
        Intro()
        PrefWrng(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Detail double input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlDInp(caption As String)
        PrefInp(caption)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Detail string input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlSInp(caption As String)
        PrefInp(caption)
        Return ReadLine.ToString()
    End Function

    ''' <summary>
    ''' Detail double input description.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <param name="description">Description.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlDInpDesc(caption As String, description As String)
        InpSty(caption)
        SfxDesc(description)
        Return Val(ReadLine)
    End Function
#End Region
End Module
