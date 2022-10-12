﻿Imports System.Diagnostics.Process
Imports System.Math
Imports System.Net
Imports System.Windows.Forms

Public Class FrmUpdate
#Region "Fields"
    Private ReadOnly _wc As New WebClient
#End Region

#Region "Events"
    ' Form shown
    Private Sub FrmUpdate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        pnlProgressBar.Width = 1
        FIFrm()
        CrtDirAdv(FRNT_PATH)
        DelFileAdv(FILE_SETUP_ADR)
        tmrMain.StrtAdv()
        AddHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        _wc.DownloadFileAsync(New Uri(_wc.DownloadString(My.Resources.link_app)), FILE_SETUP_ADR)
    End Sub

    ' Update download progress
    Private Sub Upd_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        lblCapacity.Text = String.Format("{0} MB / {1} MB", (e.BytesReceived / 1024D / 1024D).ToString("0.00"), (e.TotalBytesToReceive / 1024D / 1024D).ToString("0.00"))
        lblPercent.Text = $"{e.ProgressPercentage}%"
        pnlProgressBar.Width = CInt(Ceiling(e.ProgressPercentage * Width / 100D))
    End Sub

    ' Timer main
    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles tmrMain.Tick

    End Sub

    ' Form closing
    Private Sub FrmUpdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If GetProcessesByName(My.Resources.app_name).Count > 0 Then
            For Each item In GetProcessesByName(My.Resources.app_name)
                item.Kill()
            Next
        End If
        If GetProcessesByName(XL_NAME).Count > 0 Then
            For Each item In GetProcessesByName(XL_NAME)
                item.Kill()
            Next
        End If
    End Sub
#End Region
End Class