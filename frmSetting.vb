Imports System.Windows.Forms
Imports System.IO
Imports System
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Public Class frmSetting
    Dim fMain As frmMain

    Private Sub btSave_Click(sender As System.Object, e As System.EventArgs) Handles btSave.Click
        If MessageBox.Show("คุณต้องการบันทึกข้อมูล ใช่หรือไม่", "Setting", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            If SaveSetting() Then
                fMain.UpdateDatagridView()
                MessageBox.Show("บันทึกข้อมูลเรียบร้อยแล้ว")
            End If
        End If
    End Sub

    Public Sub InitialDatagridview(ByVal pForm As frmMain)
        Dim i As Integer
        Dim vCheckBox As New DataGridViewCheckBoxCell

        'dgvMain.Columns.Clear()
        'dgvMain.ColumnCount = 3
        'dgvMain.Columns(0).Name = "#"
        'dgvMain.Columns(1).Name = "Application Name"
        'dgvMain.Columns(2).Name = "Setting"
        fMain = pForm

        With dgvSetting
            .Visible = False
            .RowCount = fMain.totalApp

            For i = 0 To .RowCount - 1
                .Rows(i).Height = 50
                .Rows(i).Cells(0).Value = i + 1
                .Rows(i).Cells(1).Value = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", i + 1, "TITLE", "").ToString()
                vCheckBox.Value = IIf(fMain.appStartup(i)._AutoStart = "N", False, True)
                .Rows(i).Cells(2).Value = vCheckBox.Value
                .Rows(i).Cells(2).ToolTipText = "Auto Start"
                .Rows(i).Cells(3).Value = IIf(vCheckBox.Value = True, "Auto Start", "N/A")
            Next
            .Visible = True

        End With
    End Sub

    Function SaveSetting() As Boolean
        Dim vCheckBox As New DataGridViewCheckBoxCell
        Dim vRet As Boolean = False
        Try
            For i As Integer = 0 To fMain.appStartup.Length - 1
                vCheckBox.Value = dgvSetting.Item(2, i).Value
                fMain.appStartup(i)._AutoStart = IIf(vCheckBox.Value = True, "Y", "N")
                WriteIni(fMain.IniPath, i + 1, "AUTO_START", fMain.appStartup(i)._AutoStart)
            Next
            vRet = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return vRet
    End Function

    Private Sub frmSetting_Load(sender As Object, e As System.EventArgs) Handles Me.Load
    End Sub

    Private Sub dgvSetting_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSetting.CellClick
        Dim vCheckBox As New DataGridViewCheckBoxCell
        If e.RowIndex >= 0 And e.ColumnIndex = 2 Then
            vCheckBox = dgvSetting.Item(2, e.RowIndex)
            vCheckBox.Value = Not vCheckBox.Value
            dgvSetting.Item(2, e.RowIndex).Value = vCheckBox.Value
            dgvSetting.Item(3, e.RowIndex).Value = IIf(vCheckBox.Value = True, "Auto Start", "N/A")
        End If
    End Sub

End Class