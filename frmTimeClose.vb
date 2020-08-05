Imports System.IO
Imports System
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Public Class frmTimeClose

    Private Sub btSet_Click(sender As Object, e As EventArgs) Handles btSet.Click
        If MessageBox.Show("คุณต้องการบันทึกข้อมูล ใช่หรือไม่", "Setting", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            WriteIni(frmMain.IniPath, 0, "TIME_STOP", Format(Time_Close.Value, "HH:mm:ss"))
            If CheckBox1.Checked = True Then
                WriteIni(frmMain.IniPath, 0, "AUTO_STOP_ENABLE", "Y")
                WriteIni(frmMain.IniPath, 0, "DATE_STOP", txtDate.Text)
            Else
                WriteIni(frmMain.IniPath, 0, "AUTO_STOP_ENABLE", "N")
            End If

            MessageBox.Show("บันทึกข้อมูลเรียบร้อยแล้ว")
        End If

    End Sub

    Private Sub frmTimeClose_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MessageBox.Show(ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "TIME_STOP", "").ToString())
        Dim vCheckbox As String
        txtDate.Text = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "DATE_STOP", "").ToString()
        Time_Close.Text = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "TIME_STOP", "").ToString()
        vCheckbox = ReadIni(Directory.GetCurrentDirectory() & "/AppStartup.ini", 0, "AUTO_STOP_ENABLE", "").ToString()
        If vCheckbox = "Y" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

    End Sub
End Class