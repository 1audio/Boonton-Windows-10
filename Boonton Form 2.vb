
'Boonton Audio Controller - Automated GPIB Data Collection for the Boonton 1120/1121
'Copyright (C) 2006  Shannon Parks    email: separks@diytube.com
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

Imports NationalInstruments.NI4882  'must be included to reference the LangInt assembly
Imports System.IO
Imports System.Web

Imports System.Runtime.Serialization.Formatters.Binary

Imports System.Windows.Forms.DataVisualization.Charting


Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim CleanupFlag As Byte = 0
    Dim BoontonController As BoontonInterface

    'Dim buffer As String
    'Dim Dev As Short
    'Dim CalLab(27) As Double
    'Dim CalTHD(27) As Double
    'Dim Freq(27) As Double
    Friend WithEvents txtStartLevelV As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtLevelOhmLoad As System.Windows.Forms.TextBox
    Friend WithEvents chkFreqDistActive As System.Windows.Forms.CheckBox
    Friend WithEvents chkSNRActive As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents chkFullBandwidth As System.Windows.Forms.CheckBox
    Friend WithEvents chkInverseRIAA As System.Windows.Forms.CheckBox
    Friend WithEvents EndTestButton As System.Windows.Forms.Button
    Friend WithEvents QuitButton As System.Windows.Forms.Button
    Friend WithEvents tbGPIBAddress As System.Windows.Forms.TextBox
    Friend WithEvents tbRemoteAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DisableOptionalFiltersCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents chk318uS As System.Windows.Forms.CheckBox
    Friend WithEvents chkLevelSweepActive As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtLevelDistortionThreshDB As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtFreqSourceV As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSNRReferenceV As System.Windows.Forms.TextBox
    Friend WithEvents lbLevelThresh As System.Windows.Forms.Label
    Friend WithEvents txtStartLevelDB As TextBox
    Friend WithEvents txtSNRReferenceDB As TextBox
    Friend WithEvents txtFreqSourceDB As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents txtSourceMaxV As TextBox
    Friend WithEvents txtLevelDistortionThreshV As TextBox
    Friend WithEvents txtSourceMaxDB As TextBox
    'the string returned from instrument
    Dim ResByte As Integer
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)

    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents RunButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim ChartArea3 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim ChartArea4 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend2 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series3 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series4 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Title2 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Me.RunButton = New System.Windows.Forms.Button()
        Me.txtStartLevelV = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtLevelOhmLoad = New System.Windows.Forms.TextBox()
        Me.chkFreqDistActive = New System.Windows.Forms.CheckBox()
        Me.chkSNRActive = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.chkFullBandwidth = New System.Windows.Forms.CheckBox()
        Me.chkInverseRIAA = New System.Windows.Forms.CheckBox()
        Me.EndTestButton = New System.Windows.Forms.Button()
        Me.QuitButton = New System.Windows.Forms.Button()
        Me.tbGPIBAddress = New System.Windows.Forms.TextBox()
        Me.tbRemoteAddress = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.DisableOptionalFiltersCheckbox = New System.Windows.Forms.CheckBox()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.chk318uS = New System.Windows.Forms.CheckBox()
        Me.chkLevelSweepActive = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtLevelDistortionThreshDB = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFreqSourceV = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSNRReferenceV = New System.Windows.Forms.TextBox()
        Me.lbLevelThresh = New System.Windows.Forms.Label()
        Me.txtStartLevelDB = New System.Windows.Forms.TextBox()
        Me.txtSNRReferenceDB = New System.Windows.Forms.TextBox()
        Me.txtFreqSourceDB = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtSourceMaxV = New System.Windows.Forms.TextBox()
        Me.txtLevelDistortionThreshV = New System.Windows.Forms.TextBox()
        Me.txtSourceMaxDB = New System.Windows.Forms.TextBox()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RunButton
        '
        Me.RunButton.Location = New System.Drawing.Point(33, 582)
        Me.RunButton.Name = "RunButton"
        Me.RunButton.Size = New System.Drawing.Size(80, 30)
        Me.RunButton.TabIndex = 1
        Me.RunButton.Text = "Run"
        '
        'txtStartLevelV
        '
        Me.txtStartLevelV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartLevelV.Location = New System.Drawing.Point(129, 71)
        Me.txtStartLevelV.Name = "txtStartLevelV"
        Me.txtStartLevelV.Size = New System.Drawing.Size(53, 20)
        Me.txtStartLevelV.TabIndex = 6
        Me.txtStartLevelV.Text = "0.200"
        Me.txtStartLevelV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Crimson
        Me.Label1.Location = New System.Drawing.Point(56, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Start Level"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(65, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Ohm Load"
        '
        'txtLevelOhmLoad
        '
        Me.txtLevelOhmLoad.Location = New System.Drawing.Point(129, 149)
        Me.txtLevelOhmLoad.Name = "txtLevelOhmLoad"
        Me.txtLevelOhmLoad.Size = New System.Drawing.Size(53, 20)
        Me.txtLevelOhmLoad.TabIndex = 8
        Me.txtLevelOhmLoad.Text = "8"
        Me.txtLevelOhmLoad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chkFreqDistActive
        '
        Me.chkFreqDistActive.AutoSize = True
        Me.chkFreqDistActive.Checked = True
        Me.chkFreqDistActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFreqDistActive.Location = New System.Drawing.Point(15, 292)
        Me.chkFreqDistActive.Name = "chkFreqDistActive"
        Me.chkFreqDistActive.Size = New System.Drawing.Size(196, 17)
        Me.chkFreqDistActive.TabIndex = 10
        Me.chkFreqDistActive.Text = "Frequency And Distortion Response"
        Me.chkFreqDistActive.UseVisualStyleBackColor = True
        '
        'chkSNRActive
        '
        Me.chkSNRActive.AutoSize = True
        Me.chkSNRActive.Checked = True
        Me.chkSNRActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSNRActive.Location = New System.Drawing.Point(33, 208)
        Me.chkSNRActive.Name = "chkSNRActive"
        Me.chkSNRActive.Size = New System.Drawing.Size(49, 17)
        Me.chkSNRActive.TabIndex = 12
        Me.chkSNRActive.Text = "SNR"
        Me.chkSNRActive.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 445)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Description"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(84, 438)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(175, 20)
        Me.txtDescription.TabIndex = 13
        Me.txtDescription.Text = "Basic Test"
        Me.txtDescription.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chkFullBandwidth
        '
        Me.chkFullBandwidth.AutoSize = True
        Me.chkFullBandwidth.Checked = True
        Me.chkFullBandwidth.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFullBandwidth.Location = New System.Drawing.Point(51, 356)
        Me.chkFullBandwidth.Name = "chkFullBandwidth"
        Me.chkFullBandwidth.Size = New System.Drawing.Size(103, 17)
        Me.chkFullBandwidth.TabIndex = 23
        Me.chkFullBandwidth.Text = "10Hz to 140kHz"
        Me.chkFullBandwidth.UseVisualStyleBackColor = True
        '
        'chkInverseRIAA
        '
        Me.chkInverseRIAA.AutoSize = True
        Me.chkInverseRIAA.Location = New System.Drawing.Point(51, 379)
        Me.chkInverseRIAA.Name = "chkInverseRIAA"
        Me.chkInverseRIAA.Size = New System.Drawing.Size(105, 17)
        Me.chkInverseRIAA.TabIndex = 24
        Me.chkInverseRIAA.Text = "Inverse RIAA Eq"
        Me.chkInverseRIAA.UseVisualStyleBackColor = True
        '
        'EndTestButton
        '
        Me.EndTestButton.Enabled = False
        Me.EndTestButton.Location = New System.Drawing.Point(33, 627)
        Me.EndTestButton.Name = "EndTestButton"
        Me.EndTestButton.Size = New System.Drawing.Size(80, 30)
        Me.EndTestButton.TabIndex = 25
        Me.EndTestButton.Text = "End Test"
        '
        'QuitButton
        '
        Me.QuitButton.Location = New System.Drawing.Point(159, 627)
        Me.QuitButton.Name = "QuitButton"
        Me.QuitButton.Size = New System.Drawing.Size(80, 30)
        Me.QuitButton.TabIndex = 26
        Me.QuitButton.Text = "Quit"
        '
        'tbGPIBAddress
        '
        Me.tbGPIBAddress.Location = New System.Drawing.Point(126, 464)
        Me.tbGPIBAddress.Name = "tbGPIBAddress"
        Me.tbGPIBAddress.Size = New System.Drawing.Size(53, 20)
        Me.tbGPIBAddress.TabIndex = 29
        Me.tbGPIBAddress.Text = "1"
        Me.tbGPIBAddress.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbRemoteAddress
        '
        Me.tbRemoteAddress.Location = New System.Drawing.Point(126, 490)
        Me.tbRemoteAddress.Name = "tbRemoteAddress"
        Me.tbRemoteAddress.Size = New System.Drawing.Size(53, 20)
        Me.tbRemoteAddress.TabIndex = 30
        Me.tbRemoteAddress.Text = "15"
        Me.tbRemoteAddress.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(39, 471)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(73, 13)
        Me.Label12.TabIndex = 31
        Me.Label12.Text = "GPIB Address"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(30, 497)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(85, 13)
        Me.Label13.TabIndex = 32
        Me.Label13.Text = "Remote Address"
        '
        'DisableOptionalFiltersCheckbox
        '
        Me.DisableOptionalFiltersCheckbox.AutoSize = True
        Me.DisableOptionalFiltersCheckbox.Checked = True
        Me.DisableOptionalFiltersCheckbox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DisableOptionalFiltersCheckbox.Location = New System.Drawing.Point(33, 541)
        Me.DisableOptionalFiltersCheckbox.Name = "DisableOptionalFiltersCheckbox"
        Me.DisableOptionalFiltersCheckbox.Size = New System.Drawing.Size(133, 17)
        Me.DisableOptionalFiltersCheckbox.TabIndex = 38
        Me.DisableOptionalFiltersCheckbox.Text = "Disable Optional Filters"
        Me.DisableOptionalFiltersCheckbox.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        ChartArea3.Name = "Level"
        ChartArea4.Name = "Distortion"
        Me.Chart1.ChartAreas.Add(ChartArea3)
        Me.Chart1.ChartAreas.Add(ChartArea4)
        Legend2.Enabled = False
        Legend2.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend2)
        Me.Chart1.Location = New System.Drawing.Point(265, 12)
        Me.Chart1.Name = "Chart1"
        Series3.ChartArea = "Level"
        Series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series3.Legend = "Legend1"
        Series3.Name = "Series1"
        Series4.ChartArea = "Distortion"
        Series4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series4.Legend = "Legend1"
        Series4.Name = "Series2"
        Me.Chart1.Series.Add(Series3)
        Me.Chart1.Series.Add(Series4)
        Me.Chart1.Size = New System.Drawing.Size(1026, 645)
        Me.Chart1.TabIndex = 39
        Me.Chart1.Text = "Chart1"
        Title2.Name = "Basic Test"
        Title2.Text = "Basic Test"
        Me.Chart1.Titles.Add(Title2)
        '
        'chk318uS
        '
        Me.chk318uS.AutoSize = True
        Me.chk318uS.Location = New System.Drawing.Point(51, 402)
        Me.chk318uS.Name = "chk318uS"
        Me.chk318uS.Size = New System.Drawing.Size(93, 17)
        Me.chk318uS.TabIndex = 40
        Me.chk318uS.Text = "3.18uS Adl. tc"
        Me.chk318uS.UseVisualStyleBackColor = True
        '
        'chkLevelSweepActive
        '
        Me.chkLevelSweepActive.AutoSize = True
        Me.chkLevelSweepActive.Checked = True
        Me.chkLevelSweepActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkLevelSweepActive.Location = New System.Drawing.Point(33, 42)
        Me.chkLevelSweepActive.Name = "chkLevelSweepActive"
        Me.chkLevelSweepActive.Size = New System.Drawing.Size(88, 17)
        Me.chkLevelSweepActive.TabIndex = 41
        Me.chkLevelSweepActive.Text = "Level Sweep"
        Me.chkLevelSweepActive.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(22, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(101, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Distortion Threshold"
        '
        'txtLevelDistortionThreshDB
        '
        Me.txtLevelDistortionThreshDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.txtLevelDistortionThreshDB.Location = New System.Drawing.Point(188, 97)
        Me.txtLevelDistortionThreshDB.Name = "txtLevelDistortionThreshDB"
        Me.txtLevelDistortionThreshDB.Size = New System.Drawing.Size(53, 20)
        Me.txtLevelDistortionThreshDB.TabIndex = 44
        Me.txtLevelDistortionThreshDB.Text = "0"
        Me.txtLevelDistortionThreshDB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Crimson
        Me.Label5.Location = New System.Drawing.Point(74, 338)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 13)
        Me.Label5.TabIndex = 47
        Me.Label5.Text = "Source"
        '
        'txtFreqSourceV
        '
        Me.txtFreqSourceV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFreqSourceV.Location = New System.Drawing.Point(127, 335)
        Me.txtFreqSourceV.Name = "txtFreqSourceV"
        Me.txtFreqSourceV.Size = New System.Drawing.Size(53, 20)
        Me.txtFreqSourceV.TabIndex = 46
        Me.txtFreqSourceV.Text = "0.200"
        Me.txtFreqSourceV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Crimson
        Me.Label6.Location = New System.Drawing.Point(59, 243)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 13)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Reference"
        '
        'txtSNRReferenceV
        '
        Me.txtSNRReferenceV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSNRReferenceV.Location = New System.Drawing.Point(126, 240)
        Me.txtSNRReferenceV.Name = "txtSNRReferenceV"
        Me.txtSNRReferenceV.Size = New System.Drawing.Size(53, 20)
        Me.txtSNRReferenceV.TabIndex = 48
        Me.txtSNRReferenceV.Text = "0.200"
        Me.txtSNRReferenceV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbLevelThresh
        '
        Me.lbLevelThresh.AutoSize = True
        Me.lbLevelThresh.Location = New System.Drawing.Point(126, 195)
        Me.lbLevelThresh.Name = "lbLevelThresh"
        Me.lbLevelThresh.Size = New System.Drawing.Size(0, 13)
        Me.lbLevelThresh.TabIndex = 50
        '
        'txtStartLevelDB
        '
        Me.txtStartLevelDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartLevelDB.Location = New System.Drawing.Point(188, 71)
        Me.txtStartLevelDB.Name = "txtStartLevelDB"
        Me.txtStartLevelDB.Size = New System.Drawing.Size(53, 20)
        Me.txtStartLevelDB.TabIndex = 51
        Me.txtStartLevelDB.Text = "-13.98"
        Me.txtStartLevelDB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtSNRReferenceDB
        '
        Me.txtSNRReferenceDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSNRReferenceDB.Location = New System.Drawing.Point(185, 240)
        Me.txtSNRReferenceDB.Name = "txtSNRReferenceDB"
        Me.txtSNRReferenceDB.Size = New System.Drawing.Size(53, 20)
        Me.txtSNRReferenceDB.TabIndex = 52
        Me.txtSNRReferenceDB.Text = "-13.98"
        Me.txtSNRReferenceDB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtFreqSourceDB
        '
        Me.txtFreqSourceDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFreqSourceDB.Location = New System.Drawing.Point(188, 335)
        Me.txtFreqSourceDB.Name = "txtFreqSourceDB"
        Me.txtFreqSourceDB.Size = New System.Drawing.Size(53, 20)
        Me.txtFreqSourceDB.TabIndex = 53
        Me.txtFreqSourceDB.Text = "-13.98"
        Me.txtFreqSourceDB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Crimson
        Me.Label7.Location = New System.Drawing.Point(137, 55)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(42, 13)
        Me.Label7.TabIndex = 54
        Me.Label7.Text = "VRMS"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Crimson
        Me.Label8.Location = New System.Drawing.Point(188, 55)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 13)
        Me.Label8.TabIndex = 55
        Me.Label8.Text = "DB"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Crimson
        Me.Label9.Location = New System.Drawing.Point(188, 224)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 13)
        Me.Label9.TabIndex = 57
        Me.Label9.Text = "DB"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Crimson
        Me.Label10.Location = New System.Drawing.Point(137, 224)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(42, 13)
        Me.Label10.TabIndex = 56
        Me.Label10.Text = "VRMS"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Crimson
        Me.Label14.Location = New System.Drawing.Point(191, 319)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(24, 13)
        Me.Label14.TabIndex = 59
        Me.Label14.Text = "DB"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Crimson
        Me.Label15.Location = New System.Drawing.Point(140, 319)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(42, 13)
        Me.Label15.TabIndex = 58
        Me.Label15.Text = "VRMS"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(22, 123)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(93, 13)
        Me.Label11.TabIndex = 61
        Me.Label11.Text = "Source Max Level"
        '
        'txtSourceMaxV
        '
        Me.txtSourceMaxV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.txtSourceMaxV.Location = New System.Drawing.Point(129, 123)
        Me.txtSourceMaxV.Name = "txtSourceMaxV"
        Me.txtSourceMaxV.Size = New System.Drawing.Size(53, 20)
        Me.txtSourceMaxV.TabIndex = 60
        Me.txtSourceMaxV.Text = "10"
        Me.txtSourceMaxV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLevelDistortionThreshV
        '
        Me.txtLevelDistortionThreshV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.txtLevelDistortionThreshV.Location = New System.Drawing.Point(129, 97)
        Me.txtLevelDistortionThreshV.Name = "txtLevelDistortionThreshV"
        Me.txtLevelDistortionThreshV.Size = New System.Drawing.Size(53, 20)
        Me.txtLevelDistortionThreshV.TabIndex = 62
        Me.txtLevelDistortionThreshV.Text = "1"
        Me.txtLevelDistortionThreshV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtSourceMaxDB
        '
        Me.txtSourceMaxDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.txtSourceMaxDB.Location = New System.Drawing.Point(188, 123)
        Me.txtSourceMaxDB.Name = "txtSourceMaxDB"
        Me.txtSourceMaxDB.Size = New System.Drawing.Size(53, 20)
        Me.txtSourceMaxDB.TabIndex = 63
        Me.txtSourceMaxDB.Text = "20"
        Me.txtSourceMaxDB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1316, 669)
        Me.Controls.Add(Me.txtSourceMaxDB)
        Me.Controls.Add(Me.txtLevelDistortionThreshV)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtSourceMaxV)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtFreqSourceDB)
        Me.Controls.Add(Me.txtSNRReferenceDB)
        Me.Controls.Add(Me.txtStartLevelDB)
        Me.Controls.Add(Me.lbLevelThresh)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtSNRReferenceV)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtFreqSourceV)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtLevelDistortionThreshDB)
        Me.Controls.Add(Me.chkLevelSweepActive)
        Me.Controls.Add(Me.chk318uS)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.DisableOptionalFiltersCheckbox)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.tbRemoteAddress)
        Me.Controls.Add(Me.tbGPIBAddress)
        Me.Controls.Add(Me.QuitButton)
        Me.Controls.Add(Me.EndTestButton)
        Me.Controls.Add(Me.chkInverseRIAA)
        Me.Controls.Add(Me.chkFullBandwidth)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.chkSNRActive)
        Me.Controls.Add(Me.chkFreqDistActive)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtLevelOhmLoad)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtStartLevelV)
        Me.Controls.Add(Me.RunButton)
        Me.Name = "Form1"
        Me.Text = "Boonton Precision Audio"
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub GPIBCleanup(ByRef msg As String)
        Dim ErrorMnemonic() As String = {"EDVR", "ECIC", "ENOL", "EADR", "EARG", "ESAC", "EABO", "ENEB", "EDMA", "", "EOIP", "ECAP", "EFSO", "", "EBUS", "ESTB", "ESRQ", "", "", "", "ETAB"}

        ' After each GPIB call, the application checks whether the call
        ' succeeded. If an NI-488.2 call fails, the GPIB driver sets the
        ' corresponding bit in the global status variable. If the call
        ' failed, this procedure prints an error message, takes the device
        ' offline and exits.
        If (IgnoredErrors.IndexOf(CInt(BoontonController.GetErrorNumber())) <> -1) Then Return
        Dim Result As MsgBoxResult = MsgBox(BoontonController.GetErrorString() + " Click Yes to ignore further errors of this type, No to reprompt if this type of error appears, and Cancel to shut down", MsgBoxStyle.YesNoCancel, "Error " & CStr(BoontonController.GetErrorNumber()))

        Select Case Result
            Case MsgBoxResult.Yes
                IgnoredErrors.Add(BoontonController.GetErrorNumber())
            Case MsgBoxResult.Cancel  'close reference to the instrument
                BoontonController.Close() 'end program
                Me.Close()
                End
        End Select
    End Sub

    Dim IgnoredErrors As New ArrayList()

    Dim FreqLevelSeries As New Series
    Dim FreqDistortionSeries As New Series
    Dim LevelSweepSeries As New Series
    Dim SNRPercentageSeries As New Series

    Dim FreqLevelArea As New ChartArea
    Dim FreqDistortionArea As New ChartArea
    Dim LevelSweepArea As New ChartArea
    Dim SNRPercentageArea As New ChartArea

    'Brent Run
    Private Sub RunButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunButton.Click

        'RUN, QUIT buttons disabled & END TEST button enabled during the test
        RunButton.Enabled = False
        QuitButton.Enabled = False
        EndTestButton.Enabled = True


        Dim BDINDEX As Integer = Val(tbGPIBAddress.Text)           'GPIB Board Address
        Dim PRIMARY_ADDR As Integer = Val(tbRemoteAddress.Text)    'Remote Instrument Address

        BoontonController = New BoontonInterface(BDINDEX, PRIMARY_ADDR)

        If (IsNothing(BoontonController)) Then
            Throw New System.Exception("Something went wrong, check your constructor")
        End If

        If (chkLevelSweepActive.Checked And Not CleanupFlag) Then
            LevelSweepTest()
            If (chkSNRActive.Checked And Not CleanupFlag) Then
                SNRTest()
            End If
        End If

        If (chkFreqDistActive.Checked And Not CleanupFlag) Then
            FreqDistResponseTest()
        End If


        ' Take the device offline and make sure there's no output
        BoontonController.Close()

        If ((chkLevelSweepActive.Checked Or chkSNRActive.Checked Or chkFreqDistActive.Checked) And Not CleanupFlag) Then
            SaveRawData()
            SaveChartImage()
        End If

        'Cleanup


        'supposedly this frees up the object memory completely - voodoo stuff
        GC.Collect()
        GC.WaitForPendingFinalizers()


        'Enable user inputs.
        RunButton.Enabled = True
        QuitButton.Enabled = True
        EndTestButton.Enabled = False

        CleanupFlag = 0
    End Sub


    Private Sub LevelSweepTest()

        Dim OutputDB As Double
        Dim DBThresh As Double
        Dim Distortion As Double
        Dim Level As Double
        Dim IncrementByOneFlag As Boolean
        Dim MaxDB As Double


        Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()
        LevelSweepSeries.Points.Clear()

        Chart1.ChartAreas.Add(LevelSweepArea)
        Chart1.Series.Add(LevelSweepSeries)

        ' ========================================================================
        '
        '  MAIN BODY SECTION
        '
        '  In this application, the Main Body communicates with the instrument
        '  by writing a command to it and reading its response. 
        ' ========================================================================


        ' set source frequency and give a 500ms settling pause
        If (Not BoontonController.SetFloatingInput(True)) Then GPIBCleanup("Floating issue")
        If (Not BoontonController.SetFilters(0, 0)) Then GPIBCleanup("Floating issue")
        If (Not BoontonController.SetFrequency(1000)) Then GPIBCleanup("Floating issue")

        System.Threading.Thread.Sleep((500))
        OutputDB = CDbl(txtStartLevelDB.Text)
        DBThresh = CDbl(txtLevelDistortionThreshDB.Text)
        MaxDB = CDbl(txtSourceMaxDB.Text)

        'main loop
        IncrementByOneFlag = False
        Distortion = DBThresh - 1 ' Gives starting value lesser than dblFindTHD
        Do Until Distortion >= DBThresh And IncrementByOneFlag

            If (Distortion >= DBThresh) Then
                IncrementByOneFlag = True
                OutputDB -= 5
            End If

            If (OutputDB > MaxDB) Then
                Chart1.Titles(0).Text = Chart1.Titles(0).Text + " WARNING: Maximum level was reached on level sweep."
                Exit Do
            End If

            'sets output voltage at either normal level
            If (Not BoontonController.SetSourceDB(OutputDB)) Then GPIBCleanup("Error Setting Source")
            System.Threading.Thread.Sleep(5000)
            'sets 'DIST' measure mode, wait one second for settling and 
            'then reads Boonton buffer
            Distortion = BoontonController.MeasureTHDN()
            If (Double.IsNaN(Distortion)) Then GPIBCleanup("Error Collecting Distortion")
            Level = BoontonController.MeasureVRMS()
            If (Double.IsNaN(Level)) Then GPIBCleanup("Error Collecting Level")


            LevelSweepSeries.Points.AddXY(CStr(Level), Distortion)



            OutputDB += If(IncrementByOneFlag, 1, 3)




            'this allows other events, such as a hard escape
            System.Windows.Forms.Application.DoEvents()
            If CleanupFlag = 1 Then Return
        Loop
        'sets up 'LEVEL' mode, waits one second for settling and then reads Boonton buffer

        'Give measurement in dBs
        lbLevelThresh.Text = CStr(BoontonController.MeasureVRMS())



        txtStartLevelV.Text = System.Math.Round(OutputDB - 1, 3)
    End Sub


    Private Sub SNRTest()
        Dim testLevel As Double = CDbl(LevelSweepSeries.Points.Item(LevelSweepSeries.Points.Count - 1).XValue)
        BoontonController.SetZero()

        Dim Freq As Integer() = {20000, 30000, 30000, 30000}
        Dim FSet() As Integer = {0, 0, 1, 2}

        SNRPercentageSeries.Points.Clear()
        Chart1.ChartAreas.Add(SNRPercentageArea)
        Chart1.Series.Add(SNRPercentageSeries)


        For i As Integer = 0 To Freq.Length - 1
            If (Not BoontonController.SetFilters(If(DisableOptionalFiltersCheckbox.Checked, 0, FSet(i)), 0)) Then
                GPIBCleanup("Error setting the filters")
            End If

            ' set source frequency and give a 500ms settling pause
            If (Not BoontonController.SetFrequency(Freq(i))) Then
                GPIBCleanup("Error setting frequency or source")
            End If

            System.Threading.Thread.Sleep(2000)

            Dim Distortion As Double = BoontonController.MeasureVRMS()

            Dim text As String = CStr(Freq(i))
            If (FSet(i) <> 0) Then
                text = text & "F" & CStr(FSet(i))
            End If

            SNRPercentageSeries.Points.AddXY(text, testLevel + Distortion)


            'this allows other events, such as a hard escape
            System.Windows.Forms.Application.DoEvents()
            If CleanupFlag = 1 Then Return
        Next

    End Sub

    Private Sub FreqDistResponseTest()
        Dim Start, Finish As Byte
        Dim Sweep As Byte

        Dim Freq() As Integer = {10, 13, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500,
            630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000,
            25000, 32000, 40000, 50000, 64000, 80000, 100000, 128000, 140000}

        'Charts for which filters should be used at various frequencies
        Dim FSet() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                             1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
                            1, 1, 1, 1, 1, 1, 1, 1, 1, 1}



        Dim LSet() As Integer = {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
                                2, 2, 2, 2, 3, 3, 3, 3, 3,
                                0, 0, 0, 0, 0, 0}


        If (Not BoontonController.SetFloatingInput(True)) Then Call GPIBCleanup("Error setting floating input")

        FreqLevelSeries.Points.Clear()
        FreqDistortionSeries.Points.Clear()

        Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()

        Chart1.ChartAreas.Add(FreqLevelArea)
        Chart1.ChartAreas.Add(FreqDistortionArea)

        Chart1.Series.Add(FreqLevelSeries)
        Chart1.Series.Add(FreqDistortionSeries)




        'this selects the start and stop point within the Freq() array
        'eg default operation scans from 20Hz (Freq(2)) to 20kHz (Freq(15))
        'note: full bandwidth cannot be used with the inverse RIAA
        If chkFullBandwidth.CheckState = 1 And chkInverseRIAA.CheckState = 0 Then
            Start = 0
            Finish = 42
        Else
            Start = 3
            Finish = 33
        End If

        ' ========================================================================
        '
        '  MAIN BODY SECTION
        '
        '  In this application, the Main Body communicates with the instrument
        '  by writing a command to it and reading its response. 
        ' ========================================================================


        For Sweep = Start To Finish

            ' set proper filters
            If (Not BoontonController.SetFilters(If(DisableOptionalFiltersCheckbox.Checked, 0, FSet(Sweep)), LSet(Sweep))) Then
                GPIBCleanup("Error setting the filters")
            End If

            ' set source frequency and give a 500ms settling pause
            If (Not BoontonController.SetFrequency((Freq(Sweep)))) Then
                GPIBCleanup("Error setting frequency or source")
            End If

            If (Not BoontonController.SetSource(Val(txtFreqSourceV.Text), chkInverseRIAA.Checked, chk318uS.Checked)) Then
                GPIBCleanup("Err setting source")
            End If

            If (Not BoontonController.Tune(Freq(Sweep))) Then
                GPIBCleanup("Error tuning")
            End If

            System.Threading.Thread.Sleep(500)

            Dim dB As Double = BoontonController.MeasureVRMS()
            If (Double.IsNaN(dB)) Then Call GPIBCleanup("Error trying to collect Level")

            Dim Distortion As Double = BoontonController.MeasureTHDN()
            If (Double.IsNaN(Distortion)) Then Call GPIBCleanup("Error trying to collect Distortion")


            'Brent Chart
            If Double.IsNegativeInfinity(dB) Then
                dB = -40
            End If
            FreqLevelSeries.Points.AddXY(CStr(Freq(Sweep)), dB)

            If Double.IsNegativeInfinity(Distortion) Then
                Distortion = -200
            End If
            FreqDistortionSeries.Points.AddXY(CStr(Freq(Sweep)), Distortion)



            'this allows other events, such as a hard escape
            System.Windows.Forms.Application.DoEvents()
            If CleanupFlag = 1 Then Return
        Next
    End Sub

    Private Sub SaveChartImage()
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.ValidateNames = True
        saveFileDialog1.Filter = "JPG Image|*.jpg|PNG Image|*.png|Gif Image|*.gif"
        saveFileDialog1.Title = "Save charts as an image."
        If (saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim name As String = saveFileDialog1.FileName
            Dim extension As String = name.Substring(name.Length - 3, 3)

            Dim format As Imaging.ImageFormat

            If extension.Equals("gif") Then
                format = Imaging.ImageFormat.Gif
            ElseIf extension.Equals("png") Then
                format = Imaging.ImageFormat.Png
            Else
                format = Imaging.ImageFormat.Jpeg
            End If

            'Save current values from the chart object
            Dim chartTitle As String = Chart1.Titles(0).Text
            Dim chartHeight As Integer = Chart1.Height
            Dim chartWidth As Integer = Chart1.Width
            Dim count As Integer = 0


            'Change Chart For Saving to file
            'It will:
            '   Expand or Contract to fit the # of graphs made on this test
            '   Add a timestamp to the title

            ' Fill Areas
            Chart1.ChartAreas.Clear()
            Chart1.Series.Clear()

            If (chkLevelSweepActive.Checked) Then
                count += 1
                Chart1.ChartAreas.Add(LevelSweepArea)
                Chart1.Series.Add(LevelSweepSeries)
            End If

            If (chkSNRActive.Checked) Then
                count += 1
                Chart1.ChartAreas.Add(SNRPercentageArea)
                Chart1.Series.Add(SNRPercentageSeries)
            End If

            If (chkFreqDistActive.Checked) Then
                count += 2
                Chart1.ChartAreas.Add(FreqLevelArea)
                Chart1.ChartAreas.Add(FreqDistortionArea)
                Chart1.Series.Add(FreqLevelSeries)
                Chart1.Series.Add(FreqDistortionSeries)
            End If
            ' Current Chart Height is designed for 2 graphs. Expand/contract if needed.
            Chart1.Height = chartHeight * count / 2

            If (count = 4) Then
                Chart1.Width = chartWidth * 2
                Chart1.Height = chartHeight * 2
            End If



            Chart1.Titles(0).Text = chartTitle & " collected on " & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")



            'Save data
            Chart1.SaveImage(saveFileDialog1.FileName, format)



            'Revert Chart Back to Original State
            Chart1.Titles(0).Text = chartTitle
            Chart1.Height = chartHeight
            Chart1.Width = chartWidth
            Chart1.ChartAreas.Clear()
            Chart1.Series.Clear()
        End If
    End Sub

    Private Sub SaveRawData()
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.ValidateNames = True
        saveFileDialog1.Filter = "Comma Separated Values|*.csv|Tab Separated Values|*.tsv|Plain Text|*.txt|Other|*"
        saveFileDialog1.Title = "Save raw data from this test"
        If (saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then

            Dim name As String = saveFileDialog1.FileName
            Dim extension As String = name.Substring(name.Length - 3, 3)

            Dim separator As String = If(extension.Equals("csv"), ",", vbTab)


            Dim TextFile As New StreamWriter(saveFileDialog1.FileName)
            TextFile.WriteLine("Level Sweep Test" & separator & separator & "SNR Test" & separator & separator & "Frequency Distortion Response Test")

            Dim i As Integer = 0
            While (i < FreqLevelSeries.Points.Count() Or i < LevelSweepSeries.Points.Count() Or i < SNRPercentageSeries.Points.Count)
                If (i < LevelSweepSeries.Points.Count()) Then
                    Dim Point As DataPoint = LevelSweepSeries.Points.Item(i)
                    TextFile.Write(CStr(Point.AxisLabel) & separator & CStr(Point.YValues.GetValue(0)) & separator)
                Else
                    TextFile.Write(separator & separator)
                End If

                If (i < SNRPercentageSeries.Points.Count()) Then
                    Dim Point As DataPoint = SNRPercentageSeries.Points.Item(i)
                    TextFile.Write(CStr(Point.AxisLabel) & separator & CStr(Point.YValues.GetValue(0)) & separator)
                Else
                    TextFile.Write(separator & separator)
                End If

                If (i < FreqLevelSeries.Points.Count()) Then
                    Dim Point1 As DataPoint = FreqLevelSeries.Points.Item(i)
                    Dim Point2 As DataPoint = FreqDistortionSeries.Points.Item(i)

                    TextFile.Write(CStr(Point1.AxisLabel) & separator & CStr(Point1.YValues.GetValue(0)) & separator)
                    TextFile.Write(CStr(Point2.AxisLabel) & separator & CStr(Point2.YValues.GetValue(0)))
                End If
                TextFile.WriteLine()
                TextFile.Flush()
                i += 1
            End While
            TextFile.Close()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' ========================================================================
        ' INITIALIZATION SECTION
        ' ========================================================================
        'buffer = New String("x", 1024)      ' Size of read buffer

        'Read settings from object saved last time
        If File.Exists("settings.bin") Then
            Dim SettingsFileStream As Stream = File.OpenRead("settings.bin")
            Dim deserializer As New BinaryFormatter()
            Dim Data As PersistentData
            Data = CType(deserializer.Deserialize(SettingsFileStream), PersistentData)
            SettingsFileStream.Close()


            chkLevelSweepActive.Checked = Data.LevelSweep
            txtStartLevelV.Text = Data.SweepStartLevel
            txtStartLevelDB.Text = BoontonInterface.VoltsToDB(CDbl(Data.SweepStartLevel))
            txtLevelOhmLoad.Text = Data.SweepOhmLoad
            txtLevelDistortionThreshDB.Text = Data.DistortionThresh

            chkSNRActive.Checked = Data.SNR
            txtSNRReferenceV.Text = Data.SNRRef
            txtSNRReferenceDB.Text = BoontonInterface.VoltsToDB(CDbl(Data.SNRRef))

            chkFreqDistActive.Checked = Data.FreqAndDist
            txtFreqSourceV.Text = Data.FreqSource
            txtFreqSourceDB.Text = BoontonInterface.VoltsToDB(CDbl(Data.FreqSource))
            chkFullBandwidth.Checked = Data.FullRange
            chkInverseRIAA.Checked = Data.InverseRIAA
            chk318uS.Checked = Data.use318uS

            txtDescription.Text = Data.Description
            tbGPIBAddress.Text = Data.GPIBAddress
            tbRemoteAddress.Text = Data.RemoteAddress
        End If

        ' Set up chart areas and Series

        FreqLevelSeries.XValueType = ChartValueType.String
        FreqLevelSeries.ChartArea = "Level"
        FreqLevelSeries.Name = "Level"
        FreqLevelSeries.ChartType = SeriesChartType.Line
        FreqLevelSeries.XValueType = ChartValueType.String

        FreqLevelArea.Name = "Level"
        FreqLevelArea.AxisX.Interval = 1
        FreqLevelArea.AxisX.IsLabelAutoFit = True
        FreqLevelArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.DecreaseFont
        FreqLevelArea.AxisX.Title = "Frequency (Hz)"
        FreqLevelArea.AxisY.Title = "Level (dBV)"
        FreqLevelArea.AlignmentOrientation = AreaAlignmentOrientations.Horizontal
        FreqLevelArea.AxisX.MajorGrid.LineColor = Color.LightGray
        FreqLevelArea.AxisX.LineColor = Color.LightGray
        FreqLevelArea.AxisY.LineColor = Color.LightGray
        FreqLevelArea.AxisY.MajorGrid.LineColor = Color.LightGray



        FreqDistortionSeries.XValueType = ChartValueType.String
        FreqDistortionSeries.ChartArea = "Distortion"
        FreqDistortionSeries.Name = "Distortion"
        FreqDistortionSeries.ChartType = SeriesChartType.Line
        FreqDistortionSeries.XValueType = ChartValueType.String

        FreqDistortionArea.Name = "Distortion"
        FreqDistortionArea.AxisX.Interval = 1
        FreqDistortionArea.AxisX.IsLabelAutoFit = True
        FreqDistortionArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.DecreaseFont
        FreqDistortionArea.AxisX.Title = "Frequency (Hz)"
        FreqDistortionArea.AxisY.Title = "Distortion (dB)"
        FreqDistortionArea.AlignmentOrientation = AreaAlignmentOrientations.Horizontal
        FreqDistortionArea.AxisX.MajorGrid.LineColor = Color.LightGray
        FreqDistortionArea.AxisX.LineColor = Color.LightGray
        FreqDistortionArea.AxisY.LineColor = Color.LightGray
        FreqDistortionArea.AxisY.MajorGrid.LineColor = Color.LightGray

        LevelSweepSeries.XValueType = ChartValueType.String
        LevelSweepSeries.ChartArea = "Sweep"
        LevelSweepSeries.Name = "Sweep"
        LevelSweepSeries.ChartType = SeriesChartType.Line
        LevelSweepSeries.XValueType = ChartValueType.String

        LevelSweepArea.Name = "Sweep"
        LevelSweepArea.AxisX.Interval = 1
        LevelSweepArea.AxisX.IsLabelAutoFit = True
        LevelSweepArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.DecreaseFont
        LevelSweepArea.AxisX.Title = "Level (dB)"
        LevelSweepArea.AxisY.Title = "Distortion (dB)"
        LevelSweepArea.AlignmentOrientation = AreaAlignmentOrientations.Horizontal
        LevelSweepArea.AxisX.MajorGrid.LineColor = Color.LightGray
        LevelSweepArea.AxisX.LineColor = Color.LightGray
        LevelSweepArea.AxisY.LineColor = Color.LightGray
        LevelSweepArea.AxisY.MajorGrid.LineColor = Color.LightGray



        SNRPercentageSeries.XValueType = ChartValueType.String
        SNRPercentageSeries.ChartArea = "SNR"
        SNRPercentageSeries.Name = "SNR"
        SNRPercentageSeries.ChartType = SeriesChartType.Line
        SNRPercentageSeries.XValueType = ChartValueType.String

        SNRPercentageArea.Name = "SNR"
        SNRPercentageArea.AxisX.Interval = 1
        SNRPercentageArea.AxisX.IsLabelAutoFit = True
        SNRPercentageArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.DecreaseFont
        SNRPercentageArea.AxisX.Title = "Frequency (Hz)"
        SNRPercentageArea.AxisY.Title = "Distortion (dB)"
        SNRPercentageArea.AlignmentOrientation = AreaAlignmentOrientations.Horizontal
        SNRPercentageArea.AxisX.MajorGrid.LineColor = Color.LightGray
        SNRPercentageArea.AxisX.LineColor = Color.LightGray
        SNRPercentageArea.AxisY.LineColor = Color.LightGray
        SNRPercentageArea.AxisY.MajorGrid.LineColor = Color.LightGray
    End Sub

    ' Records settings onto object to be saved onto a file
    Private Sub Form1_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim Data As New PersistentData

        Data.LevelSweep = chkLevelSweepActive.Checked
        Data.SweepStartLevel = txtStartLevelV.Text
        Data.SweepOhmLoad = txtLevelOhmLoad.Text
        Data.DistortionThresh = txtLevelDistortionThreshDB.Text

        Data.SNR = chkSNRActive.Checked
        Data.SNRRef = txtSNRReferenceV.Text

        Data.FreqAndDist = chkFreqDistActive.Checked
        Data.FreqSource = txtFreqSourceV.Text
        Data.FullRange = chkFullBandwidth.Checked
        Data.InverseRIAA = chkInverseRIAA.Checked
        Data.use318uS = chk318uS.Checked

        Data.Description = txtDescription.Text
        Data.GPIBAddress = tbGPIBAddress.Text
        Data.RemoteAddress = tbRemoteAddress.Text


        Dim SettingsFile As Stream = File.Create("settings.bin")
        Dim serializer As New BinaryFormatter()
        serializer.Serialize(SettingsFile, Data)
        SettingsFile.Close()
    End Sub


    Private Sub EndTestButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndTestButton.Click
        CleanupFlag = 1
    End Sub

    Private Sub QuitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitButton.Click
        Me.Close()
        Application.Exit()
    End Sub


    '    Private Sub FindWattButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '        Dim BDINDEX As Byte = Val(tbGPIBAddress.Text)           'GPIB Board Address
    '        Dim PRIMARY_ADDR As Byte = Val(tbRemoteAddress.Text)    'Remote Instrument Address

    '        'Dim Start, Finish As Byte
    '        Dim rdbuf As String
    '        'Dim Sweep As Byte
    '        Dim WattPower As Double
    '        'Dim dB As Double
    '        'Dim SNR As Double
    '        Dim OutputVolts As Double
    '        Dim dblFindWatt As Double
    '        Dim InputVolts As Double
    '        'Dim Distortion As Double

    '        'RUN, QUIT buttons disabled & END TEST button enabled during the test
    '        RunButton.Enabled = False
    '        QuitButton.Enabled = False
    '        EndTestButton.Enabled = True


    '        'Prevents crash if garbage buffer data is retrieve from Boonton

    '        'This brings the Boonton online using ildev. A device handle,
    '        'Boonton, is returned and is used in all subsequent calls to the device.
    '        'Boonton = li.ibdev(BDINDEX, PRIMARY_ADDR, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
    '        'If (li.ibsta And c.ERR) Then Call GPIBCleanup("Unable to open device")


    '        ' ========================================================================
    '        '
    '        '  MAIN BODY SECTION
    '        '
    '        '  In this application, the Main Body communicates with the instrument
    '        '  by writing a command to it and reading its response. 
    '        ' ========================================================================




    '        ' set source frequency and give a 500ms settling pause
    '        SetFrequency(1000)
    '        System.Threading.Thread.Sleep((500))
    '        OutputVolts = 0.02          'CDbl(tbOutputVolts.Text)
    '        dblFindWatt = CDbl(tbFindWatt.Text)

    '        'main loop
    '        Do Until WattPower > (dblFindWatt + dblFindWatt * 0.005)

    '            'sets output voltage at either normal level
    '            SetSource((OutputVolts))

    '            'sets up 'LEVEL' mode, waits one second for settling and then reads Boonton buffer
    '            MeasureVRMS()
    '            System.Threading.Thread.Sleep((100))
    '            rdbuf = Space(20)
    '            'li.ilrd(Boonton, rdbuf, Len(rdbuf))
    '            'If (li.ibsta And c.ERR) Then Call GPIBCleanup("Unable to read from device")

    '            'take left seven digits from the buffer, ie trim garbage to prevent math error
    '            rdbuf = Strings.Left(rdbuf, 7)

    '            'rounds incoming input VRMS to the millivolt
    '            InputVolts = System.Math.Round(Val(rdbuf), 3)

    '            WattPower = System.Math.Round(InputVolts * InputVolts / Val(txtLevelOhmLoad.Text), 3)

    '            If WattPower < dblFindWatt Then
    '                If dblFindWatt / WattPower > 3.3 Then
    '                    OutputVolts = OutputVolts + 0.03
    '                ElseIf dblFindWatt / WattPower > 2.1 Then
    '                    OutputVolts = OutputVolts + 0.02
    '                ElseIf dblFindWatt / WattPower > 1.5 Then
    '                    OutputVolts = OutputVolts + 0.01
    '                ElseIf dblFindWatt / WattPower > 1.15 Then
    '                    OutputVolts = OutputVolts + 0.01
    '                    'ElseIf dblFindWatt / WattPower > 1.02 Then
    '                    '    OutputVolts = OutputVolts + 0.005
    '                ElseIf dblFindWatt / WattPower > 1 Then
    '                    OutputVolts = OutputVolts + 0.001
    '                End If


    '            End If
    '            If WattPower > dblFindWatt Then
    '                OutputVolts = OutputVolts + 0.001
    '            End If

    '            'this allows other events, such as a hard escape
    '            System.Windows.Forms.Application.DoEvents()
    '            If CleanupFlag = 1 Then GoTo cleanup
    '        Loop




    '        'this allows other events, such as a hard escape
    '        System.Windows.Forms.Application.DoEvents()
    '        If CleanupFlag = 1 Then GoTo cleanup

    '        txtStartLevelDB.Text = System.Math.Round(OutputVolts, 3)

    '        ' ========================================================================
    '        '
    '        ' CLEANUP SECTION
    '        '
    '        ' ========================================================================
    'Cleanup:



    '        ' Take the device offline and make sure there's no output
    '        SetZero()
    '        'li.ilonl(Boonton, 0)
    '        'If (li.ibsta And c.ERR) Then Call GPIBCleanup("Error putting device offline.")

    '        'Enable user inputs.
    '        RunButton.Enabled = True
    '        QuitButton.Enabled = True
    '        EndTestButton.Enabled = False
    '        MsgBox("Done!")

    '        'resets cleanup flag
    '        CleanupFlag = 0


    '    End Sub

    ' This class stores form data to put back in on restartup.
    <Serializable()> Private Class PersistentData
        Public LevelSweep As Boolean
        Public SweepStartLevel As String
        Public SweepOhmLoad As String
        Public DistortionThresh As String

        Public SNR As Boolean
        Public SNRRef As String

        Public FreqAndDist As Boolean
        Public FreqSource As String
        Public FullRange As Boolean
        Public InverseRIAA As Boolean
        Public use318uS As Boolean

        Public Description As String
        Public GPIBAddress As String
        Public RemoteAddress As String
    End Class

    Private Sub txtDescription_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDescription.TextChanged
        If (Chart1.Titles.Count <> 0) Then
            Chart1.Titles(0).Text = txtDescription.Text
        End If

    End Sub

    'If change was from DB, sets V to proper value, and the other way around.
    Private Sub DBVConverter(sender As Object, e As EventArgs) Handles txtStartLevelV.TextChanged, txtStartLevelDB.TextChanged, txtSNRReferenceV.TextChanged, txtSNRReferenceDB.TextChanged, txtFreqSourceV.TextChanged, txtFreqSourceDB.TextChanged, txtLevelDistortionThreshV.TextChanged, txtLevelDistortionThreshDB.TextChanged, txtSourceMaxV.TextChanged, txtSourceMaxDB.TextChanged
        Dim textBox As TextBox = CType(sender, TextBox)
        If (Not textBox.Focused) Then Return
        Dim name As String = textBox.Name.ToString
        'The last character of the name is either V or B
        Dim dimention As Char = name.Substring(name.Length - 1)
        If (dimention = "V"c) Then
            Dim dbName As String = name.Substring(0, name.Length - 1) & "DB"
            Dim dbText As Control = Me.Controls.Item(dbName)
            If (IsNothing(dbText)) Then Return

            Dim vVal As Double
            If (Not Double.TryParse(textBox.Text, vVal)) Then
                dbText.Text = "NaN"
                Return
            End If
            dbText.Text = BoontonInterface.VoltsToDB(vVal)

        Else
            Dim vName As String = name.Substring(0, name.Length - 2) & "V"
            Dim vText As Control = Me.Controls.Item(vName)
            If (IsNothing(vText)) Then Return

            Dim dbVal As Double
            If (Not Double.TryParse(textBox.Text, dbVal)) Then
                vText.Text = "NaN"
                Return
            End If
            vText.Text = BoontonInterface.DBToVolts(CDbl(dbVal))
        End If
    End Sub
End Class