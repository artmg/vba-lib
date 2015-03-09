VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mod_prj_PlanInACan_Virt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' Plan in a can
'
' 080229.AMG added basic documentation
' 080228.AMG minor logic changes
' 080228.AMG created
'
' This module is designed to go in a Microsoft Excel 2003 workbook
' that contains a Named Range as per the constant cRange below
' which has a series of columns as per the enum SvrCol below
' (NB: leave the row containing column headers OUTSIDE the Named Range)
'
' Execute this Macro on a PC with Microsoft Project 2003 installed
' and it will automatically create a project based upon the contents
' of the table
'

' Tech note: if you export this module, when you re-import it elsewhere
' use Tools/References to include the following:
'        Microsoft Project 11.0 Object Library (C:\Program Files\Microsoft Office\OFFICE11\MSPRJ.OLB)
'        Microsoft Excel 11.0 Object Library   (C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE)


Const cCalendar As String = "24 Hours"
Const cRange As String = "ServerNames"
'Const cRange As String = "TestServerNames"

Enum SvrCol
    ServerName = 1
    FateName
    When
    WhenNotes
    FateText
    BusinessConstraints
    PerforMeasured
    Site
    HostOwner
    Comments
    MainAppOwner
    OSFromServerDB
    DiskUsedFromServerDB
    ServerIsMoving
    RebootedRecently
    ScheduleABackup
    ExecuteScripts
    WriteScripts
    ValidateScripts
End Enum

Dim prj As Project
Dim svrRow As Range


Sub CreatePlanFromServerList()
    
    Set pj = CreateObject("MSProject.Project")
    Set prj = pj.Application.Projects.Add
    
    prj.Application.SidepaneToggle (False)
'    prj.Application.Calculation = pjManual
    prj.DefaultDurationUnits = pjHour
    ProjectSummaryInfo Calendar:=cCalendar
    OptionsView DateFormat:=8

    With prj.Tasks.Add("Server Virtualisation - Plan in a can")
        .OutlineLevel = "1"
    End With

    'server
    For Each svrRow In Range(cRange).Rows
        With prj.Tasks.Add
            .Name = DataFrom(SvrCol.ServerName)
            .OutlineLevel = "2"
            .ResourceNames = DataFrom(SvrCol.HostOwner)
        End With

' CONVERSION
        
        Dim ConversionID As Integer
        With prj.Tasks.Add("Conversion")
            .OutlineLevel = "3"
            ConversionID = .id
        End With

        Dim LastID As Long
        With prj.Tasks.Add("Suspend Application")
            .Start = DataFrom(SvrCol.When) & " 10:00"
            .Duration = 60
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.MainAppOwner)
            .OutlineLevel = "4"
        End With

        With prj.Tasks.Add("Disable Server")
            .Duration = 60
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.HostOwner)
            FollowLast .id
        End With
        
        With prj.Tasks.Add("Pre-migration checks")
            .Duration = 60
            .Estimated = False
            .ResourceNames = "IP Soft"
            FollowLast .id
        End With
        
        With prj.Tasks.Add("P2V")
            If DataFrom(DiskUsedFromServerDB) = "" Then
                .Duration = 180
                .Estimated = True
            Else
                .Duration = DataFrom(DiskUsedFromServerDB) * 2.25
                .Estimated = False
            End If
            .ResourceNames = "P2V Resource"
            FollowLast .id
        End With
        
        With prj.Tasks.Add("post-migration checks")
            .Duration = 60
            .Estimated = False
            .ResourceNames = "IP Soft"
            FollowLast .id
        End With

        With prj.Tasks.Add("Test Application")
            .Duration = DataFrom(ExecuteScripts) * 60
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.MainAppOwner) & "'s Testers"
            FollowLast .id
        End With

        With prj.Tasks.Add("Resolve issues")
            .Duration = 60
            .Estimated = True
            If DataFrom(SvrCol.HostOwner) <> DataFrom(MainAppOwner) Then
                .ResourceNames = DataFrom(SvrCol.HostOwner) & ", " & DataFrom(MainAppOwner)
            Else
                .ResourceNames = DataFrom(SvrCol.HostOwner)
            End If
            FollowLast .id
        End With

        If LCase(DataFrom(SvrCol.Site)) <> "FS" Then
            With prj.Tasks.Add("Refresh ARP")
                .Duration = 10
                .Estimated = False
                .ResourceNames = "OpsTelecom"
                FollowLast .id
            End With
        End If

        With prj.Tasks.Add("Return to Service")
            .Duration = 60
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.HostOwner)
            FollowLast .id
        End With

' PREREQS
        With prj.Tasks.Add("Prerequisites")
            .OutlineLevel = "3"
        End With

        If LCase(DataFrom(SvrCol.PerforMeasured)) <> "yes" Then
            With prj.Tasks.Add("Start PerfMon")
                .Duration = 30
                .Estimated = False
                .ResourceNames = DataFrom(SvrCol.HostOwner)
                .OutlineLevel = "4"
                .Predecessors = ConversionID & "SF-" & 24 * 7 * 2
            End With
        End If
        
        With prj.Tasks.Add("Clean cold start")
            .Duration = 30
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.HostOwner)
            .OutlineLevel = "4"
            .Predecessors = ConversionID & "SF-" & 24 * 7
        End With

        If LCase(DataFrom(SvrCol.ScheduleABackup)) <> "no" Then
            With prj.Tasks.Add("Back server up")
                .Duration = 120
                .Estimated = False
                .ResourceNames = DataFrom(SvrCol.HostOwner)
                .OutlineLevel = "4"
                .Predecessors = ConversionID & "SF"
            End With
        End If

        With prj.Tasks.Add("Write test and rollback scripts")
            .Duration = DataFrom(WriteScripts) * 60
            .Estimated = False
            .ResourceNames = DataFrom(SvrCol.HostOwner) & "'s Testers"
            .OutlineLevel = "4"
            .Predecessors = ConversionID & "SF-" & 24 * 7 * 2
        End With

        With prj.Tasks.Add("validate test and rollback scripts")
            .Duration = DataFrom(ValidateScripts) * 60
            .Estimated = False
            .ResourceNames = "My Resource Name"
            FollowLast .id
        End With

    Next

' tidy up
    
    Dim res As Resource
    For Each res In prj.Resources
        res.BaseCalendar = cCalendar
    Next
'    prj.Application.Calculation = pjAutomatic

    ColumnBestFit Column:=5
    ColumnBestFit Column:=6
    ColumnBestFit Column:=3

End Sub

Function DataFrom(cl As SvrCol)
    DataFrom = svrRow.Cells(1, cl).Value
End Function
'Function DataFrom(rw As Range, cl As SvrCol)
'    DataFrom = rw.Cells(1, cl).Value
'End Function

Sub FollowLast(id)
    With prj
        .Tasks(id).Predecessors = id - 1
        .Tasks(id).OutlineLevel = .Tasks(id - 1).OutlineLevel
    End With
End Sub

