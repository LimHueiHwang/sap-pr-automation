
Option Explicit

Sub Close_Fixed_PR()

    ' ============================================================
    ' SAP GUI objects
    ' ============================================================
    Dim SapGuiApp As Object
    Dim SapConnection As Object
    Dim SapSession As Object

    ' ============================================================
    ' Excel objects
    ' ============================================================
    Dim wsCloseFixedPR As Worksheet
    Dim currentPRCell As Range

    ' ============================================================
    ' Variables
    ' ============================================================
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim sapMessage As String

    On Error GoTo ErrorHandler

    ' ============================================================
    ' Connect to SAP
    ' ============================================================
    On Error Resume Next
    Set SapGuiApp = GetObject("SAPGUI").GetScriptingEngine
    On Error GoTo ErrorHandler

    If SapGuiApp Is Nothing Then

        MsgBox _
            "Please log in to SAP before running this script.", _
            vbCritical, _
            "SAP Connection"

        Exit Sub

    End If

    If SapGuiApp.Children.Count = 0 Then

        MsgBox _
            "No SAP connection was found.", _
            vbCritical, _
            "SAP Connection"

        Exit Sub

    End If

    Set SapConnection = SapGuiApp.Children(0)

    If SapConnection.Children.Count = 0 Then

        MsgBox _
            "No active SAP session was found.", _
            vbCritical, _
            "SAP Session"

        Exit Sub

    End If

    Set SapSession = SapConnection.Children(0)

    ' ============================================================
    ' Reference worksheet
    ' ============================================================
    Set wsCloseFixedPR = _
        ThisWorkbook.Sheets("Close Fixed PR")

    Application.DisplayAlerts = False

    ' ============================================================
    ' Determine last row
    ' ============================================================
    lastRow = wsCloseFixedPR.Cells( _
                  wsCloseFixedPR.Rows.Count, _
                  "A" _
              ).End(xlUp).Row

    If lastRow < 4 Then

        MsgBox _
            "No Purchase Requisition numbers were found.", _
            vbExclamation, _
            "Close/Fix PR"

        GoTo CleanExit

    End If

    ' ============================================================
    ' Clear previous results
    ' ============================================================
    wsCloseFixedPR.Range( _
        "B4:C" & lastRow _
    ).ClearContents

    ' ============================================================
    ' Start ME52N
    ' ============================================================
    With SapSession

        .findById("wnd[0]").maximize

        .findById( _
            "wnd[0]/tbar[0]/okcd" _
        ).Text = "/nme52n"

        .findById("wnd[0]").sendVKey 0

    End With

    ' ============================================================
    ' Process each PR
    ' ============================================================
    For rowIndex = 4 To lastRow

        Set currentPRCell = _
            wsCloseFixedPR.Range("A" & rowIndex)

        ' --------------------------------------------------------
        ' Select Other Purchase Requisition
        ' --------------------------------------------------------
        SapSession.findById( _
            "wnd[0]/tbar[1]/btn[17]" _
        ).press

        ' --------------------------------------------------------
        ' Enter PR number
        ' --------------------------------------------------------
        With SapSession.findById( _
            "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/" & _
            "ctxtMEPO_SELECT-BANFN" _
        )

            .Text = currentPRCell.Value
            .caretPosition = Len(.Text)

        End With

        SapSession.findById("wnd[1]").sendVKey 0

        ' --------------------------------------------------------
        ' Read SAP status message
        ' --------------------------------------------------------
        sapMessage = SapSession.findById( _
            "wnd[0]/sbar/pane[0]" _
        ).Text

        ' --------------------------------------------------------
        ' Process existing PR
        ' --------------------------------------------------------
        If Not sapMessage Like "*does not exist*" Then

            With SapSession

                .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT5/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/chkMEREQ3321-FIXKZ").Selected = False
                .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT5/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/chkMEREQ3321-EBAKZ").Selected = True

                ' Save
                .findById( _
                    "wnd[0]/tbar[0]/btn[11]" _
                ).press

            End With

        End If

        ' --------------------------------------------------------
        ' Capture final SAP message
        ' --------------------------------------------------------
        sapMessage = SapSession.findById( _
            "wnd[0]/sbar/pane[0]" _
        ).Text

        ' --------------------------------------------------------
        ' Handle result
        ' --------------------------------------------------------
        If sapMessage = "" Then

            On Error Resume Next
            SapSession.findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).press
            On Error GoTo ErrorHandler

            wsCloseFixedPR.Range( _
                "B" & rowIndex _
            ).Value = "No changes made"

        Else

            wsCloseFixedPR.Range( _
                "B" & rowIndex _
            ).Value = sapMessage

        End If

    Next rowIndex

    MsgBox _
        "Process completed. Please check column B for results.", _
        vbInformation, _
        "Close/Fix PR"

CleanExit:

    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True

    MsgBox _
        "Close/Fix PR stopped." & _
        vbCrLf & vbCrLf & _
        "Row: " & rowIndex & _
        vbCrLf & _
        "PR: " & IIf(currentPRCell Is Nothing, "", currentPRCell.Value) & _
        vbCrLf & _
        "Error: " & Err.Description, _
        vbCritical, _
        "Close/Fix PR Error"

End Sub

