Option Explicit

Sub validate_PR()

    ' ============================================================
    ' SAP GUI objects
    ' ============================================================
    Dim SapGuiApp As Object
    Dim SapConnection As Object
    Dim SapSession As Object

    ' ============================================================
    ' Excel objects
    ' ============================================================
    Dim CombinePR As Worksheet
    Dim Plan_Order As Worksheet
    Dim wbSource As Workbook
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim obj As Object

    ' ============================================================
    ' Variables
    ' ============================================================
    Dim totalRows As Long
    Dim lastRow As Long
    Dim i As Long
    Dim rowIndex As Long

    Dim nPlant As String
    Dim nPart As Variant

    Dim fieldID As String
    Dim wbNamePattern As String

    Dim startTime As Double
    Dim found As Boolean

    On Error GoTo ErrorHandler

    ' ============================================================
    ' Connect to SAP
    ' ============================================================
    On Error Resume Next
    Set SapGuiApp = GetObject("SAPGUI").GetScriptingEngine
    On Error GoTo ErrorHandler

    If SapGuiApp Is Nothing Then

        MsgBox _
            "Please log on to the SAP system before proceeding.", _
            vbCritical, _
            "SAP Connection"

        Exit Sub

    End If

    If SapGuiApp.Children.Count = 0 Then

        MsgBox _
            "Please log on to the SAP system before proceeding.", _
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
    ' Reference worksheets
    ' ============================================================
    Set CombinePR = ThisWorkbook.Sheets("Combine PR")
    Set Plan_Order = ThisWorkbook.Sheets("Plan Order")

    ' ============================================================
    ' Determine total material rows
    ' ============================================================
    totalRows = CombinePR.Cells( _
                    CombinePR.Rows.Count, _
                    "A" _
                ).End(xlUp).Row

    If totalRows < 4 Then

        MsgBox _
            "No material data was found in the Combine PR sheet.", _
            vbExclamation, _
            "Plan Order Validation"

        Exit Sub

    End If

    ' ============================================================
    ' Determine result range
    ' ============================================================
    lastRow = CombinePR.Cells( _
                  CombinePR.Rows.Count, _
                  "G" _
              ).End(xlUp).Row

    If lastRow < 4 Then
        lastRow = totalRows
    End If

    ' ============================================================
    ' Disable Excel alerts during automation
    ' ============================================================
    Application.DisplayAlerts = False

    ' ============================================================
    ' Clear previous results
    ' ============================================================
    CombinePR.Range( _
        "C4:G" & lastRow _
    ).ClearContents

    ' ============================================================
    ' Setup validation formulas
    ' ============================================================
    For i = 4 To totalRows

        ' Extract PR number from SAP result
        CombinePR.Range("E" & i).Formula = _
            "=TEXTBEFORE(TEXTAFTER(C" & i & _
            ",""purchase requisition ""), "" "")"

        ' Look up available Plan Order quantity
        CombinePR.Range("F" & i).Formula = _
            "=IFERROR(XLOOKUP(A" & i & _
            ",Summary!A:A,Summary!B:B),0)"

        ' Determine validation result
        CombinePR.Range("G" & i).Formula = _
            "=IF(F" & i & ">=B" & i & ",""Ok""," & _
            "IF(F" & i & "=0,""Check if SA part"",""Not enough Plan Order""))"

    Next i

    ' ============================================================
    ' Validate plant
    ' ============================================================
    nPlant = Trim(CStr(CombinePR.Range("B2").Value))

    If nPlant = "" Then

        MsgBox _
            "Please fill in the Plant Code.", _
            vbExclamation, _
            "Missing Plant Code"

        GoTo CleanExit

    End If

    ' ============================================================
    ' Clear previous Plan Order results
    ' ============================================================
    Plan_Order.Cells.Clear

    ' ============================================================
    ' Run SAP Query
    ' ============================================================
    With SapSession

        .findById("wnd[0]").maximize

        .findById( _
            "wnd[0]/tbar[0]/okcd" _
        ).Text = "/nsq00"

        .findById("wnd[0]").sendVKey 0

        .findById( _
            "wnd[0]/tbar[1]/btn[19]" _
        ).press

        .findById( _
            "wnd[1]/usr/cntlGRID1/shellcont/shell" _
        ).currentCellRow = 8

        .findById( _
            "wnd[1]/usr/cntlGRID1/shellcont/shell" _
        ).selectedRows = "8"

        .findById( _
            "wnd[1]/usr/cntlGRID1/shellcont/shell" _
        ).doubleClickCurrentCell

        .findById( _
            "wnd[0]/usr/ctxtRS38R-QNUM" _
        ).Text = "DEMO_SAP_QUERY"

        .findById( _
            "wnd[0]/tbar[1]/btn[8]" _
        ).press

        .findById( _
            "wnd[0]/usr/ctxtLANGUAGE-LOW" _
        ).Text = "EN"

        .findById( _
            "wnd[0]/usr/ctxtPLANT-LOW" _
        ).Text = nPlant

        ' --------------------------------------------------------
        ' Clear existing selection values
        ' --------------------------------------------------------
        .findById( _
            "wnd[0]/usr/btn%_SP$00019_%_APP_%-VALU_PUSH" _
        ).press

        .findById("wnd[1]/tbar[0]/btn[16]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press

        .findById( _
            "wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH" _
        ).press

        .findById("wnd[1]/tbar[0]/btn[16]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press

        ' --------------------------------------------------------
        ' Configure Excel output
        ' --------------------------------------------------------
        .findById( _
            "wnd[0]/usr/rad%EXCEL" _
        ).Select

        .findById( _
            "wnd[0]/usr/ctxtSP$00003-LOW" _
        ).Text = "PA"

        .findById( _
            "wnd[0]/usr/ctxtMATERIAL-LOW" _
        ).Text = "1"

        .findById( _
            "wnd[0]/usr/btn%_MATERIAL_%_APP_%-VALU_PUSH" _
        ).press

    End With

    ' ============================================================
    ' Prepare material list
    ' ============================================================
    If totalRows = 4 Then

        ReDim nPart(1 To 1, 1 To 1)

        nPart(1, 1) = CombinePR.Range("A4").Value

    Else

        nPart = CombinePR.Range( _
                    "A4:A" & totalRows _
                ).Value

    End If

    ' ============================================================
    ' Enter materials into SAP selection
    ' ============================================================
    For i = 1 To UBound(nPart, 1)

        If i <= 7 Then

            rowIndex = i - 1

        Else

            rowIndex = ((i - 1) Mod 7) + 1

        End If

        fieldID = _
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/" & _
            "ssubSCREEN_HEADER:SAPLALDB:3010/" & _
            "tblSAPLALDBSINGLE/" & _
            "ctxtRSCSEL_255-SLOW_I[1," & rowIndex & "]"

        ' --------------------------------------------------------
        ' Scroll SAP material table when required
        ' --------------------------------------------------------
        If rowIndex = 1 And i > 7 Then

            SapSession.findById( _
                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/" & _
                "ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                "tblSAPLALDBSINGLE" _
            ).VerticalScrollbar.Position = i - 1

        End If

        ' --------------------------------------------------------
        ' Locate material input field
        ' --------------------------------------------------------
        Set obj = Nothing
        Err.Clear

        On Error Resume Next
        Set obj = SapSession.findById(fieldID)
        On Error GoTo ErrorHandler

        If obj Is Nothing Then

            MsgBox _
                "Unable to find SAP material input field:" & _
                vbCrLf & vbCrLf & _
                fieldID, _
                vbCritical, _
                "SAP Control Error"

            GoTo CleanExit

        End If

        obj.Text = nPart(i, 1)

    Next i

    ' ============================================================
    ' Execute SAP query
    ' ============================================================
    SapSession.findById( _
        "wnd[1]/tbar[0]/btn[8]" _
    ).press

    SapSession.findById( _
        "wnd[0]/tbar[1]/btn[8]" _
    ).press

    ' ============================================================
    ' Select Excel export format
    ' ============================================================
    SapSession.findById( _
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/" & _
        "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]" _
    ).Select

    SapSession.findById( _
        "wnd[1]/tbar[0]/btn[0]" _
    ).press

    SapSession.findById( _
        "wnd[1]/tbar[0]/btn[0]" _
    ).press

    ' ============================================================
    ' Wait for exported Excel workbook
    ' ============================================================
    wbNamePattern = "Worksheet in Basis*"
    startTime = Timer
    found = False

    Do While Not found And Timer - startTime < 60

        DoEvents

        For Each wb In Application.Workbooks

            If wb.Name Like wbNamePattern Then

                Set wbSource = wb
                found = True

                Exit For

            End If

        Next wb

    Loop

    If Not found Then

        MsgBox _
            "Workbook was not opened within 60 seconds.", _
            vbExclamation, _
            "SAP Export Timeout"

        GoTo CleanExit

    End If

    ' ============================================================
    ' Copy SAP export to Plan Order sheet
    ' ============================================================
    Set wsSource = wbSource.Sheets(1)
    Set wsTarget = Plan_Order

    wsTarget.Cells.Clear

    wsSource.UsedRange.Copy _
        Destination:=wsTarget.Range("A1")

    ' ============================================================
    ' Format Plan Order quantity column
    ' ============================================================
    With wsTarget.Columns("C")

        .NumberFormat = "0.00"
        .Value = .Value

    End With

    MsgBox _
        "Validate_PR successfully!", _
        vbInformation, _
        "Done"

CleanExit:

    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True

    MsgBox _
        "Plan Order validation stopped." & _
        vbCrLf & vbCrLf & _
        "Error: " & Err.Description, _
        vbCritical, _
        "Validation Error"

End Sub

