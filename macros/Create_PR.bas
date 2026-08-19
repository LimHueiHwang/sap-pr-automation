Option Explicit

Sub Create_PR()

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

    ' ============================================================
    ' Variables
    ' ============================================================
    Dim totalRows As Long
    Dim rowComPR As Long
    Dim i As Long

    Dim nPlant As String
    Dim DRemoveVendor As String

    Dim nMaterial As Range
    Dim nQuantity As Range

    Dim MRPElement As String
    Dim recordQty As Long

    Dim tableViewSize As Long
    Dim nR As Long
    Dim planOrderRow As Long
    Dim foundPlanOrder As Boolean

    Dim Result As String

    On Error GoTo ErrorHandler

    ' ============================================================
    ' Connect to SAP
    ' ============================================================
    On Error Resume Next
    Set SapGuiApp = GetObject("SAPGUI").GetScriptingEngine
    On Error GoTo ErrorHandler

    If SapGuiApp Is Nothing Then

        MsgBox _
            "Please logon to SAP system before proceeding.", _
            vbCritical, _
            "SAP Connection"

        Exit Sub

    End If

    If SapGuiApp.Children.Count = 0 Then

        MsgBox _
            "Please logon to SAP system before proceeding.", _
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
    ' Excel worksheet
    ' ============================================================
    Set CombinePR = ThisWorkbook.Sheets("Combine PR")

    ' ============================================================
    ' Disable Excel alerts during automation
    ' ============================================================
    Application.DisplayAlerts = False

    ' ============================================================
    ' Get total rows
    ' ============================================================
    totalRows = CombinePR.Cells( _
                    CombinePR.Rows.Count, _
                    "A" _
                ).End(xlUp).Row

    ' ============================================================
    ' Clear previous results
    ' ============================================================
    If totalRows >= 4 Then

        CombinePR.Range( _
            "C4:E" & totalRows _
        ).ClearContents

    End If

    ' ============================================================
    ' Extract PR number from SAP result
    ' ============================================================
    For i = 4 To totalRows

        CombinePR.Range("E" & i).Formula = _
            "=TEXTBEFORE(TEXTAFTER(C" & i & _
            ",""purchase requisition ""), "" "")"

    Next i

    ' ============================================================
    ' Read configuration
    ' ============================================================
    nPlant = Trim(CStr(CombinePR.Range("B2").Value))
    DRemoveVendor = Trim(CStr(CombinePR.Range("H1").Value))

    ' ============================================================
    ' Validate plant
    ' ============================================================
    If nPlant = "" Then

        MsgBox _
            "Please fill in the Plant Code.", _
            vbExclamation, _
            "Missing Plant Code"

        GoTo CleanExit

    End If

    ' ============================================================
    ' Process each material
    ' ============================================================
    For rowComPR = 4 To totalRows

        Set nMaterial = CombinePR.Range("A" & rowComPR)
        Set nQuantity = CombinePR.Range("B" & rowComPR)

        ' --------------------------------------------------------
        ' Reset MRP detection variables for this material
        ' --------------------------------------------------------
        MRPElement = ""
        foundPlanOrder = False
        planOrderRow = -1
        recordQty = 0

        ' ========================================================
        ' MD03 - Refresh MRP
        ' ========================================================
        SapSession.findById("wnd[0]").maximize

        SapSession.findById( _
            "wnd[0]/tbar[0]/okcd" _
        ).Text = "/nmd03"

        SapSession.findById( _
            "wnd[0]/tbar[0]/btn[0]" _
        ).press

        SapSession.findById( _
            "wnd[0]/usr/ctxtRM61X-MATNR" _
        ).Text = nMaterial

        SapSession.findById( _
            "wnd[0]/usr/ctxtRM61X-BERID" _
        ).Text = nPlant

        SapSession.findById( _
            "wnd[0]/usr/ctxtRM61X-WERKS" _
        ).Text = nPlant

        SapSession.findById( _
            "wnd[0]/tbar[0]/btn[0]" _
        ).press

        ' ========================================================
        ' MD04 - Display MRP
        ' ========================================================
        SapSession.findById("wnd[0]").maximize

        SapSession.findById( _
            "wnd[0]/tbar[0]/okcd" _
        ).Text = "/nmd04"

        SapSession.findById( _
            "wnd[0]/tbar[0]/btn[0]" _
        ).press

        SapSession.findById( _
            "wnd[0]/usr/tabsTAB300/tabpF01/" & _
            "ssubINCLUDE300:SAPMM61R:0301/" & _
            "ctxtRM61R-MATNR" _
        ).Text = nMaterial

        SapSession.findById( _
            "wnd[0]/usr/tabsTAB300/tabpF01/" & _
            "ssubINCLUDE300:SAPMM61R:0301/" & _
            "ctxtRM61R-BERID" _
        ).Text = nPlant

        SapSession.findById( _
            "wnd[0]/usr/tabsTAB300/tabpF01/" & _
            "ssubINCLUDE300:SAPMM61R:0301/" & _
            "ctxtRM61R-WERKS" _
        ).Text = nPlant

        SapSession.findById( _
            "wnd[0]/tbar[0]/btn[0]" _
        ).press

        ' ========================================================
        ' Ensure GR view
        ' ========================================================
        If Trim( _
            SapSession.findById( _
                "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
                "btnBUTTON_DAT00" _
            ).Text _
        ) <> "AV" Then

            SapSession.findById( _
                "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
                "btnBUTTON_DAT00" _
            ).press

        End If

        ' ========================================================
        ' Apply filter
        ' ========================================================
        SapSession.findById( _
            "wnd[0]/tbar[1]/btn[29]" _
        ).press

        SapSession.findById( _
            "wnd[0]/usr/subINCLUDE12XX:SAPMM61R:1200/" & _
            "cmbRM61R-FILBZ" _
        ).Key = "SAP00001"

        ' ========================================================
        ' Read visible MRP rows
        ' ========================================================
        tableViewSize = SapSession.findById( _
            "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
            "tblSAPMM61RTC_EZ" _
        ).VisibleRowCount

        For nR = 0 To tableViewSize - 1

            MRPElement = SapSession.findById( _
                "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
                "tblSAPMM61RTC_EZ/" & _
                "txtMDEZ-DELB0[2," & nR & "]" _
            ).Text

            If Not (MRPElement Like "[_]*") Then

                recordQty = CLng( _
                    SapSession.findById( _
                        "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
                        "tblSAPMM61RTC_EZ/" & _
                        "txtMDEZ-MNG01[8," & nR & "]" _
                    ).Text _
                )

                ' ------------------------------------------------
                ' Special handling for SchLne
                ' ------------------------------------------------
                If MRPElement = "SchLne" Then

                    SapSession.findById( _
                        "wnd[0]/tbar[1]/btn[41]" _
                    ).press

                    SapSession.findById( _
                        "wnd[0]/usr/ctxtRM61X-BANER" _
                    ).Text = "3"

                    SapSession.findById( _
                        "wnd[0]/usr/ctxtRM61X-LIFKZ" _
                    ).Text = "1"

                    SapSession.findById( _
                        "wnd[0]/usr/ctxtRM61X-DISER" _
                    ).Text = "1"

                    SapSession.findById( _
                        "wnd[0]/usr/ctxtRM61X-PLMOD" _
                    ).Text = "1"

                    SapSession.findById("wnd[0]").sendVKey 0
                    SapSession.findById("wnd[0]").sendVKey 0

                    SapSession.findById( _
                        "wnd[0]/tbar[1]/btn[6]" _
                    ).press

                    Exit For

                End If

                ' ------------------------------------------------
                ' Plan Order found
                ' ------------------------------------------------
                If MRPElement = "PlOrd." Then

                    foundPlanOrder = True
                    planOrderRow = nR

                    Exit For

                End If

            End If

        Next nR

        ' ========================================================
        ' Create PR only when Plan Order was found
        ' ========================================================
        If foundPlanOrder Then

            SapSession.findById( _
                "wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/" & _
                "tblSAPMM61RTC_EZ/" & _
                "txtMDEZ-EXTRA[5," & planOrderRow & "]" _
            ).SetFocus

            SapSession.findById("wnd[0]").sendVKey 2

            SapSession.findById( _
                "wnd[1]/tbar[0]/btn[27]" _
            ).press

            SapSession.findById( _
                "wnd[0]/usr/txtMDBA-MENGE" _
            ).Text = Trim(nQuantity)

            SapSession.findById("wnd[0]").sendVKey 0

            ' ----------------------------------------------------
            ' Handle quantity logic
            ' ----------------------------------------------------
            If nQuantity < recordQty Then

                SapSession.findById( _
                    "wnd[0]/usr/txtPLAF-GSMNG" _
                ).Text = ""

            End If

            ' ----------------------------------------------------
            ' Remove vendor if configured
            ' ----------------------------------------------------
            If DRemoveVendor = "X" Then

                SapSession.findById( _
                    "wnd[0]/usr/ctxtRM61P-EPSTP" _
                ).Text = ""

                SapSession.findById( _
                    "wnd[0]/usr/ctxtMDBA-FLIEF" _
                ).Text = ""

                SapSession.findById( _
                    "wnd[0]/usr/ctxtMDBA-KONNR" _
                ).Text = ""

                SapSession.findById( _
                    "wnd[0]/usr/txtMDBA-KONPS" _
                ).Text = ""

                SapSession.findById( _
                    "wnd[0]/usr/txtMDBA-RESWK" _
                ).Text = ""

                SapSession.findById("wnd[0]").sendVKey 0

            End If

            ' ----------------------------------------------------
            ' Save / create PR
            ' ----------------------------------------------------
            SapSession.findById( _
                "wnd[0]/tbar[0]/btn[11]" _
            ).press

            Result = SapSession.findById( _
                "wnd[0]/sbar" _
            ).Text

            CombinePR.Range("C" & rowComPR).Value = Result

        End If

    Next rowComPR

    MsgBox _
        "Done. Please check the result.", _
        vbInformation, _
        "PR Creation"

CleanExit:

    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True

    MsgBox _
        "PR creation stopped." & vbCrLf & vbCrLf & _
        "Row: " & rowComPR & vbCrLf & _
        "Material: " & IIf(nMaterial Is Nothing, "", nMaterial.Value) & vbCrLf & _
        "Error: " & Err.Description, _
        vbCritical, _
        "PR Creation Error"

End Sub

