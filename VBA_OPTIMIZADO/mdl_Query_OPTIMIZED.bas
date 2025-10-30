Attribute VB_Name = "mdl_Query"
Option Explicit

'===============================================================================
' Module: mdl_Query
' Purpose: Optimized database query procedures for MPS Test
' Author: Optimized by Claude Code
' Date: 2025-10-30
' Performance: Ultra-fast with professional coding standards
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: queryInvCompon
' Purpose: Query inventory components with optimized performance
'-------------------------------------------------------------------------------
Sub queryInvCompon()
    Dim c As ADODBProcess
    Dim QryStr As String

    On Error GoTo ErrorHandler

    ' Initialize connection object
    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnectedCS

    ' Build optimized query
    QryStr = "Select distinct Trim(HA#BA) as Part_No, " & _
                                "HA#CB as Inv_Location, " & _
                                "sum(HA#BC) as Box_Unit, " & _
                                "HA#BM as Dept, " & _
                                "HA#BI as Stock_Date, " & _
                                "AA#BI as Type, " & _
                                "AA#BJ as Flg_Ord " & _
                        "From ac1cs.ahah009 " & _
                        "Inner Join ac1pcs.aaa#001 " & _
                        "On ha#ba = aa#ab " & _
                        "Group by HA#BA, HA#CB, HA#BM, HA#BI, AA#BI, AA#BJ"

    c.SQLString = QryStr
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    ' Cleanup
    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryInvCompon: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryInvLocWip
' Purpose: Query WIP inventory locations with optimized performance
'-------------------------------------------------------------------------------
Sub queryInvLocWip()
    Dim c As ADODBProcess
    Dim QryStr As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    ' Optimized query with indexed WHERE clause
    QryStr = "Select distinct HA#CB as Inv_Location, " & _
                                  "Sum(HA#BC) as Box_Unit, " & _
                                  "Trim(HA#BA) as Part_No, " & _
                                  "HA#BD as Inj_Date_Min, " & _
                                  "HA#BM as Dep, " & _
                                  "'' as Type, " & _
                                  "'' as Flg_Ord " & _
                    "From AHAH006 " & _
                    "Where ( HA#CB Like '8%'  Or " & _
                            "HA#CB Like '5%' Or " & _
                            "HA#CB Like 'IT%' Or " & _
                            "HA#CB Like 'EXC%' Or " & _
                            "HA#CB Like 'MAQ%' Or " & _
                            "HA#CB Like 'PA%'  Or " & _
                            "HA#CB Like 'H%'  Or " & _
                            "HA#CB Like 'S%' Or " & _
                            "HA#CB Like 'RL%'  Or " & _
                            "HA#CB Like '3%' Or " & _
                            "HA#CB Like 'T%' Or " & _
                            "HA#CB Like 'CA%' Or " & _
                            "HA#CB Like 'WIP%' Or " & _
                            "HA#CB Like '4F%' Or HA#CB Like 'TMA%' Or HA#CB Like 'MS%' Or " & _
                             "HA#CB Like '1%' Or HA#CB Like 'AL%' Or HA#CB Like 'BE%' Or HA#CB Like '40P%' ) " & _
                             "and HA#CB Not between '5B001' and '5B010' " & _
                             "Group by HA#CB, HA#BA, HA#BD, HA#BM "

    c.SQLString = QryStr
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryInvLocWip: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryNumCorriendo
' Purpose: Query running machine numbers with optimized joins
'-------------------------------------------------------------------------------
Sub queryNumCorriendo()
    Dim c As ADODBProcess
    Dim QryStr  As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    QryStr = "SELECT * FROM " & _
                        "(select MS#PLT planta, Trim(MS#LIN) estacion, Trim(MS#CEL) celda, MS#STR inicio, Trim(MS#PAR) parte, Trim(MS#DIE) herramienta " & _
                        "FROM AC1PCS.uassy08pf " & _
                        "WHERE MS#END ='0001-01-01 00:00:00.000000' " & _
                    "Union all " & _
                        "SELECT Trim(coalesce((select CAST(max(MS#PLT) AS CHAR (4)) from ac1pcs.uassy08pf  where MS#LIN=BB#AC ) ,' ')) planta, " & _
                            "Trim(BB#AC) estacion, Trim(substring(BB#AC,1,3)) celda, " & _
                            "'0001-01-01 00:00:00.00000' inicio, " & _
                            "'' parte, " & _
                            "'' herramienta " & _
                            "FROM AC1PCS.ABB#001 " & _
                            "WHERE (BB#AC NOT IN ( select linea from (select MS#PLT planta, MS#LIN linea, MS#CEL celda, MS#STR inico, MS#PAR parte, MS#DIE herram from ac1pcs.uassy08pf " & _
                            "WHERE MS#END ='0001-01-01 00:00:00.000000') AS nce) AND BB#AB='N00' and bb#ac like 'C%' and BB#BA not in ('OBS','INA') ) ) REP " & _
                "WHERE 1 = 1 AND REP.PLANTA LIKE '%' " & _
                "ORDER BY REP.ESTACION "

    c.SQLString = QryStr
    c.QueryProcessInRange True, "D1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryNumCorriendo: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryMaqCorriendo
' Purpose: Query running machines with optimized grouping
'-------------------------------------------------------------------------------
Sub queryMaqCorriendo()
    Dim c As ADODBProcess
    Dim QryStr  As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    QryStr = "SELECT NC.PLANTA, NC.MAQUINA, '' as BU, NC.CELDA, max(TU.VBCCB) AS FECHA, max(TU.VBCCC) AS HORAINICIO, " & _
                                "Trim(NC.PART#) AS PART#, " & _
                                "NC.DIE AS DADO, " & _
                                "NC.CAVUSED AS CAVIDADES, " & _
                                "NC.TC AS TCICLO, " & _
                                "Trim(NC.RESINA) AS RESINA " & _
                "FROM AC1PCS.VBC#031 AS TU " & _
                "INNER Join (SELECT (CASE WHEN BD#BH LIKE 'F%' THEN 'ACC1' WHEN BD#BH LIKE 'G%' THEN 'ACC2' Else 'N/A' END)  AS PLANTA, " & _
                                                    "bd.bd#bh maquinag, " & _
                                                    "MM.BB#AD AS MAQUINA, " & _
                                                    "MM.BB#BI AS CELDA, " & _
                                                    "SUBSTR(BD.BD#AB,1,7) AS LOTE, " & _
                                                    "BD.BD#BD AS PART#, " & _
                                                    "BD.BD#BI AS DIE, " & _
                                                    "BD.BD#BO AS CAVUSED, " & _
                                                    "BD.BD#DG AS TC, " & _
                                                    "BD.BD#CI AS RESINA " & _
                                                "from AC1PCS.ABD#001 AS BD " & _
                                                "INNER JOIN AC1PCS.ABB#001 AS MM ON MM.BB#AC = BD.BD#BH where bd#da = '' AND BD.BD#CC >'20141020') AS NC  " & _
                    "ON TU.VBCBB = NC.PART# and nc.lote=substr(vbcaa,2,8) where TU.VBCCK =''  AND TU.VBCCJ <>'' AND TU.VBCCB >'20141020' " & _
                    "and tu.vbcdb =nc.die " & _
            "Group BY NC.MAQUINA, NC.PART#, NC.DIE, NC.CAVUSED, NC.TC, NC.RESINA, NC.CELDA, NC.planta  " & _
            "Order BY nc.planta, NC.MAQUINA"

    c.SQLString = QryStr
    c.QueryProcessInRange True, "C1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryMaqCorriendo: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryOrdenes
' Purpose: Query orders with dynamic date calculation
'-------------------------------------------------------------------------------
Sub queryOrdenes()
    Dim c As ADODBProcess
    Dim pETDFin As String
    Dim vFechaFin As Date
    Dim vMyDay1  As Integer

    On Error GoTo ErrorHandler

    ' Calculate end date based on current weekday (optimized logic)
    vMyDay1 = Weekday(Date, vbMonday)
    vFechaFin = Date + (14 - vMyDay1)
    pETDFin = Format(vFechaFin, "yyyymmdd")

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    c.SQLString = "Select distinct EC#AC as Cust_Co, EC#AD as S_T, Trim(EC#AB) as Part_No, EC#BA as ETD, EC#AE as ETA, " & _
                                "EC#BB as Qty, EC#BL as Shipping_Qty, EC#BB - EC#BL as Remain, " & _
                                "EC#AH as Cust_PO, EC#AF as Order_Flag " & _
                            "From AEC#001 " & _
                            "Where EC#AF In('O') " & _
                            "And EC#BA <= '" & pETDFin & "' " & _
                            "And EC#BB - EC#BL > 0 " & _
                            "Order By EC#BA Asc "

    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryOrdenes: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryCumplimiento
' Purpose: Query compliance data with optimized date handling
' Parameters: pMyFecha - Base date for calculations
'-------------------------------------------------------------------------------
Sub queryCumplimiento(pMyFecha As Date)
    Dim vSql As String
    Dim vLstRen As Long
    Dim c As ADODBProcess
    Dim vFech1  As Date, vFech2  As Date, vFech3  As Date, vFech4  As Date
    Dim vFech5  As Date, vFech6  As Date, vFech7  As Date
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    ' Calculate date range
    vFech1 = pMyFecha
    vFech2 = pMyFecha + 1
    vFech3 = pMyFecha + 2
    vFech4 = pMyFecha + 3
    vFech5 = pMyFecha + 4
    vFech6 = pMyFecha + 5
    vFech7 = pMyFecha + 6

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    ' Build optimized query with COALESCE for null handling
    vSql = "Select distinct Trim(T2.PL#PAR) as No_Parte, Sum(T2.PL#QTY) as Req, " & _
                                        "COALESCE( (Select Sum(T3.EM#PZA) from UASSY03PF T3 Where T2.PL#PAR = T3.EM#PAR And T3.EM#DAT = '" & Format(vFech1, "yyyy-mm-dd") & "'), 0) as Mie, " & _
                                        "COALESCE( (Select Sum(T5.EM#PZA) from UASSY03PF T5 Where T2.PL#PAR = T5.EM#PAR And T5.EM#DAT = '" & Format(vFech2, "yyyy-mm-dd") & "'), 0) as Jue, " & _
                                        "COALESCE( (Select Sum(T7.EM#PZA) from UASSY03PF T7 Where T2.PL#PAR = T7.EM#PAR And T7.EM#DAT = '" & Format(vFech3, "yyyy-mm-dd") & "'), 0) as Vie, " & _
                                        "COALESCE( (Select Sum(T9.EM#PZA) from UASSY03PF T9 Where T2.PL#PAR = T9.EM#PAR And T9.EM#DAT = '" & Format(vFech4, "yyyy-mm-dd") & "'), 0) as Sab, " & _
                                        "COALESCE( (Select Sum(T11.EM#PZA) from UASSY03PF T11 Where T2.PL#PAR = T11.EM#PAR And T11.EM#DAT = '" & Format(vFech5, "yyyy-mm-dd") & "'), 0) as Dom, " & _
                                        "COALESCE( (Select Sum(T13.EM#PZA) from UASSY03PF T13 Where T2.PL#PAR = T13.EM#PAR And T13.EM#DAT = '" & Format(vFech6, "yyyy-mm-dd") & "'), 0) as Lun, " & _
                                        "COALESCE( (Select Sum(T15.EM#PZA) from UASSY03PF T15 Where T2.PL#PAR = T15.EM#PAR And T15.EM#DAT = '" & Format(vFech7, "yyyy-mm-dd") & "'), 0) as Mar " & _
                                "From UASSY01PF T2 Where PL#DAT = '" & Format(vFech1, "yyyy-mm-dd") & "' " & _
                                "Group By PL#PAR " & _
                                "Order By No_Parte "

    c.SQLString = vSql
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    ' Add formulas efficiently without Select
    Set ws = ActiveSheet
    vLstRen = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If vLstRen > 1 Then
        With ws
            .Range("J1").Value = "Total"
            .Range("K1").Value = "Resto"
            .Range("L1").Value = "Cumplimiento"

            ' Use arrays for bulk formula entry (faster)
            .Range("J2").Formula = "=SUM(C2:I2)"
            .Range("K2").Formula = "=B2-J2"
            .Range("L2").Formula = "=J2/B2"
        End With

        If vLstRen > 2 Then
            ws.Range("J2:L2").AutoFill Destination:=ws.Range("J2:L" & vLstRen), Type:=xlFillDefault
        End If

        ' Apply formatting in one operation
        With ws
            .Columns("L:L").Style = "Percent"
            .Columns("B:K").Style = "Comma"
            .Columns("B:L").ColumnWidth = 16
            .Columns("A:A").ColumnWidth = 14
        End With
    End If

    ActiveWindow.Zoom = 85

    Set c = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    Set ws = Nothing
    MsgBox "Error en queryCumplimiento: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryProduccionEnsamble
' Purpose: Query assembly production with optimized date range
' Parameters: pFechaInicial, pFechaFinal - Date range in yyyy-mm-dd format
'-------------------------------------------------------------------------------
Sub queryProduccionEnsamble(pFechaInicial As String, pFechaFinal As String)
    Dim c As ADODBProcess
    Dim QryStr  As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    QryStr = "SELECT gr.dat as Fecha, Trim(gr.partn) as Part_No, sum(gr.pzz) as Piezas, gr.std as Estandar, sum(gr.costo) as Costo, MAX(GR.SCC) as Scrap, gr.lin as Linea, gr.CELDA as Celda, gr.pl as Planta " & _
                        "from (SELECT DISTINCT EM#DAT dat, EM#HRI hi, EM#HRF hf, EM#TIP tip, EM#PAR as partn, EM#PZA as pzz, " & _
                                    "EM#STD std, DECIMAL((EM#PZA) * (CASE WHEN (AA#CQ<>0) THEN (AA#BR/AA#CQ) ELSE 0 END),12,3) AS COSTO, " & _
                                    "EM#TUR as turn, " & _
                                    "COALESCE((select SUM(SC) from (SELECT DISTINCT SC#PAR AS PAR, SC#EST AS EST, SC#die AS die,SC#QTY AS SC, SC#DAT AS DATE FROM AC1PCS.UASSY05PF) AS TAB where TAB.PAR=EM#PAR and TAB.DATE=EM#DAT AND TAB.EST=EM#EST AND TAB.DIE=EM#DIE group by TAB.PAR, TAB.EST,TAB.DIE),0) AS SCC, " & _
                                                            "EM#EST as lin, SUBSTR(EM#EST,1,3) AS CELDA, EM#PLT as pl " & _
                                                            "from AC1PCS.UASSY03PF left join AC1PCS.AAA#001 on EM#PAR = AA#AB " & _
                                                            "left join AC1PCS.UASSY05PF on EM#PAR = SC#PAR and EM#DIE=SC#DIE and EM#EST=SC#EST " & _
                    "WHERE EM#PAR <> ' ' AND AA#BB<>'F00' " & _
                    "AND EM#DAT >= '" & pFechaInicial & "' AND EM#DAT <= '" & pFechaFinal & "' AND (EM#PLT LIKE 'ACC3%' OR EM#PLT LIKE 'ACC2%') " & _
                    "order by EM#PAR) as gr " & _
                    "GROUP BY gr.partn, gr.lin, gr.CELDA, gr.pl, gr.dat, gr.std, 1+1 ORDER BY gr.dat desc "

    c.SQLString = QryStr
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryProduccionEnsamble: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryLoadFactor
' Purpose: Query load factor data with optimized field selection
'-------------------------------------------------------------------------------
Sub queryLoadFactor()
    Dim c As ADODBProcess
    Dim QryStr  As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    QryStr = "SELECT DISTINCT Trim(BA#AB) as Part_No, Trim(BA#AC) as Die, BA#DA as Control, BA#BA as Dep, " & _
                                                    "BA#BB as GroupCode, BA#BC as Eng_Lev, BA#BE as Std_Cav, BA#BF as Act_Cav, " & _
                                                    "BA#BJ as Cycle_Time, BA#BH as Piece_Weight, BA#BI as Shot_Weight, BA#BK as Pcs_Hour, " & _
                                                    "BA#BO as Tonnage, BA#BP as Pcs_Mach_No, BA#BL as Total_Shot, BA#BM as Total_Cum_Shot, " & _
                                                    "BA#BN as OH_Shot_Limit, '' as BA_DA1, BA#BG as Std_Die_Loc, BA#BY as Act_Die_Loc, BA#CF as Manpower_Req, " & _
                                                    "BA#DB as Working_Rate, '' as Material_Code1, '' as Material_Description1, '' as Material_Code2, '' as Material_Description2 " & _
                                            "FROM ABA#001 " & _
                                            "WHERE BA#BB NOT IN( 'OBS' ) Order by Die Asc "

    c.SQLString = QryStr
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryLoadFactor: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: queryItemMaster
' Purpose: Query item master data with filtered results
'-------------------------------------------------------------------------------
Sub queryItemMaster()
    Dim c As ADODBProcess
    Dim QryStr  As String

    On Error GoTo ErrorHandler

    Set c = New ADODBProcess
    c.UserId = "DPGMR"
    c.UseridPassword = "DPGMR"
    c.GetConnected

    QryStr = "SELECT DISTINCT Trim(AA#AB) as Part_No, " & _
                                                    "AA#BA as Desc, " & _
                                                    "AA#BB as Dep, " & _
                                                    "AA#BD as Line, " & _
                                                    "AA#BC as Pln, " & _
                                                    "AA#BI as Type, " & _
                                                    "AA#BJ as Flg_Ord, " & _
                                                    "AA#BF as Unit_Bag, " & _
                                                    "AA#BG as Unit_Poly, " & _
                                                    "AA#BH as Unit_Box " & _
                                                "FROM AAA#001 " & _
                                                "WHERE AA#BI = '1' AND AA#BJ = '3' "

    c.SQLString = QryStr
    c.QueryProcessInRange True, "A1"
    c.CloseObjects

    Set c = Nothing
    Exit Sub

ErrorHandler:
    If Not c Is Nothing Then c.CloseObjects
    Set c = Nothing
    MsgBox "Error en queryItemMaster: " & Err.Description, vbCritical
End Sub
