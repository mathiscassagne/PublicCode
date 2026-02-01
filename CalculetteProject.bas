Attribute VB_Name = "Module1"
Option Explicit

' === CONFIG G�N�RALE ===
Private Const CALC_SHEET As String = "Calculette"
Private Const SOURCE_SHEET_NAME As String = "---CQL_TAUX_ACTIVITE_DETAIL_BBB"
Private Const USE_SECOND_SHEET_IF_MISSING As Boolean = True

Private Const FIRST_DATA_ROW As Long = 6
Private Const HEADER_ROW_SRC As Long = 2

' === DRILL-DOWN (cellules filtre) ===
Private Const FILTER_GT_CELL As String = "B1"
Private Const FILTER_SUPP_CELL As String = "B2"
Private Const FILTER_ORDER_CELL As String = "B3"

' === CIBLES (CALCULETTE) ===
Private Const TGT_GT_COL As String = "A"          ' GT
Private Const TGT_SUPP_NO_COL As String = "B"     ' N� fournisseur
Private Const TGT_SUPP_NAME_COL As String = "C"   ' Libell� fournisseur
Private Const TGT_ORDER_COL As String = "D"       ' N� commande
Private Const TGT_GTIN_COL As String = "K"        ' GTIN
Private Const TGT_CODELEC_COL As String = "E"     ' CODELEC (cl� de bloc)
Private Const TGT_AMOUNT_COL As String = "F"      ' Montant ligne
Private Const TGT_ARTICLE_COL As String = "G"     ' Article / libell� article
Private Const TGT_SERVICE_RATE_COL As String = "H" ' Taux de service (K source)
Private Const TGT_OLD_D_COL As String = "I"       ' ex-ancienne D -> I
Private Const TGT_OLD_E_COL As String = "J"       ' ex-ancienne E -> J
Private Const TGT_OLD_G_COL As String = "L"       ' ex-ancienne G -> L
Private Const TGT_QTY_MISS_COL As String = "M"    ' UVC manquantes (par ligne)
Private Const TGT_SURCOST_COL As String = ""     ' Surco�t annexe (42�/commande p�nalisable)
Private Const TGT_PENALTY_COL As String = "U"     ' P�nalit� ventil�e
Private Const TGT_PACK_MISS_COL As String = "P"   ' qt� manquante au conditionnement (r�cap bloc)
Private Const TGT_HELPER_PACKTYPE_COL As String = "V"  ' Helper: "COLIS" / "PALETTE" / "ERREUR"


' Colonnes r�cap bloc
Private Const REC_TOTAL_L_COL As String = "N"     ' total L
Private Const REC_TOTAL_M_COL As String = "O"     ' total M
Private Const REC_PACK_MISS_COL As String = "P"   ' total conditionnement (nouveau)
Private Const REC_TOTAL_AMT_COL As String = "R"   ' total F
Private Const REC_CAP_COL As String = "S"         ' MIN(R;U)
Private Const REC_TWO_PCT_COL As String = ""     ' 2%
Private Const REC_BLOCK_CHOSEN_COL As String = "T"

' === SOURCES (FEUILLE SOURCE) ===
Private Const SRC_GT_COL As String = "C"
Private Const SRC_SUPP_NO_COL As String = "F"
Private Const SRC_SUPP_NAME_COL As String = "G"
Private Const SRC_ORDER_COL As String = "M"
Private Const SRC_CODELEC_COL As String = "Q"
Private Const SRC_AMOUNT_COL As String = "U"
Private Const SRC_ARTICLE_COL As String = "L"
Private Const SRC_OLD_D_COL As String = "R"
Private Const SRC_OLD_E_COL As String = "S"
Private Const SRC_OLD_G_COL As String = "X"
Private Const SRC_QTY_MISS_COL As String = "Y"
Private Const SRC_SERVICE_RATE_COL As String = "K"
Private Const SRC_GTIN_COL As String = "T"

' Colonnes Gamme (fichier externe)
Private Const GAMME_SHEET_NAME As String = "Base"
Private Const GAMME_HEADER_ROW As Long = 5
Private Const GAMME_ARTICLE_COL As String = "D"
Private Const GAMME_PACKTYPE_COL As String = "Y"
Private Const GAMME_UVC_PER_COLIS_COL As String = "Z"
Private Const GAMME_COLIS_PER_PALETTE_COL As String = "AA"

' Feuille "calculatrice" (3e feuille)
Private Const PREJ_SHEET_INDEX As Long = 3
Private Const PREJ_ROW_INPUT As Long = 65
Private Const PREJ_ROW_OUTPUT As Long = 68

' Colonnes de la 3e feuille (selon flux)
Private Const FT_COL_PAL As String = "E"
Private Const FT_COL_COLIS As String = "F"
Private Const NFT_COL_PAL As String = "I"
Private Const NFT_COL_COLIS As String = "J"
Private Const OPP_COL_PAL As String = "M"
Private Const OPP_COL_COLIS As String = "N"

' ====== CACHE GAMME ======
Private Const GAMME_CACHE_SHEET As String = "_GammeCache"
Private Const GAMME_CACHE_HEADER_ROW As Long = 1
Private Const GAMME_CACHE_FIRST_DATA_ROW As Long = 2
Private Const GAMME_CACHE_LASTUPDATED_COL As Long = 7 ' Col G pour le libell� + H pour la date
Private Const GAMME_CACHE_LASTUPDATED_VAL_COL As Long = 8
Private Const GAMME_CACHE_INFO_CELL As String = "B4"  ' Calculette: info "Cache Gamme MAJ: ..."

' === Param�trage surco�t annexe (V2) ===
Private Const SURCOST_DEFAULT As Double = 42#   ' fallback / flux mixtes / non-OPP
Private Const SURCOST_OPP As Double = 28#       ' si commande 100% OPP
Private Const TGT_FLUX_COL As String = "G"      ' Colonne flux dans Calculette: FT/NFT/OPP

' === V2 : int�grer le surco�t dans Q (pas de colonne d�di�e) ===
Private Const V2_INTEGRATE_SURCOST_IN_Q As Boolean = True   ' activer V2

' (Optionnel) cellules de la feuille Calculette pour surcharger les montants
Private Const SURCOST_DEFAULT_CELL As String = "Q2"         ' laisser vide ou un nombre
Private Const SURCOST_OPP_CELL     As String = "Q3"

' Limites de plage pour nettoyages / tri / styles
Private Const REPORT_END_COL As String = "U"           ' derni�re colonne �rapport�e� (avant: "V")
Private Const WORK_END_COL As String = "V"             ' jusqu�o� on nettoie (inclut helper)

' Filtre �lignes vides� : ignorer si montant source vide
Private Const FILTER_ON_AMOUNT_NOT_EMPTY As Boolean = True


' =================== MACROS PRINCIPALES ===================

Public Sub RemplirPuisCalculer()
    Remplir_Depuis_Source_Avec_Filtres
    Calculs_Blocs_V2
    Calculer_Qte_Conditionnement
    Calculer_Prejudice_Q
End Sub
' Cr�e/retourne la feuille cache (tr�s masqu�e)
Private Function EnsureGammeCacheSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GAMME_CACHE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = GAMME_CACHE_SHEET
        ' ent�tes
        ws.Cells(GAMME_CACHE_HEADER_ROW, 1).Value = "ArticleRaw"
        ws.Cells(GAMME_CACHE_HEADER_ROW, 2).Value = "ArticleKey"
        ws.Cells(GAMME_CACHE_HEADER_ROW, 3).Value = "PackType"
        ws.Cells(GAMME_CACHE_HEADER_ROW, 4).Value = "UVC_per_colis_Z"
        ws.Cells(GAMME_CACHE_HEADER_ROW, 5).Value = "Colis_per_palette_AA"
        ws.Cells(GAMME_CACHE_HEADER_ROW, GAMME_CACHE_LASTUPDATED_COL).Value = "LastUpdatedLabel"
        ws.Cells(GAMME_CACHE_HEADER_ROW, GAMME_CACHE_LASTUPDATED_VAL_COL).Value = "LastUpdatedValue"
        ws.Visible = xlSheetVeryHidden
    End If
    Set EnsureGammeCacheSheet = ws
End Function

' Met � jour le cache � partir du fichier Gamme (Base, � partir de la ligne 5)
Public Function Update_Gamme_Cache(Optional ByVal filePath As String = "") As Boolean
    Dim wsCache As Worksheet, wbG As Workbook, wsG As Worksheet
    Dim lastRowG As Long, r As Long, outRow As Long
    Dim art As String, artKey As String, packType As String
    Dim z As Double, aa As Double, f As Variant, openedHere As Boolean
    
    Update_Gamme_Cache = False
    Set wsCache = EnsureGammeCacheSheet()
    On Error Resume Next
    If Len(filePath) = 0 Then
        ' essayer de trouver un classeur "Gamme" d�j� ouvert
        Dim wb As Workbook
        For Each wb In Application.Workbooks
            If InStr(1, UCase$(wb.Name), "GAMME", vbTextCompare) > 0 Then Set wbG = wb: Exit For
        Next wb
    End If
    On Error GoTo 0
    
    If wbG Is Nothing Then
        If Len(filePath) = 0 Then
            f = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Choisir le fichier Gamme (feuille 'Base')")
            If VarType(f) = vbBoolean Then Exit Function
            filePath = CStr(f)
        End If
        Set wbG = Application.Workbooks.Open(filePath, ReadOnly:=True)
        openedHere = True
    End If
    
    On Error Resume Next
    Set wsG = wbG.Worksheets(GAMME_SHEET_NAME) ' "Base"
    On Error GoTo 0
    If wsG Is Nothing Then
        MsgBox "Feuille 'Base' introuvable dans le fichier Gamme.", vbExclamation
        GoTo Quitter
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' purge cache
    With wsCache
        .Cells.ClearContents
        .Cells(GAMME_CACHE_HEADER_ROW, 1).Value = "ArticleRaw"
        .Cells(GAMME_CACHE_HEADER_ROW, 2).Value = "ArticleKey"
        .Cells(GAMME_CACHE_HEADER_ROW, 3).Value = "PackType"
        .Cells(GAMME_CACHE_HEADER_ROW, 4).Value = "UVC_per_colis_Z"
        .Cells(GAMME_CACHE_HEADER_ROW, 5).Value = "Colis_per_palette_AA"
        .Cells(GAMME_CACHE_HEADER_ROW, GAMME_CACHE_LASTUPDATED_COL).Value = "LastUpdatedLabel"
        .Cells(GAMME_CACHE_HEADER_ROW, GAMME_CACHE_LASTUPDATED_VAL_COL).Value = "LastUpdatedValue"
    End With
    
    lastRowG = wsG.Cells(wsG.Rows.Count, ColNum(GAMME_ARTICLE_COL)).End(xlUp).Row
    outRow = GAMME_CACHE_FIRST_DATA_ROW
    
    For r = GAMME_HEADER_ROW To lastRowG
        art = CStr(wsG.Cells(r, ColNum(GAMME_ARTICLE_COL)).Value)   ' D
        artKey = NormalizeKey(art)
        If Len(artKey) > 0 Then
            packType = CStr(wsG.Cells(r, ColNum(GAMME_PACKTYPE_COL)).Value)      ' Y
            z = CDbl(Val(wsG.Cells(r, ColNum(GAMME_UVC_PER_COLIS_COL)).Value))   ' Z
            aa = CDbl(Val(wsG.Cells(r, ColNum(GAMME_COLIS_PER_PALETTE_COL)).Value)) ' AA
            wsCache.Cells(outRow, 1).Value = art
            wsCache.Cells(outRow, 2).Value = artKey
            wsCache.Cells(outRow, 3).Value = packType
            wsCache.Cells(outRow, 4).Value = z
            wsCache.Cells(outRow, 5).Value = aa
            outRow = outRow + 1
        End If
    Next r
    
    ' timestamp
    wsCache.Cells(1, GAMME_CACHE_LASTUPDATED_COL).Value = "LastUpdated"
    wsCache.Cells(1, GAMME_CACHE_LASTUPDATED_VAL_COL).Value = Now
    
    ' masque fort
    wsCache.Visible = xlSheetVeryHidden
    Update_Gamme_Cache = True

Quitter:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If openedHere Then On Error Resume Next: wbG.Close SaveChanges:=False: On Error GoTo 0
End Function

' Charge le dictionnaire depuis le cache (ArticleKey -> Array(packType, Z, AA))
Private Function LoadGammeDictFromCache(ByRef dict As Object) As Boolean
    Dim wsCache As Worksheet, lastRow As Long, r As Long
    Set wsCache = EnsureGammeCacheSheet()
    lastRow = wsCache.Cells(wsCache.Rows.Count, 1).End(xlUp).Row
    If lastRow < GAMME_CACHE_FIRST_DATA_ROW Then
        LoadGammeDictFromCache = False
        Exit Function
    End If
    dict.RemoveAll
    For r = GAMME_CACHE_FIRST_DATA_ROW To lastRow
        Dim k As String, packType As String, z As Double, aa As Double
        k = CStr(wsCache.Cells(r, 2).Value)
        If Len(k) > 0 Then
            packType = CStr(wsCache.Cells(r, 3).Value)
            z = CDbl(Val(wsCache.Cells(r, 4).Value))
            aa = CDbl(Val(wsCache.Cells(r, 5).Value))
            If Not dict.Exists(k) Then dict.Add k, Array(packType, z, aa)
        End If
    Next r
    LoadGammeDictFromCache = (dict.Count > 0)
End Function

' Affiche la date de MAJ du cache en B4 sur la calculette
Public Sub ShowGammeCacheInfo()
    Dim wsCalc As Worksheet, wsCache As Worksheet
    Dim dt As Variant, txt As String
    If Len(CALC_SHEET) > 0 Then Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET) Else Set wsCalc = ActiveSheet
    Set wsCache = EnsureGammeCacheSheet()
    dt = wsCache.Cells(1, GAMME_CACHE_LASTUPDATED_VAL_COL).Value
    If IsDate(dt) Then
        txt = "Cache Gamme MAJ: " & Format(dt, "dd/mm/yyyy HH:nn")
    Else
        txt = "Cache Gamme MAJ: (inconnu)"
    End If
    wsCalc.Range(GAMME_CACHE_INFO_CELL).Value = txt
End Sub

' Bouton pour forcer la MAJ du cache
Public Sub Actualiser_Gamme_Cache()
    If Update_Gamme_Cache() Then
        ShowGammeCacheInfo
        MsgBox "Cache Gamme mis � jour.", vbInformation
    End If
End Sub


Public Sub Remplir_Depuis_Source_Avec_Filtres()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsCalc As Worksheet, wsSrc As Worksheet
    Dim lastRowSrc As Long, r As Long, outRow As Long
    Dim fGT As String, fSupp As String, fOrder As String
    Dim passFilter As Boolean
    
    ' Feuilles
    If Len(CALC_SHEET) > 0 Then
        Set wsCalc = wb.Worksheets(CALC_SHEET)
    Else
        Set wsCalc = ActiveSheet
    End If
    Set wsSrc = GetSourceSheet(wb)
    If wsSrc Is Nothing Then
        MsgBox "Feuille source introuvable.", vbExclamation
        Exit Sub
    End If
    
    ' Filtres (drill-down)
    fGT = Trim$(CStr(wsCalc.Range(FILTER_GT_CELL).Value))       ' B1
    fSupp = Trim$(CStr(wsCalc.Range(FILTER_SUPP_CELL).Value))   ' B2
    fOrder = Trim$(CStr(wsCalc.Range(FILTER_ORDER_CELL).Value)) ' B3
    
    If Len(fGT & fSupp & fOrder) = 0 Then
        MsgBox "Renseignez au moins un filtre (B1 GT, B2 Fournisseur, B3 Commande).", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Reset complet de la zone de travail (contenus + formats)
    With wsCalc.Range("A" & FIRST_DATA_ROW & ":W" & wsCalc.Rows.Count)
        .ClearContents
        .ClearFormats
    End With
    
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, ColNum(SRC_ORDER_COL)).End(xlUp).Row
    outRow = FIRST_DATA_ROW
    
    ' --- Remplissage depuis la source ---
    For r = HEADER_ROW_SRC + 1 To lastRowSrc
        passFilter = True
        
        If Len(fGT) > 0 Then
            passFilter = passFilter And SameKey(wsSrc.Cells(r, ColNum(SRC_GT_COL)).Value, fGT)
        End If
        If Len(fSupp) > 0 Then
            passFilter = passFilter And SameKey(wsSrc.Cells(r, ColNum(SRC_SUPP_NO_COL)).Value, fSupp)
        End If
        If Len(fOrder) > 0 Then
            passFilter = passFilter And SameKey(wsSrc.Cells(r, ColNum(SRC_ORDER_COL)).Value, fOrder)
        End If
        
        If passFilter Then
            If (Not FILTER_ON_AMOUNT_NOT_EMPTY) Or _
               (Trim$(CStr(wsSrc.Cells(r, ColNum(SRC_AMOUNT_COL)).Value)) <> "") Then
               
                ' Nouveaux champs en t�te
                wsCalc.Cells(outRow, ColNum(TGT_GT_COL)).Value = wsSrc.Cells(r, ColNum(SRC_GT_COL)).Value                ' A <- C
                wsCalc.Cells(outRow, ColNum(TGT_SUPP_NO_COL)).Value = wsSrc.Cells(r, ColNum(SRC_SUPP_NO_COL)).Value       ' B <- F
                wsCalc.Cells(outRow, ColNum(TGT_SUPP_NAME_COL)).Value = wsSrc.Cells(r, ColNum(SRC_SUPP_NAME_COL)).Value   ' C <- G
                wsCalc.Cells(outRow, ColNum(TGT_ORDER_COL)).Value = wsSrc.Cells(r, ColNum(SRC_ORDER_COL)).Value           ' D <- M
                
                ' Anciens champs d�cal�s de +4 (+ GTIN en K)
                wsCalc.Cells(outRow, ColNum(TGT_CODELEC_COL)).Value = wsSrc.Cells(r, ColNum(SRC_CODELEC_COL)).Value       ' E <- Q
                wsCalc.Cells(outRow, ColNum(TGT_AMOUNT_COL)).Value = wsSrc.Cells(r, ColNum(SRC_AMOUNT_COL)).Value         ' F <- U
                wsCalc.Cells(outRow, ColNum(TGT_ARTICLE_COL)).Value = wsSrc.Cells(r, ColNum(SRC_ARTICLE_COL)).Value       ' G <- L
                wsCalc.Cells(outRow, ColNum(TGT_SERVICE_RATE_COL)).Value = wsSrc.Cells(r, ColNum(SRC_SERVICE_RATE_COL)).Value ' H <- K
                wsCalc.Cells(outRow, ColNum(TGT_GTIN_COL)).Value = wsSrc.Cells(r, ColNum(SRC_GTIN_COL)).Value             ' K <- T (GTIN)
                wsCalc.Cells(outRow, ColNum(TGT_OLD_D_COL)).Value = wsSrc.Cells(r, ColNum(SRC_OLD_D_COL)).Value           ' I <- R
                wsCalc.Cells(outRow, ColNum(TGT_OLD_E_COL)).Value = wsSrc.Cells(r, ColNum(SRC_OLD_E_COL)).Value           ' J <- S
                ' K = GTIN (rempli ci-dessus)
                wsCalc.Cells(outRow, ColNum(TGT_OLD_G_COL)).Value = wsSrc.Cells(r, ColNum(SRC_OLD_G_COL)).Value           ' L <- X
                wsCalc.Cells(outRow, ColNum(TGT_QTY_MISS_COL)).Value = wsSrc.Cells(r, ColNum(SRC_QTY_MISS_COL)).Value     ' M <- Y
                
                outRow = outRow + 1
            End If
        End If
    Next r
    
    ' Tri pour des blocs lisibles (GT, Fournisseur, Commande, CODELEC)
    If outRow > FIRST_DATA_ROW Then
        With wsCalc.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsCalc.Range(TGT_GT_COL & FIRST_DATA_ROW & ":" & TGT_GT_COL & outRow - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add key:=wsCalc.Range(TGT_SUPP_NO_COL & FIRST_DATA_ROW & ":" & TGT_SUPP_NO_COL & outRow - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add key:=wsCalc.Range(TGT_ORDER_COL & FIRST_DATA_ROW & ":" & TGT_ORDER_COL & outRow - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add key:=wsCalc.Range(TGT_CODELEC_COL & FIRST_DATA_ROW & ":" & TGT_CODELEC_COL & outRow - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsCalc.Range("A" & FIRST_DATA_ROW & ":V" & outRow - 1)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    Else
        MsgBox "Aucune ligne trouv�e pour ces filtres.", vbInformation
    End If
    If outRow > FIRST_DATA_ROW Then
        Dim rngH As Range, c As Range, v
        Set rngH = wsCalc.Range(TGT_SERVICE_RATE_COL & FIRST_DATA_ROW & ":" & TGT_SERVICE_RATE_COL & outRow - 1)
        For Each c In rngH.Cells
            v = c.Value
            If IsNumeric(v) Then
                ' si >1 (ex: 97, 100), on ram�ne sur 0�1
                If CDbl(v) > 1 Then c.Value = CDbl(v) / 100
            End If
        Next c
        rngH.NumberFormatLocal = "0,00%"
    End If
    If outRow > FIRST_DATA_ROW Then
        With wsCalc.Range(TGT_GTIN_COL & FIRST_DATA_ROW & ":" & TGT_GTIN_COL & outRow - 1)
            .NumberFormat = "0"          ' entier, pas de d�cimales, pas de 1,23E+12
        End With
        wsCalc.Columns(TGT_GTIN_COL).AutoFit   ' ajuste la largeur pour tout voir
    End If
    wsCalc.Columns(TGT_SUPP_NAME_COL).AutoFit
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub



Public Sub Calculs_Blocs_V2()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim blockStart As Long, blockEnd As Long, curOrder As String, key As Variant
    Dim nextOrder As String, nextCodelec As Variant
    Dim totalAmt As Double, totalMiss As Double, totalOldG As Double
    Dim miss As Double, twoPct As Double, penaltyCap As Double, rr As Long

    If Len(CALC_SHEET) > 0 Then Set ws = ThisWorkbook.Worksheets(CALC_SHEET) Else Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastRow = ws.Cells(ws.Rows.Count, TGT_CODELEC_COL).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then GoTo Fin

    ' Reset zones r�cap/affichage
    ws.Range(REC_TOTAL_L_COL & FIRST_DATA_ROW & ":" & REC_CAP_COL & lastRow).ClearContents   ' N:S
    ws.Range(REC_BLOCK_CHOSEN_COL & FIRST_DATA_ROW & ":" & REC_BLOCK_CHOSEN_COL & lastRow).ClearContents ' T (ou U si T existe)
    If Len(TGT_SURCOST_COL) > 0 Then ws.Range(TGT_SURCOST_COL & FIRST_DATA_ROW & ":" & TGT_SURCOST_COL & lastRow).ClearContents
    ws.Range(TGT_PENALTY_COL & FIRST_DATA_ROW & ":" & TGT_PENALTY_COL & lastRow).ClearContents

    ' Parcours par blocs (Commande + CODELEC)
    r = FIRST_DATA_ROW
    Do While r <= lastRow
        If Len(Trim$(CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value))) = 0 Or IsRecapRow(ws, r) Then
            r = r + 1
        Else
            key = ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value
            curOrder = CStr(ws.Cells(r, ColNum(TGT_ORDER_COL)).Value)
            blockStart = r: blockEnd = r
            Do While blockEnd + 1 <= lastRow
                If IsRecapRow(ws, blockEnd + 1) Then Exit Do
                nextOrder = CStr(ws.Cells(blockEnd + 1, ColNum(TGT_ORDER_COL)).Value)
                nextCodelec = ws.Cells(blockEnd + 1, ColNum(TGT_CODELEC_COL)).Value
                If (nextOrder = curOrder) And (nextCodelec = key) Then
                    blockEnd = blockEnd + 1
                Else
                    Exit Do
                End If
            Loop

            ' Tri local du bloc par UVC manquantes d�croissant
            With ws.Sort
                .SortFields.Clear
                .SortFields.Add key:=ws.Range(TGT_QTY_MISS_COL & blockStart & ":" & TGT_QTY_MISS_COL & blockEnd), _
                                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange ws.Range("A" & blockStart & ":" & REPORT_END_COL & blockEnd)
                .Header = xlNo: .MatchCase = False: .Orientation = xlTopToBottom
                .Apply
            End With

            ' Totaux bloc
            totalOldG = Application.WorksheetFunction.Sum(ws.Range("L" & blockStart & ":L" & blockEnd))
            totalMiss = Application.WorksheetFunction.Sum(ws.Range(TGT_QTY_MISS_COL & blockStart & ":" & TGT_QTY_MISS_COL & blockEnd))
            totalAmt = Application.WorksheetFunction.Sum(ws.Range(TGT_AMOUNT_COL & blockStart & ":" & TGT_AMOUNT_COL & blockEnd))
            twoPct = totalAmt * 0.02

            ' N,O,R,S
            ws.Cells(blockStart, REC_TOTAL_L_COL).Value = totalOldG
            ws.Cells(blockStart, REC_TOTAL_M_COL).Value = totalMiss
            ws.Cells(blockStart, REC_TOTAL_AMT_COL).Value = totalAmt
            If totalMiss > 0 Then
                penaltyCap = WorksheetFunction.Min(totalAmt, twoPct)
                ws.Cells(blockStart, REC_CAP_COL).Value = penaltyCap
                ws.Cells(blockStart, REC_BLOCK_CHOSEN_COL).ClearContents ' rempli plus tard (capFinal)
            Else
                ws.Cells(blockStart, REC_CAP_COL).ClearContents
                ws.Cells(blockStart, REC_BLOCK_CHOSEN_COL).Value = "Non P�nalisable"
            End If

            ' Nettoyage lisibilit� sur les autres lignes du bloc
            If blockEnd > blockStart Then
                ws.Range(REC_TOTAL_L_COL & (blockStart + 1) & ":" & REC_CAP_COL & blockEnd).ClearContents
                ws.Range(REC_BLOCK_CHOSEN_COL & (blockStart + 1) & ":" & REC_BLOCK_CHOSEN_COL & blockEnd).ClearContents
            End If

            ' Ventilation V au prorata des M si cap > 0
            If totalMiss > 0 And penaltyCap > 0 Then
                Dim alloc As Double, lastAllocRow As Long, remainder As Double
                alloc = 0: lastAllocRow = -1
                For rr = blockStart To blockEnd
                    miss = Val(ws.Cells(rr, ColNum(TGT_QTY_MISS_COL)).Value)
                    If miss > 0 Then
                        ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value = penaltyCap * (miss / totalMiss)
                        alloc = alloc + ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value
                        lastAllocRow = rr
                    Else
                        ws.Cells(rr, ColNum(TGT_PENALTY_COL)).ClearContents
                    End If
                Next rr
                If lastAllocRow <> -1 Then
                    remainder = penaltyCap - alloc
                    ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value = ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value + remainder
                End If
            Else
                ws.Range(TGT_PENALTY_COL & blockStart & ":" & TGT_PENALTY_COL & blockEnd).ClearContents
            End If

            r = blockEnd + 1
        End If
    Loop

    ' ===== R�CAP PAR COMMANDE : insertion lignes �R�CAP CDE� + int�gration surco�t dans Q =====
    Dim groupStarts() As Long, groupEnds() As Long, groupOrders() As String, cnt As Long
    Dim startCur As Long, endCur As Long, ord As String, i As Long
    Dim sumL As Double, sumM As Double, sumN As Double, sumO As Double, sumR As Double, sumS As Double
    Dim scDefault As Double, scOPP As Double, allOPP As Boolean, flux As String, surcostVal As Double, rr2 As Long

    ' rep�rer les groupes �par commande�
    r = FIRST_DATA_ROW: cnt = 0
    Do While r <= lastRow
        ord = CStr(ws.Cells(r, ColNum(TGT_ORDER_COL)).Value)
        If Len(ord) = 0 Then
            r = r + 1
        Else
            startCur = r: endCur = r
            Do While endCur + 1 <= lastRow And CStr(ws.Cells(endCur + 1, ColNum(TGT_ORDER_COL)).Value) = ord
                endCur = endCur + 1
            Loop
            cnt = cnt + 1
            ReDim Preserve groupStarts(1 To cnt)
            ReDim Preserve groupEnds(1 To cnt)
            ReDim Preserve groupOrders(1 To cnt)
            groupStarts(cnt) = startCur: groupEnds(cnt) = endCur: groupOrders(cnt) = ord
            r = endCur + 1
        End If
    Loop

    Call GetSurcostConfig(ws, scDefault, scOPP)

    For i = cnt To 1 Step -1
        startCur = groupStarts(i): endCur = groupEnds(i): ord = groupOrders(i)

        sumL = Application.WorksheetFunction.Sum(ws.Range("L" & startCur & ":L" & endCur))
        sumM = Application.WorksheetFunction.Sum(ws.Range("M" & startCur & ":M" & endCur))
        sumN = Application.WorksheetFunction.Sum(ws.Range(REC_TOTAL_L_COL & startCur & ":" & REC_TOTAL_L_COL & endCur))
        sumO = Application.WorksheetFunction.Sum(ws.Range(REC_TOTAL_M_COL & startCur & ":" & REC_TOTAL_M_COL & endCur))
        sumR = Application.WorksheetFunction.Sum(ws.Range(REC_TOTAL_AMT_COL & startCur & ":" & REC_TOTAL_AMT_COL & endCur))
        sumS = Application.WorksheetFunction.Sum(ws.Range(REC_CAP_COL & startCur & ":" & REC_CAP_COL & endCur))

        ' D�tection 100% OPP (blancs autoris�s)
        allOPP = True
        For rr2 = startCur To endCur
            flux = UCase$(Trim$(CStr(ws.Cells(rr2, ColNum(TGT_FLUX_COL)).Value)))
            If Len(flux) > 0 And flux <> "OPP" Then allOPP = False
        Next rr2
        If sumM > 0 Then surcostVal = IIf(allOPP, scOPP, scDefault) Else surcostVal = 0

        ' Ins�rer la ligne r�cap au-dessus
        ws.Rows(startCur).Insert xlShiftDown
        ws.Cells(startCur, ColNum(TGT_GT_COL)).Value = ws.Cells(startCur + 1, ColNum(TGT_GT_COL)).Value
        ws.Cells(startCur, ColNum(TGT_SUPP_NO_COL)).Value = ws.Cells(startCur + 1, ColNum(TGT_SUPP_NO_COL)).Value
        ws.Cells(startCur, ColNum(TGT_SUPP_NAME_COL)).Value = ws.Cells(startCur + 1, ColNum(TGT_SUPP_NAME_COL)).Value
        ws.Cells(startCur, ColNum(TGT_ORDER_COL)).Value = ord

        ws.Cells(startCur, ColNum(TGT_CODELEC_COL)).Value = "R�CAP CDE" & vbLf & ord
        ws.Cells(startCur, ColNum(TGT_CODELEC_COL)).WrapText = True

        ws.Cells(startCur, "L").Value = sumL
        ws.Cells(startCur, "M").Value = sumM
        ws.Cells(startCur, REC_TOTAL_L_COL).Value = sumN
        ws.Cells(startCur, REC_TOTAL_M_COL).Value = sumO
        ws.Cells(startCur, REC_TOTAL_AMT_COL).Value = sumR
        ws.Cells(startCur, REC_CAP_COL).Value = sumS

        ' Q (r�cap) = S Q blocs + surco�t int�gr� (V2)
        Dim rr3 As Long, sumQBlocks As Double
        sumQBlocks = 0
        For rr3 = startCur + 1 To endCur + 1
            If Not IsRecapRow(ws, rr3) Then
                If IsNumeric(ws.Cells(rr3, "Q").Value) Then sumQBlocks = sumQBlocks + CDbl(ws.Cells(rr3, "Q").Value)
            End If
        Next rr3
        If V2_INTEGRATE_SURCOST_IN_Q And (surcostVal > 0) Then
            ws.Cells(startCur, "Q").Value = CeilTo(sumQBlocks + surcostVal, 0.01)
        Else
            ws.Cells(startCur, "Q").Value = sumQBlocks
        End If
        ws.Cells(startCur, "Q").NumberFormatLocal = "# ##0,00 �"

        ' Style de la ligne r�cap (limit� � A:REPORT_END_COL)
        With ws.Range(Chr(Asc(REPORT_END_COL) + 1) & startCur & ":XFD" & startCur)
            .Interior.ColorIndex = xlColorIndexNone
            .Font.Bold = False
        End With
        With ws.Range("A" & startCur & ":" & REPORT_END_COL & startCur)
            .Font.Bold = True
            .Interior.Color = RGB(222, 235, 247)
        End With
    Next i

    ' Formats �
    lastRow = ws.Cells(ws.Rows.Count, TGT_CODELEC_COL).End(xlUp).Row
    ws.Range(REC_TOTAL_AMT_COL & FIRST_DATA_ROW & ":" & REC_TOTAL_AMT_COL & lastRow).NumberFormatLocal = "# ##0,00 �"
    ws.Range(REC_CAP_COL & FIRST_DATA_ROW & ":" & REC_CAP_COL & lastRow).NumberFormatLocal = "# ##0,00 �"
    ws.Range(REC_BLOCK_CHOSEN_COL & FIRST_DATA_ROW & ":" & REC_BLOCK_CHOSEN_COL & lastRow).NumberFormatLocal = "# ##0,00 �"
    ws.Range(TGT_PENALTY_COL & FIRST_DATA_ROW & ":" & TGT_PENALTY_COL & lastRow).NumberFormatLocal = "# ##0,00 �"
    If Len(TGT_SURCOST_COL) > 0 Then ws.Range(TGT_SURCOST_COL & FIRST_DATA_ROW & ":" & TGT_SURCOST_COL & lastRow).NumberFormatLocal = "# ##0,00 �"

    ' Centrage & bordures A:REPORT_END_COL
    If lastRow >= FIRST_DATA_ROW Then
        With ws.Range("A" & FIRST_DATA_ROW & ":" & REPORT_END_COL & lastRow)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Dim rngTable As Range, b As Variant
        Set rngTable = ws.Range("A" & FIRST_DATA_ROW & ":" & REPORT_END_COL & lastRow)
        For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
            With rngTable.Borders(b)
                .LineStyle = xlContinuous
                .Weight = xlHairline
            End With
        Next b
    End If

Fin:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Public Sub Calculer_Qte_Conditionnement()
    Dim ws As Worksheet, lastRow As Long
    Dim wbG As Workbook, wsG As Worksheet
    Dim f As Variant, openedHere As Boolean
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, lastRowG As Long
    
    Dim key As Variant, curOrder As String
    Dim blockStart As Long, blockEnd As Long
    Dim nextOrder As String, nextCodelec As Variant
    
    Dim art As String, artKey As String, packType As String
    Dim uvcPerColis As Double, colisPerPalette As Double, denom As Double
    Dim miss As Double, sumCond As Double, sumCondRounded As Double, hasErr As Boolean, rr As Long
    Dim isRecap As Boolean
    
    If Len(CALC_SHEET) > 0 Then Set ws = ThisWorkbook.Worksheets(CALC_SHEET) Else Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, TGT_CODELEC_COL).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "Aucune donn�e � traiter.", vbInformation
        Exit Sub
    End If
    ' ====== Charger dict depuis le CACHE (et proposer MAJ si vide) ======
    If Not LoadGammeDictFromCache(dict) Then
        If MsgBox("Le cache Gamme est vide. Actualiser maintenant ?", vbYesNo + vbQuestion) = vbYes Then
            If Update_Gamme_Cache() Then
                If Not LoadGammeDictFromCache(dict) Then
                    MsgBox "�chec du chargement du cache.", vbExclamation
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    ShowGammeCacheInfo

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ---------- Reset P (r�cap conditionnement) et W (helper) ----------
    ws.Range(REC_PACK_MISS_COL & FIRST_DATA_ROW & ":" & REC_PACK_MISS_COL & lastRow).ClearContents  ' P
    ws.Range(TGT_HELPER_PACKTYPE_COL & FIRST_DATA_ROW & ":" & TGT_HELPER_PACKTYPE_COL & lastRow).ClearContents ' W
    ws.Range(REC_PACK_MISS_COL & FIRST_DATA_ROW & ":" & REC_PACK_MISS_COL & lastRow).Interior.ColorIndex = xlColorIndexNone
    
    ' ---------- Parcours par blocs (Commande + CODELEC) ----------
    r = FIRST_DATA_ROW
    Do While r <= lastRow
        ' ignorer les lignes "R�CAP CDE"
        isRecap = (InStr(1, CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value), "R�CAP CDE", vbTextCompare) > 0) _
                  Or (InStr(1, CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value), "RECAP CDE", vbTextCompare) > 0)
        If isRecap Then
            r = r + 1
        Else
            key = ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value
            curOrder = CStr(ws.Cells(r, ColNum(TGT_ORDER_COL)).Value)
            If Len(Trim$(CStr(key))) = 0 Then
                r = r + 1
            Else
                blockStart = r: blockEnd = r
                Do While blockEnd + 1 <= lastRow
                    nextOrder = CStr(ws.Cells(blockEnd + 1, ColNum(TGT_ORDER_COL)).Value)
                    nextCodelec = ws.Cells(blockEnd + 1, ColNum(TGT_CODELEC_COL)).Value
                    isRecap = (InStr(1, CStr(ws.Cells(blockEnd + 1, ColNum(TGT_CODELEC_COL)).Value), "R�CAP CDE", vbTextCompare) > 0) _
                              Or (InStr(1, CStr(ws.Cells(blockEnd + 1, ColNum(TGT_CODELEC_COL)).Value), "RECAP CDE", vbTextCompare) > 0)
                    If isRecap Then Exit Do
                    If (nextOrder = curOrder) And (nextCodelec = key) Then
                        blockEnd = blockEnd + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                ' -- calcul du total "conditionnement" pour le bloc --
                sumCond = 0: hasErr = False
                For rr = blockStart To blockEnd
                    miss = Val(ws.Cells(rr, ColNum(TGT_QTY_MISS_COL)).Value) ' M
                    If miss > 0 Then
                        art = CStr(ws.Cells(rr, ColNum(TGT_OLD_D_COL)).Value) ' I = n� article
                        artKey = NormalizeKey(art)
                        If dict.Exists(artKey) Then
                            packType = LCase$(CStr(dict(artKey)(0)))       ' Y
                            uvcPerColis = CDbl(dict(artKey)(1))            ' Z
                            colisPerPalette = CDbl(dict(artKey)(2))        ' AA
                            
                            ' r�gle : "colis" OU "box" -> Z ; sinon -> AA
                            If (InStr(1, packType, "colis", vbTextCompare) > 0) Or _
                               (InStr(1, packType, "box", vbTextCompare) > 0) Then
                                denom = uvcPerColis
                                ws.Cells(rr, ColNum(TGT_HELPER_PACKTYPE_COL)).Value = "COLIS"
                            Else
                                denom = colisPerPalette
                                ws.Cells(rr, ColNum(TGT_HELPER_PACKTYPE_COL)).Value = "PALETTE"
                            End If
                            If denom > 0 Then
                                ' arrondi au sup�rieur PAR LIGNE
                                sumCond = sumCond + Application.WorksheetFunction.RoundUp(miss / denom, 0)
                            Else
                                hasErr = True
                                ws.Cells(rr, ColNum(TGT_HELPER_PACKTYPE_COL)).Value = "ERREUR"
                            End If
                        Else
                            hasErr = True
                            ws.Cells(rr, ColNum(TGT_HELPER_PACKTYPE_COL)).Value = "ERREUR"
                        End If
                    End If
                Next rr
                
                ' �crire P (arrondi sup�rieur, entier) sur la 1�re ligne du bloc
                If hasErr Then
                    ws.Cells(blockStart, ColNum(REC_PACK_MISS_COL)).Value = "ERREUR"
                    ws.Cells(blockStart, ColNum(REC_PACK_MISS_COL)).Interior.Color = vbYellow
                Else
                    sumCondRounded = Application.WorksheetFunction.RoundUp(sumCond, 0)
                    ws.Cells(blockStart, ColNum(REC_PACK_MISS_COL)).Value = sumCondRounded
                    ws.Cells(blockStart, ColNum(REC_PACK_MISS_COL)).NumberFormat = "0"
                    ws.Cells(blockStart, ColNum(REC_PACK_MISS_COL)).Interior.ColorIndex = xlColorIndexNone
                End If
                ' vider P sur les autres lignes du bloc
                If blockEnd > blockStart Then
                    ws.Range(REC_PACK_MISS_COL & (blockStart + 1) & ":" & REC_PACK_MISS_COL & blockEnd).ClearContents
                End If
                
                r = blockEnd + 1
            End If
        End If
    Loop
    
    ' ---------- Totaux P (r�cap commande) : somme et arrondi sup�rieur ----------
    Dim rowCur As Long, nextRecap As Long, sumPcmd As Double, sumPcmdRounded As Double, anyErr As Boolean
    Dim iRow As Long, v As Variant
    rowCur = FIRST_DATA_ROW
    Do While rowCur <= lastRow
        isRecap = (InStr(1, CStr(ws.Cells(rowCur, ColNum(TGT_CODELEC_COL)).Value), "R�CAP CDE", vbTextCompare) > 0) _
                  Or (InStr(1, CStr(ws.Cells(rowCur, ColNum(TGT_CODELEC_COL)).Value), "RECAP CDE", vbTextCompare) > 0)
        If isRecap Then
            ' prochaine r�cap (ou fin)
            nextRecap = rowCur + 1
            Do While nextRecap <= lastRow _
                And Not ((InStr(1, CStr(ws.Cells(nextRecap, ColNum(TGT_CODELEC_COL)).Value), "R�CAP CDE", vbTextCompare) > 0) _
                         Or (InStr(1, CStr(ws.Cells(nextRecap, ColNum(TGT_CODELEC_COL)).Value), "RECAP CDE", vbTextCompare) > 0))
                nextRecap = nextRecap + 1
            Loop
            
            sumPcmd = 0: anyErr = False
            For iRow = rowCur + 1 To nextRecap - 1
                v = ws.Cells(iRow, ColNum(REC_PACK_MISS_COL)).Value   ' P
                If VarType(v) = vbString Then
                    If UCase$(CStr(v)) = "ERREUR" Then anyErr = True: Exit For
                ElseIf IsNumeric(v) Then
                    sumPcmd = sumPcmd + CDbl(v)
                End If
            Next iRow
            
            If anyErr Then
                ws.Cells(rowCur, ColNum(REC_PACK_MISS_COL)).Value = "ERREUR"
                ws.Cells(rowCur, ColNum(REC_PACK_MISS_COL)).Interior.Color = vbYellow
            Else
                sumPcmdRounded = Application.WorksheetFunction.RoundUp(sumPcmd, 0)
                ws.Cells(rowCur, ColNum(REC_PACK_MISS_COL)).Value = sumPcmdRounded
                ws.Cells(rowCur, ColNum(REC_PACK_MISS_COL)).NumberFormat = "0"
                ws.Cells(rowCur, ColNum(REC_PACK_MISS_COL)).Interior.ColorIndex = xlColorIndexNone
            End If
            
            rowCur = nextRecap
        Else
            rowCur = rowCur + 1
        End If
    Loop

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If openedHere Then On Error Resume Next: wbG.Close SaveChanges:=False: On Error GoTo 0
End Sub

Public Sub Calculer_Prejudice_Q()
    Dim ws As Worksheet, wsPrej As Worksheet
    Dim lastRow As Long, rowCur As Long, nextRecap As Long
    Dim r As Long, rr As Long
    
    If Len(CALC_SHEET) > 0 Then
        Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    Else
        Set ws = ActiveSheet
    End If
    lastRow = ws.Cells(ws.Rows.Count, ColNum(TGT_CODELEC_COL)).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub
    
    On Error Resume Next
    Set wsPrej = ThisWorkbook.Worksheets(PREJ_SHEET_INDEX)
    On Error GoTo 0
    If wsPrej Is Nothing Then
        MsgBox "Impossible de trouver la 3e feuille (calculatrice).", vbExclamation
        Exit Sub
    End If
    
    ' Charger Gamme (cache)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If Not LoadGammeDictFromCache(dict) Then
        If MsgBox("Le cache Gamme est vide. Actualiser maintenant ?", vbYesNo + vbQuestion) = vbYes Then
            If Not Update_Gamme_Cache() Then Exit Sub
            If Not LoadGammeDictFromCache(dict) Then
                MsgBox "�chec du chargement du cache.", vbExclamation
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    ShowGammeCacheInfo
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Reset Q bloc + choix bloc + ventilations
    ws.Range("Q" & FIRST_DATA_ROW & ":Q" & lastRow).ClearContents
    ws.Range(REC_BLOCK_CHOSEN_COL & FIRST_DATA_ROW & ":" & REC_BLOCK_CHOSEN_COL & lastRow).ClearContents
    ws.Range(TGT_PENALTY_COL & FIRST_DATA_ROW & ":" & TGT_PENALTY_COL & lastRow).ClearContents
    
    ' === Parcours par COMMANDE (de r�cap � r�cap) ===
    rowCur = FIRST_DATA_ROW
    Do While rowCur <= lastRow
        If Not IsRecapRow(ws, rowCur) Then
            rowCur = rowCur + 1
        Else
            ' borne du groupe commande
            nextRecap = rowCur + 1
            Do While nextRecap <= lastRow And Not IsRecapRow(ws, nextRecap)
                nextRecap = nextRecap + 1
            Loop
            
            ' --- 1) Lister les BLOCS (par CODELEC) et collecter donn�es brutes ---
            Dim bCnt As Long
            Dim bStarts() As Long, bEnds() As Long
            Dim bMiss() As Double, bQraw() As Double, bS() As Double
            Dim bHasErr() As Boolean
            Dim sOrder As String, codelec As Variant
            Dim bsBlk As Long, beBlk As Long
            Dim flux As String
            
            bCnt = 0
            sOrder = CStr(ws.Cells(rowCur, ColNum(TGT_ORDER_COL)).Value)
            r = rowCur + 1
            Do While r <= nextRecap - 1
                If Len(Trim$(CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value))) = 0 Then
                    r = r + 1
                Else
                    ' d�but de bloc
                    codelec = ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value
                    bsBlk = r: beBlk = r
                    Do While beBlk + 1 <= nextRecap - 1
                        If ws.Cells(beBlk + 1, ColNum(TGT_ORDER_COL)).Value <> sOrder Then Exit Do
                        If ws.Cells(beBlk + 1, ColNum(TGT_CODELEC_COL)).Value <> codelec Then Exit Do
                        beBlk = beBlk + 1
                    Loop
                    
                    ' agrandir chaque tableau (un par un)
                    bCnt = bCnt + 1
                    ReDim Preserve bStarts(1 To bCnt): bStarts(bCnt) = bsBlk
                    ReDim Preserve bEnds(1 To bCnt):   bEnds(bCnt) = beBlk
                    ReDim Preserve bMiss(1 To bCnt)
                    ReDim Preserve bQraw(1 To bCnt)
                    ReDim Preserve bS(1 To bCnt)
                    ReDim Preserve bHasErr(1 To bCnt)
                    
                    ' total manquants du bloc (prend O sur 1re ligne si dispo, sinon recalcule)
                    Dim missB As Double
                    missB = ToDbl(ws.Cells(bsBlk, REC_TOTAL_M_COL).Value)
                    If missB = 0 Then
                        missB = Application.WorksheetFunction.Sum(ws.Range(TGT_QTY_MISS_COL & bsBlk & ":" & TGT_QTY_MISS_COL & beBlk))
                    End If
                    bMiss(bCnt) = missB
                    
                    ' plafond S du bloc
                    bS(bCnt) = ToDbl(ws.Cells(bsBlk, REC_CAP_COL).Value)
                    
                    ' ==== Q "brut" du bloc via la 3e feuille ====
                    Dim hasErr As Boolean, sumColis As Double, sumPal As Double
                    hasErr = False: sumColis = 0: sumPal = 0
                    
                    ' flux (1re ligne du bloc)
                    flux = UCase$(Trim$(CStr(ws.Cells(bsBlk, ColNum(TGT_FLUX_COL)).Value)))
                    
                    ' agr�gation volumes par pack
                    Dim art As String, artKey As String, miss As Double
                    Dim packType As String, uvcPerColis As Double, colisPerPalette As Double, denom As Double
                    For rr = bsBlk To beBlk
                        miss = Val(ws.Cells(rr, ColNum(TGT_QTY_MISS_COL)).Value)
                        If miss > 0 Then
                            art = CStr(ws.Cells(rr, ColNum(TGT_OLD_D_COL)).Value)
                            artKey = NormalizeKey(art)
                            If dict.Exists(artKey) Then
                                packType = LCase$(CStr(dict(artKey)(0)))
                                uvcPerColis = CDbl(dict(artKey)(1))
                                colisPerPalette = CDbl(dict(artKey)(2))
                                If (InStr(1, packType, "colis", vbTextCompare) > 0) Or (InStr(1, packType, "box", vbTextCompare) > 0) Then
                                    denom = uvcPerColis
                                    If denom > 0 Then sumColis = sumColis + (miss / denom) Else hasErr = True
                                Else
                                    denom = colisPerPalette
                                    If denom > 0 Then sumPal = sumPal + (miss / denom) Else hasErr = True
                                End If
                            Else
                                hasErr = True
                            End If
                        End If
                    Next rr
                    
                    Dim amountColis As Double, amountPal As Double
                    If Not hasErr Then
                        Dim cPal As String, cCol As String
                        Select Case flux
                            Case "FT":  cPal = FT_COL_PAL:  cCol = FT_COL_COLIS
                            Case "NFT": cPal = NFT_COL_PAL: cCol = NFT_COL_COLIS
                            Case "OPP": cPal = OPP_COL_PAL: cCol = OPP_COL_COLIS
                            Case Else:  hasErr = True
                        End Select
                        If Not hasErr Then
                            wsPrej.Cells(PREJ_ROW_INPUT, ColNum(cPal)).Value = sumPal
                            wsPrej.Cells(PREJ_ROW_INPUT, ColNum(cCol)).Value = sumColis
                            wsPrej.Calculate
                            amountPal = ToDbl(wsPrej.Cells(PREJ_ROW_OUTPUT, ColNum(cPal)).Value)
                            amountColis = ToDbl(wsPrej.Cells(PREJ_ROW_OUTPUT, ColNum(cCol)).Value)
                            wsPrej.Cells(PREJ_ROW_INPUT, ColNum(cPal)).Value = 0
                            wsPrej.Cells(PREJ_ROW_INPUT, ColNum(cCol)).Value = 0
                        End If
                    End If
                    
                    bHasErr(bCnt) = hasErr
                    If hasErr Then
                        ws.Cells(bsBlk, "Q").Value = "ERREUR"
                        ws.Cells(bsBlk, "Q").Interior.Color = vbYellow
                        bQraw(bCnt) = 0
                    Else
                        bQraw(bCnt) = amountPal + amountColis
                        ws.Cells(bsBlk, "Q").Value = bQraw(bCnt)
                        ws.Cells(bsBlk, "Q").NumberFormatLocal = "# ##0,00 �"
                        ws.Cells(bsBlk, "Q").Interior.ColorIndex = xlColorIndexNone
                    End If
                    
                    ' vider Q sur les autres lignes du bloc (lisibilit�)
                    If beBlk > bsBlk Then ws.Range("Q" & (bsBlk + 1) & ":Q" & beBlk).ClearContents
                    
                    r = beBlk + 1
                End If
            Loop
            
            ' --- 2) Surco�t commande & r�partition prorata O (blocs p�nalisables) ---
            Dim scDefault As Double, scOPP As Double, allOPP As Boolean
            Dim sumMissOrder As Double, surcostValOrder As Double
            Call GetSurcostConfig(ws, scDefault, scOPP)
            
            ' 100% OPP (ignorer blancs)
            allOPP = True
            For r = rowCur + 1 To nextRecap - 1
                flux = UCase$(Trim$(CStr(ws.Cells(r, ColNum(TGT_FLUX_COL)).Value)))
                If Len(flux) > 0 And flux <> "OPP" Then allOPP = False
            Next r
            
            ' somme des O (par blocs p�nalisables)
            sumMissOrder = 0
            For r = 1 To bCnt
                If bMiss(r) > 0 Then sumMissOrder = sumMissOrder + bMiss(r)
            Next r
            
            If sumMissOrder > 0 Then
                surcostValOrder = IIf(allOPP, SURCOST_OPP, SURCOST_DEFAULT)
            Else
                surcostValOrder = 0
            End If
            
            ' --- 3) Par bloc : Q_adj = Q_raw + part surco�t ; capFinal = MIN(S ; Q_adj) ; ventilation ---
            Dim sumVOrder As Double: sumVOrder = 0
            Dim i As Long
            For i = 1 To bCnt
                If bHasErr(i) Then
                    ws.Range(TGT_PENALTY_COL & bStarts(i) & ":" & TGT_PENALTY_COL & bEnds(i)).ClearContents
                Else
                    Dim share As Double, Qadj As Double, capFinal As Double
                    If (bMiss(i) > 0) And (sumMissOrder > 0) And (surcostValOrder > 0) Then
                        share = surcostValOrder * (bMiss(i) / sumMissOrder)
                    Else
                        share = 0
                    End If
                    
                    Qadj = bQraw(i) + share
                    ws.Cells(bStarts(i), "Q").Value = Qadj
                    ws.Cells(bStarts(i), "Q").NumberFormatLocal = "# ##0,00 �"
                    
                    capFinal = WorksheetFunction.Min(bS(i), Qadj)
                    capFinal = CeilTo(capFinal, 0.01)
                    ws.Cells(bStarts(i), REC_BLOCK_CHOSEN_COL).Value = capFinal
                    ws.Cells(bStarts(i), REC_BLOCK_CHOSEN_COL).NumberFormatLocal = "# ##0,00 �"
                    
                    ' ventilation U au prorata des manquants du bloc
                    Dim alloc As Double, lastAllocRow As Long, remainder As Double, missLine As Double
                    alloc = 0: lastAllocRow = -1
                    If capFinal > 0 And bMiss(i) > 0 Then
                        For rr = bStarts(i) To bEnds(i)
                            missLine = Val(ws.Cells(rr, ColNum(TGT_QTY_MISS_COL)).Value)
                            If missLine > 0 Then
                                ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value = capFinal * (missLine / bMiss(i))
                                alloc = alloc + ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value
                                lastAllocRow = rr
                            Else
                                ws.Cells(rr, ColNum(TGT_PENALTY_COL)).ClearContents
                            End If
                        Next rr
                        If lastAllocRow <> -1 Then
                            remainder = capFinal - alloc
                            ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value = ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value + remainder
                        End If
                        sumVOrder = sumVOrder + capFinal
                    Else
                        ws.Range(TGT_PENALTY_COL & bStarts(i) & ":" & TGT_PENALTY_COL & bEnds(i)).ClearContents
                    End If
                End If
            Next i
            
            ' --- 4) R�cap commande : Q = S Q_blocs ajust�s ; U = S ventilations (pas d�ajout surco�t ici) ---
            Dim sumQOrder As Double: sumQOrder = 0
            For i = 1 To bCnt
                If Not bHasErr(i) Then sumQOrder = sumQOrder + ToDbl(ws.Cells(bStarts(i), "Q").Value)
            Next i
            
            ws.Cells(rowCur, "Q").Value = CeilTo(sumQOrder, 0.01)
            ws.Cells(rowCur, "Q").NumberFormatLocal = "# ##0,00 �"
            
            ws.Cells(rowCur, "T").Value = CeilTo(sumVOrder, 0.01)
            ws.Cells(rowCur, "T").NumberFormatLocal = "# ##0,00 �"
            
            rowCur = nextRecap
        End If
    Loop
    
Fin:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Public Sub Ventiler_Q_Manuel()
    Dim ws As Worksheet
    Dim lastRow As Long, rowCur As Long, nextRecap As Long, rr As Long

    If Len(CALC_SHEET) > 0 Then Set ws = ThisWorkbook.Worksheets(CALC_SHEET) Else Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ColNum(TGT_CODELEC_COL)).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Vider ventilations
    ws.Range(TGT_PENALTY_COL & FIRST_DATA_ROW & ":" & TGT_PENALTY_COL & lastRow).ClearContents
    ws.Range(REC_BLOCK_CHOSEN_COL & FIRST_DATA_ROW & ":" & REC_BLOCK_CHOSEN_COL & lastRow).ClearContents

    ' Parcours par COMMANDE
    rowCur = FIRST_DATA_ROW
    Do While rowCur <= lastRow
        If Not IsRecapRow(ws, rowCur) Then
            rowCur = rowCur + 1
        Else
            nextRecap = rowCur + 1
            Do While nextRecap <= lastRow And Not IsRecapRow(ws, nextRecap)
                nextRecap = nextRecap + 1
            Loop

            ' Lister blocs + collecter O, S, Q (SAISIS manuellement sur 1re ligne de bloc)
            Dim bStarts() As Long, bEnds() As Long, bCnt As Long
            Dim bMiss() As Double, bQmanual() As Double, bS() As Double
            Dim i As Long, r As Long, sOrder As String, codelec As Variant

            bCnt = 0
            sOrder = CStr(ws.Cells(rowCur, ColNum(TGT_ORDER_COL)).Value)

            r = rowCur + 1
            Do While r <= nextRecap - 1
                If Len(Trim$(CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value))) = 0 Then
                    r = r + 1
                Else
                    Dim bS As Long, be As Long
                    codelec = ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value
                    bS = r: be = r
                    Do While be + 1 <= nextRecap - 1
                        If ws.Cells(be + 1, ColNum(TGT_ORDER_COL)).Value <> sOrder Then Exit Do
                        If ws.Cells(be + 1, ColNum(TGT_CODELEC_COL)).Value <> codelec Then Exit Do
                        be = be + 1
                    Loop

                    bCnt = bCnt + 1
                    ReDim Preserve bStarts(1 To bCnt), bEnds(1 To bCnt), bMiss(1 To bCnt), bQmanual(1 To bCnt), bS(1 To bCnt)
                    bStarts(bCnt) = bS: bEnds(bCnt) = be
                    bMiss(bCnt) = ToDbl(ws.Cells(bS, REC_TOTAL_M_COL).Value)
                    If bMiss(bCnt) = 0 Then bMiss(bCnt) = Application.WorksheetFunction.Sum(ws.Range(TGT_QTY_MISS_COL & bS & ":" & TGT_QTY_MISS_COL & be))
                    bS(bCnt) = ToDbl(ws.Cells(bS, REC_CAP_COL).Value)

                    bQmanual(bCnt) = ToDbl(ws.Cells(bS, "Q").Value) ' ? Q saisi � la main sur la 1re ligne du bloc

                    ' vider Q sur les autres lignes du bloc (lisibilit�)
                    If be > bS Then ws.Range("Q" & (bS + 1) & ":Q" & be).ClearContents
                End If
                r = be + 1
            Loop

            ' Surco�t COMMANDE & r�partition prorata O
            Dim scDefault As Double, scOPP As Double, allOPP As Boolean, flux As String
            Dim sumMissOrder As Double, surcostValOrder As Double
            Call GetSurcostConfig(ws, scDefault, scOPP)

            allOPP = True
            For r = rowCur + 1 To nextRecap - 1
                flux = UCase$(Trim$(CStr(ws.Cells(r, ColNum(TGT_FLUX_COL)).Value)))
                If Len(flux) > 0 And flux <> "OPP" Then allOPP = False
            Next r

            sumMissOrder = 0
            For i = 1 To bCnt
                If bMiss(i) > 0 Then sumMissOrder = sumMissOrder + bMiss(i)
            Next i

            If sumMissOrder > 0 Then
                surcostValOrder = IIf(allOPP, scOPP, scDefault)
            Else
                surcostValOrder = 0
            End If

            ' cap + ventilation avec Q_adj = Q_manual + part_surco�t
            Dim sumVOrder As Double: sumVOrder = 0
            For i = 1 To bCnt
                Dim share As Double, Qadj As Double, capFinal As Double
                If (bMiss(i) > 0) And (sumMissOrder > 0) And (surcostValOrder > 0) Then
                    share = surcostValOrder * (bMiss(i) / sumMissOrder)
                Else
                    share = 0
                End If

                Qadj = bQmanual(i) + share
                ws.Cells(bStarts(i), "Q").Value = Qadj
                ws.Cells(bStarts(i), "Q").NumberFormatLocal = "# ##0,00 �"

                capFinal = WorksheetFunction.Min(bS(i), Qadj)
                capFinal = CeilTo(capFinal, 0.01)
                ws.Cells(bStarts(i), REC_BLOCK_CHOSEN_COL).Value = capFinal
                ws.Cells(bStarts(i), REC_BLOCK_CHOSEN_COL).NumberFormatLocal = "# ##0,00 �"

                Dim alloc As Double, lastAllocRow As Long, remainder As Double, missLine As Double
                alloc = 0: lastAllocRow = -1
                If capFinal > 0 And bMiss(i) > 0 Then
                    For rr = bStarts(i) To bEnds(i)
                        missLine = Val(ws.Cells(rr, ColNum(TGT_QTY_MISS_COL)).Value)
                        If missLine > 0 Then
                            ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value = capFinal * (missLine / bMiss(i))
                            alloc = alloc + ws.Cells(rr, ColNum(TGT_PENALTY_COL)).Value
                            lastAllocRow = rr
                        Else
                            ws.Cells(rr, ColNum(TGT_PENALTY_COL)).ClearContents
                        End If
                    Next rr
                    If lastAllocRow <> -1 Then
                        remainder = capFinal - alloc
                        ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value = ws.Cells(lastAllocRow, ColNum(TGT_PENALTY_COL)).Value + remainder
                    End If
                    sumVOrder = sumVOrder + capFinal
                Else
                    ws.Range(TGT_PENALTY_COL & bStarts(i) & ":" & TGT_PENALTY_COL & bEnds(i)).ClearContents
                End If
            Next i

            ' R�cap COMMANDE : Q = S Q_blocs (d�j� ajust�s) ; U = S V (pas d�ajout de surco�t)
            Dim sumQOrder As Double: sumQOrder = 0
            For i = 1 To bCnt
                sumQOrder = sumQOrder + ToDbl(ws.Cells(bStarts(i), "Q").Value)
            Next i
            ws.Cells(rowCur, "Q").Value = CeilTo(sumQOrder, 0.01)
            ws.Cells(rowCur, "Q").NumberFormatLocal = "# ##0,00 �"

            ws.Cells(rowCur, "T").Value = CeilTo(sumVOrder, 0.01)
            ws.Cells(rowCur, "T").NumberFormatLocal = "# ##0,00 �"

            rowCur = nextRecap
        End If
    Loop

    ' Formats �
    ws.Range(REC_BLOCK_CHOSEN_COL & FIRST_DATA_ROW & ":" & REC_BLOCK_CHOSEN_COL & lastRow).NumberFormatLocal = "# ##0,00 �"
    ws.Range(TGT_PENALTY_COL & FIRST_DATA_ROW & ":" & TGT_PENALTY_COL & lastRow).NumberFormatLocal = "# ##0,00 �"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub Refresh_Calculette()
    On Error GoTo Fin
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ? Respecte les Q manuels : recalcule uniquement la ventilation & les r�cap
    Ventiler_Q_Manuel
    
Fin:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' =================== OUTILS ===================

Private Function GetSourceSheet(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetSourceSheet = wb.Worksheets(SOURCE_SHEET_NAME)
    On Error GoTo 0
    If GetSourceSheet Is Nothing And USE_SECOND_SHEET_IF_MISSING Then
        If wb.Worksheets.Count >= 2 Then Set GetSourceSheet = wb.Worksheets(2)
    End If
End Function

' "AB" -> 28
Private Function ColNum(ByVal colLetter As String) As Long
    ColNum = Range(UCase$(colLetter) & "1").Column
End Function

Private Function SameKey(a As Variant, b As Variant) As Boolean
    SameKey = (NormalizeKey(a) = NormalizeKey(b))
End Function

Private Function NormalizeKey(v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If IsDigitsOnly(s) Then
        ' retire les z�ros de t�te (conserve "0")
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If
    NormalizeKey = UCase$(s)
End Function

Private Function IsDigitsOnly(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then IsDigitsOnly = False: Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsDigitsOnly = False: Exit Function
        End If
    Next
    IsDigitsOnly = True
End Function

Private Function IsRecapRow(ws As Worksheet, ByVal r As Long) As Boolean
    Dim txt As String
    txt = CStr(ws.Cells(r, ColNum(TGT_CODELEC_COL)).Value) ' colonne E = libell� codelec / r�cap
    IsRecapRow = (InStr(1, txt, "R�CAP CDE", vbTextCompare) > 0) Or _
                 (InStr(1, txt, "RECAP CDE", vbTextCompare) > 0)
End Function
Private Function CeilTo(ByVal v As Double, ByVal step_ As Double) As Double
    ' Arrondi au sup�rieur au multiple "step_" (ex: 0.1 pour 10 centimes)
    CeilTo = step_ * Application.WorksheetFunction.RoundUp(v / step_, 0)
End Function
Private Function ToDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = 0
    End If
End Function
' Lis un nombre depuis une cellule (supporte �28 ��, espaces ins�cables, virgules, etc.)
Private Function ReadEuroNumber(ByVal rng As Range, ByVal fallback As Double) As Double
    Dim s As String
    s = CStr(rng.Value)
    s = Replace(s, "�", "")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, " ", "")
    s = Replace(s, ",", Application.International(xlDecimalSeparator))
    If IsNumeric(s) Then ReadEuroNumber = CDbl(s) Else ReadEuroNumber = fallback
End Function
' R�cup�re les montants de surco�t (cellule si dispo, sinon constantes)
Private Sub GetSurcostConfig(ByVal ws As Worksheet, ByRef scDefault As Double, ByRef scOPP As Double)
    scDefault = SURCOST_DEFAULT
    scOPP = SURCOST_OPP
    On Error Resume Next
    If Len(SURCOST_DEFAULT_CELL) > 0 Then scDefault = ReadEuroNumber(ws.Range(SURCOST_DEFAULT_CELL), SURCOST_DEFAULT)
    If Len(SURCOST_OPP_CELL) > 0 Then scOPP = ReadEuroNumber(ws.Range(SURCOST_OPP_CELL), SURCOST_OPP)
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    Application.EnableEvents = True
End Sub
