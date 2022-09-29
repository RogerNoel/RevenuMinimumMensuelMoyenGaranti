Attribute VB_Name = "Module1"
Public CP As Integer
Public Employeur As Integer
Public num_emp As Integer
Public Gestionnaire As Integer
Public nom_trav As Integer
Public pren_trav As Integer
Public num_trav As Integer
Public dt_naiss As Integer
Public ancien1 As Integer
Public Dt_sortie As Integer
Public base_sal As Integer
Public Q_S As Integer
Public dur_trav_eff As Integer
Public dur_trav_ref As Integer
Public Mois As Integer
Public Annee As Integer
Public Salaire As Integer
Public PFA_3520 As Integer
Public prm_an_3290 As Integer
Public sal_moy_fer_3822 As Integer
Public Prest As Integer
Public SalHor As Integer
Public Statut As Integer
Public ancien2 As Integer
Public Age As Integer
Public RevMMMG As Integer
Public TotalPercu As Integer
Public Prorata As Integer
Public RMMMG_pro_rat As Integer
Public Minimum As Integer
Public diff1 As Integer
Public nbre_lignes As Integer
Public sal3000bis As Integer


' Il ne faut plus travailler avec des salaires horaires; à la place, on a cree une colonne "3000bis" qui regroupe tous les codes de 1 à 999: cette colonne remplacera la case salaire qui n'existe pas pour un salaire horaire.
' La premiere operation va copier ce salaire bis et le coller dans la colonne salaire si celle-ci est vide.

Sub salaire3000bis()
        Call compte_lignes
        For i = 2 To nbre_lignes
                If Cells(i, Salaire) = "" Then
                        Cells(i, Salaire).value = Cells(i, sal3000bis).value
                End If
        Next i
End Sub

Sub compte_lignes()
nbre_lignes = Cells.Find(what:="*", searchdirection:=xlPrevious).Row
End Sub
' ************************************************************************
Sub calcProrata()
'       Verifie si on parle de salaire mensuel ou horaire
'               si mensuel
'                       si le salaire 3000 est inferieur ou egal à la base salariale 1050 (salaire mensuel effectif), on calcule en % le prorata salaire/base salariale
'                       sinon le prorata est 1
'                       Ce pourcentage va proratiser le RMMMG que le travailleur promerite: RMMMG * % * Q/S
'               si horaire
'                       le code divise le RMMMG par le nombre moyen d'heures par mois soit 164.67
'                       le code multiplie ce RMMMG moyen par le nombre d'heures effectivement prestees
'                       le code compare les montants

Set adresse = Range("a1:az1").Find("dur_trav_eff")
dur_trav_eff = adresse.column
Set adresse = Range("a1:az1").Find("dur_trav_ref")
dur_trav_ref = adresse.column
Dim interm As Double

        For i = 2 To nbre_lignes
                ' calcul du Q/S
                Cells(i, Q_S).value = Cells(i, dur_trav_eff).value / Cells(i, dur_trav_ref).value
                If Cells(i, Salaire).value <> "" Then
                ' -------------- SALAIRE MENSUEL --------------
                'range("ag" & j & ":" & "ah" & j) = ""
                        If Cells(i, base_sal) <> "" Then
                                If Cells(i, Salaire).value <= Cells(i, base_sal).value Then
                                        Cells(i, Prorata).value = Cells(i, Salaire) / Cells(i, base_sal)
                                Else
                                        Cells(i, Prorata).value = 1
                                End If
                                With Cells(i, RMMMG_pro_rat)
                                        .value = Cells(i, RevMMMG) * Cells(i, Prorata) * Cells(i, Q_S)  '<<-- uniquement le dernier = Q/S --
                                        .NumberFormat = "#0.00"
                                End With
                        Else
                                Cells(i, base_sal).value = ((13 * Cells(i, dur_trav_ref)) / 3 * Cells(i, SalHor))
                                Cells(i, Prorata).value = Cells(i, Prest) / ((13 * Cells(i, dur_trav_ref)) / 3)
                                Cells(i, RMMMG_pro_rat).value = Cells(i, RevMMMG).value * Cells(i, Prorata).value
                        End If
                End If
        Next i
End Sub
' ************************************************************************************************************

Sub determination_montant_RMMMG()
        Dim age_trav As Integer
                Dim tableauBareme As Range
                Dim moisAct
                Set tableauBareme = Worksheets("Baremes").Range("b8:e19")
        For i = 2 To nbre_lignes
                        age_trav = Cells(i, Age).value
                        moisAct = Cells(i, Mois).value
                        
                        If taille_entreprise < 20 Then
                                        If Cells(i, Statut).value = "FIX" Then ' <<--------------------------------------------------------- Statut contrat fixe ------------------------------
                                                        If Cells(i, ancien2).value < 6 Then  ' <<-- Anciennete --
                                                                revenu_min = Application.VLookup(moisAct, tableauBareme, 2, False)
                                                        ElseIf Cells(i, ancien2).value > 12 Then     ' <<-- Anciennete--
                                                                revenu_min = Application.VLookup(moisAct, tableauBareme, 4, False)
                                                        Else
                                                                revenu_min = Application.VLookup(moisAct, tableauBareme, 3, False)
                                                        End If
                                        Else ' si etudiant    <<--------------------------------------------------------- Statut contrat étudiant ------------------------------
                                                If age_trav > 22 Then
                                                        age_trav = 22
                                                End If
                                        Dim tableauEtudJanvToNov As Range
                                        Dim tableauEtudDec As Range
                                        Set tableauEtudJanvToNov = Worksheets("Baremes").Range("b27:f33")
                                        Set tableauEtudDec = Worksheets("Baremes").Range("h27:l33")
                                                        If Cells(i, ancien2).value < 6 Then                  ' <<-- Anciennete --
                                                                revenu_min = Application.VLookup(age_trav, tableauEtudJanvToNov, 3, False)
                                                        ElseIf Cells(i, ancien2).value >= 12 Then     ' <<-- CAnciennete --
                                                                revenu_min = Application.VLookup(age_trav, tableauEtudJanvToNov, 5, False)
                                                        Else ' anciennete intermediaire
                                                                revenu_min = Application.VLookup(age_trav, tableauEtudJanvToNov, 4, False)
                                                        End If
                                        End If
                        Else ' taille >= 20
                                        If Cells(i, Statut).value = "FIX" Then                      ' <<-- Statut --
                                                        If Cells(i, ancien2).value < 6 Then          ' <<-- Anciennete --
                                                                        revenu_min = 1695.25
                                                        ElseIf Cells(i, ancien2).value > 12 Then ' <<-- Anciennete --
                                                                        revenu_min = 1787.2
                                                        Else
                                                                        revenu_min = 1738.41
                                                        End If
                                        Else ' si etudiant
                                                        If Cells(i, ancien2).value < 6 Then          ' <<-- Anciennete --
                                                                        Select Case age_trav
                                                                                        Case 21
                                                                                        revenu_min = 1695.25
                                                                                        Case 20
                                                                                        revenu_min = 1593.54
                                                                                        Case 19
                                                                                        revenu_min = 1491.82
                                                                                        Case 18
                                                                                        revenu_min = 1390.11
                                                                                        Case 17
                                                                                        revenu_min = 1288.39
                                                                                        Case Else ' 16 et moins
                                                                                        revenu_min = 1186.68
                                                                        End Select
                                                        ElseIf Cells(i, ancien2).value > 12 Then     ' <<-- Anciennete --
                                                                        Select Case age_trav
                                                                                        Case 22
                                                                                        revenu_min = 1787.2
                                                                                        Case 21
                                                                                        revenu_min = 1787.2
                                                                                        Case 20
                                                                                        revenu_min = 1679.97
                                                                                        Case 19
                                                                                        revenu_min = 1572.74
                                                                                        Case 18
                                                                                        revenu_min = 1465.5
                                                                                        Case 17
                                                                                        revenu_min = 1358.27
                                                                                        Case Else ' 16 et mois
                                                                                        revenu_min = 1251.04
                                                                        End Select
                                                        Else 'anciennete intermediaire
                                                                        Select Case age_trav
                                                                                        Case 21
                                                                                        revenu_min = 1738.41
                                                                                        Case 20
                                                                                        revenu_min = 1634.11
                                                                                        Case 19
                                                                                        revenu_min = 1529.8
                                                                                        Case 18
                                                                                        revenu_min = 1425.5
                                                                                        Case 17
                                                                                        revenu_min = 1321.19
                                                                                        Case Else ' 16 et moins
                                                                                        revenu_min = 1216.89
                                                                        End Select
                                                        End If
                                        End If
                        End If
                        Cells(i, RevMMMG).value = revenu_min          ' <<-- Rmmmg --
        Next i ' Fin determination RMMMG -------------------------------------------------------
End Sub

' **************************************************************************************************************

Sub determination_rmmmg()
' **************************
' code macro principal qui appelle les autres
' *************************
        Worksheets("Export Prisma").Activate
        Range("a1").Select
        With ActiveWindow
                        .SplitColumn = 0
                        .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True

        Columns("L:L").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("l1").value = "Q_S"

        Set adresse = Range("a1:az1").Find("CP")
        CP = adresse.column
        Set adresse = Range("a1:az1").Find("Employeur")
        Employeur = adresse.column
        Set adresse = Range("a1:az1").Find("num_emp")
        num_emp = adresse.column
        Set adresse = Range("a1:az1").Find("Gestionnaire")
        Gestionnaire = adresse.column
        Set adresse = Range("a1:az1").Find("nom_trav")
        nom_trav = adresse.column
        Set adresse = Range("a1:az1").Find("pren_trav")
        pren_trav = adresse.column
        Set adresse = Range("a1:az1").Find("num_trav")
        num_trav = adresse.column
        Set adresse = Range("a1:az1").Find("dt_naiss")
        dt_naiss = adresse.column
        Set adresse = Range("a1:az1").Find("ancien1")
        ancien1 = adresse.column
        Set adresse = Range("a1:az1").Find("Dt_sortie")
        Dt_sortie = adresse.column
        Set adresse = Range("a1:az1").Find("base_sal")
        base_sal = adresse.column
        Set adresse = Range("a1:az1").Find("Q_S")
        Q_S = adresse.column
        Set adresse = Range("a1:az1").Find("Mois")
        Mois = adresse.column
        Set adresse = Range("a1:az1").Find("Annee")
        Annee = adresse.column
        Set adresse = Range("a1:az1").Find("Salaire")
        Salaire = adresse.column
        Set adresse = Range("a1:az1").Find("PFA_3520")
        PFA_3520 = adresse.column
        Set adresse = Range("a1:az1").Find("prm_an_3290")
        prm_an_3290 = adresse.column
        Set adresse = Range("a1:az1").Find("sal_moy_fer_3822")
        sal_moy_fer_3822 = adresse.column
        Set adresse = Range("a1:az1").Find("Prest")
        Prest = adresse.column
        Set adresse = Range("a1:az1").Find("SalHor")
        SalHor = adresse.column
        Set adresse = Range("a1:az1").Find("sal3000bis")
        sal3000bis = adresse.column
        
        Application.ScreenUpdating = False
        Call compte_lignes
        Worksheets("Export Prisma").Activate

        Dim revenu_min As Double
        Dim taille_entreprise As Integer
        taille_entreprise = 1 ' /!\ Il n'y a actuellement pas d'entreprises de + de 20, donc pas traite
        Dim j As Integer, k As Integer
        
        Call salaire3000bis

'  entetes
        Cells(1, 26).value = "Statut"
        Cells(1, 27).value = "ancien2"
        Cells(1, 28).value = "Age"
        Cells(1, 29).value = "RevMMMG" ' --> determination du RMMMG selon bareme
        Cells(1, 30).value = "TotalPercu"  ' --> somme des salaires perçus en O-P-Q-R
        Cells(1, 31).value = "Prorata"  ' --> rapport entre la base salariale K et le salaire O pour calculer en AD un RMMMG Pro-ratise
        Cells(1, 32).value = "RMMMG_pro_rat" ' --> voir ligne precedente
        Cells(1, 33).value = "Minimum"  ' --> minimum à percevoir sur l'annee: somme de tous les RMMMG pro-ratisee (AD)
        Cells(1, 34).value = "diff1"

        Worksheets("Export Prisma").Activate

        '******************************************************************************************
        ' Determination de  la fin de la serie de lignes d'un seul travailleur pour y faire la somme de ses paies et la mise en forme visuelle
        '******************************************************************************************
        Set adresse = Range("a1:az1").Find("TotalPercu")
        TotalPercu = adresse.column

        Dim debut, fin As Long
        debut = 2

        For i = 2 To nbre_lignes
                Range("A" & i & ":AM" & i).font.italic = True
                k = i + 1
                ' ... comparaison ligne par ligne des num d'employeur et de travailleur
                If (Cells(i, num_trav).value <> Cells(k, num_trav).value) Or (Cells(i, num_trav).value = Cells(k, num_trav).value And Cells(i, num_emp).value <> Cells(k, num_emp).value) Then
                ' trace d'une ligne en gras à la fin de chaque travailleur
                        Range("A" & i & ":AM" & i).Select
                        With Selection.Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .ColorIndex = xlAutomatic
                                .TintAndShade = 0
                                .Weight = xlMedium
                        End With
                        Selection.Borders(xlEdgeRight).LineStyle = xlNone
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                                        ' ************************************
                                        ' calcul du total des salaires perçus: salaire (code 3000) + PFA (code 3520) + Prime annuelle (code 3290) + salaire moyen jours feries (code 3822)
                                        ' ***********************************
                        Dim total As Double
                        total = 0
                        Dim colonne1, colonne2 As Long
                        colonne1 = Salaire
                        colonne2 = sal_moy_fer_3822
                        fin = i
                        For colonnes = colonne1 To colonne2
                                For lignes = debut To fin
                                                If IsNumeric(Cells(lignes, colonnes).value) = True Then
                                                        total = total + Cells(lignes, colonnes).value
                                                End If
                                Next lignes
                        Next colonnes
                        Cells(i, TotalPercu).value = total
                        total = 0
                        debut = i + 1
                        colonnes = colonne1
                End If
        Next i

        Set adresse = Range("a1:az1").Find("Prorata")
        Prorata = adresse.column
        Set adresse = Range("a1:az1").Find("RMMMG_pro_rat")
        RMMMG_pro_rat = adresse.column
        Set adresse = Range("a1:az1").Find("RevMMMG")
        RevMMMG = adresse.column
        Columns(Prorata).Select
        Selection.style = "Percent"

                        '******************************************************************************************
                        ' Verification - du statut du travailleur; soit etudiant soit "fixe"
                        '                   - de l'anciennete du travailleur en mois
                        '                   - de l'age du travailleur en annees
                        '******************************************************************************************
        Set adresse = Range("a1:az1").Find("Statut")
        Statut = adresse.column
        Set adresse = Range("a1:az1").Find("Age")
        Age = adresse.column
        Set adresse = Range("a1:az1").Find("ancien2")
        ancien2 = adresse.column
        
        For i = 2 To nbre_lignes
                If Cells(i, num_trav) >= 40000 And Cells(i, num_trav) < 50000 Then
                        Cells(i, Statut).value = "STU"
                Else
                        Cells(i, Statut).value = "FIX"
                End If

                Dim date1 As Date
                Dim date2 As Date
                date1 = Cells(i, ancien1).value
                date2 = Cells(i, Mois).value
                
                If year(date1) = year(date2) And month(date1) = month(date2) Then
                        Cells(i, ancien2) = 0
                Else
                        Cells(i, ancien2).value = DateDiff("m", date1, date2)
                End If

                Cells(i, Age).Select                                 ' <<-- Age --
                Dim naiss2 As Date
                naiss2 = Cells(i, dt_naiss).value
                Dim anCourant As Date
                anCourant = Cells(i, Mois).value
                Dim days As Integer
                days = anCourant - naiss2
                Cells(i, Age).value = Format(days, "yy")
        Next i

        '******************************************************************************************
        ' Determination du montant du RMMMG
        '******************************************************************************************
        Call determination_montant_RMMMG
        Call calcProrata

'******************************************************************************************
' Ces lignes calculent la somme des montants de RMMMG proratises pour comparer avec le salaire reellement perçu
'******************************************************************************************
        Set adresse = Range("a1:az1").Find("Minimum")
        Minimum = adresse.column
        Dim ligne_bas As Integer
        cellule_bas = 2
        Dim cellule_haut As Integer
        cellule_haut = 2
        
        For i = 2 To nbre_lignes
                If Cells(i, TotalPercu) <> "" Then
                        cellule_haut = i
                        Cells(cellule_haut, Minimum).Formula = "=SUM(AF" & cellule_bas & ":AF" & cellule_haut & ")"
                        cellule_bas = i + 1
                        cellule_haut = 0
                End If
        Next i

        '******************************************************************************************
        ' Ces lignes calculent la difference entre minimum et salaire reellement perçu
        '******************************************************************************************
        Set adresse = Range("a1:az1").Find("diff1")
        diff1 = adresse.column
        
        For i = 2 To nbre_lignes
                If Cells(i, TotalPercu) <> "" Then
                        'Cells(i, diff1).FormulaR1C1 = "=RC[-4]-RC[-1]"
                        Cells(i, diff1).value = Cells(i, TotalPercu) - Cells(i, Minimum)
                        Cells(i, diff1).style = "currency"
                        If Cells(i, diff1) < 0 Then
                                Cells(i, diff1).Interior.ColorIndex = 44
                        Else
                                Cells(i, diff1).Interior.ColorIndex = 50
                        End If
                End If
        Next i

        Range("a:am").Select
        Selection.font.size = 9
        Columns("A:Ah").autofit
        Columns("h:j").Hidden = True
        Columns("n:n").Hidden = True
        Columns("p:ab").Hidden = True
        Range("a2").Select
        msgbox ("Calculs terminés.")

End Sub
