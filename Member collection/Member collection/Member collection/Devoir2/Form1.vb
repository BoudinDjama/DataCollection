'******************************************************************
''**********************************************************************************************************************************************************************************
'  * Author    :  Ahmed, Boudin
'  * Date      :  2018-10-24
'  * Purpose   :  Devoir 2
'  * Tectonics :  ...........
'**********************************************************************************************************************************************************************************
'FORMULAIRE POUR LE DEUXIEME DEVOIR

Public Class Form1


    '======================================DECLARATION===================================
    '-----------------------------------------
    'ou se retrouve la liste de ...
    Dim Provin = "..\..\..\BB_Provinces.txt"
    Dim Lang = "..\..\..\BB_Langues.txt"
    Dim Type = "..\..\..\BB_TypesMembres.txt"
    Dim Membre = "..\..\..\BB_Membres.txt"
    '-----------------------------------------

    '--------------------------------
    'Liseur des texts
    Dim ProvinStringliseur As String
    Dim LanguesStringliseur As String
    Dim TypesStringliseur As String
    Dim MembresStringliseur As String
    '--------------------------------


    '-----------------------------------
    'Enregistrements et mise a jour
    Public enreg As String
    Dim ficher As New List(Of String)
    Dim ficFinal As Array
    Dim save As Boolean = False
    Dim UpdateAJour As Boolean = False
    '-----------------------------------


    '------------------------------------------------------
    'Structures
    Public Structure Membres
        Dim NoMembre As String
        Dim TypeMembre As String
        Dim LangCode As String
        Dim MembreNom As String
        Dim MembrePrenom As String
        Dim MembreAdresse As String
        Dim MembreVille As String
        Dim ProvCode As String
        Dim MembreCodePostal As String
        Dim MembreNoTel As String
        Dim MembreEMail As String
    End Structure
    Public Structure Provinces
        Dim ProvCode As String
        Dim ProvDesc As String
    End Structure
    Public Structure Langues
        Dim LangCode As String
        Dim LangDesc As String
    End Structure
    Public Structure TypesMembres
        Dim TypeMembre As String
        Dim TypeMembreDesc As String
    End Structure
    '----------------------------------------------------


    '----------------------------------------------------------------
    Public FicheMembre As New SortedList(Of String, Membres)
    Public FicheProvince As New SortedList(Of String, Provinces)
    Public FicheLangue As New SortedList(Of String, Langues)
    Public FicheTypeMembre As New SortedList(Of String, TypesMembres)

    Public RecMembre As New Membres
    Public RecProvince As New Provinces
    Public RecTypeMembre As New TypesMembres
    Public RecLangue As New Langues
    '----------------------------------------------------------------


    '======================================================================================






    Private Sub Button8_Click_1(sender As System.Object, e As System.EventArgs) Handles Button8.Click


        Dim Unmembre As New Membres
        Dim Key As String = RecMembre.NoMembre
        '-------------------------------------------------
        ' Update les info du fiche membre
        Unmembre.NoMembre = Me.NumeroMem.Text
        Unmembre.TypeMembre = Me.TypeComb.Text
        Unmembre.LangCode = Me.LangComb.Text
        Unmembre.MembreNom = Me.NomMem.Text
        Unmembre.MembrePrenom = Me.PrenomMem.Text
        Unmembre.MembreAdresse = Me.AdrseMem.Text
        Unmembre.MembreVille = Me.VilleMem.Text
        Unmembre.ProvCode = Me.ProvinceComb.Text
        Unmembre.MembreCodePostal = Me.CodePostalComb.Text
        Unmembre.MembreNoTel = Me.NumeroTel.Text
        Unmembre.MembreEMail = Me.AdrseCourriel.Text
        '--------------------------------------------------


        '---------------------------------------------------------------------
        Province_descrip.Text = FicheProvince.Item(ProvinceComb.Text).ProvDesc

        Lang_descrip.Text = FicheLangue.Item(LangComb.Text).LangDesc

        Type_descrip.Text = FicheTypeMembre.Item(TypeComb.Text).TypeMembreDesc
        '------------------------------------------------------------------------


        '--------------------------------
        'Update la fiche
        RecMembre = Unmembre

        FicheMembre.Item(Key) = RecMembre
        UpdateAJour = True
        '---------------------------------
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        'AFFICHER LA PREMIERE PAGE QUAND CELUI-CI EST CLIQUER (<<)
        RecMembre = FicheMembre.First.Value
        Affichage(RecMembre)

    End Sub




    '*************************************************************************************************************
    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim Position As Integer
        Dim Key As String

        'QUAND LE (<) EST CLIQUER, ON PREND UN PAS A L'ARRIERE


        Position = CInt(RecMembre.NoMembre.Substring(1))
        'Si on est pas a la premiere position
        If Position > 0 Then
            Position -= 1
        End If
        Key = "M" & Format(Position, "000000")

        'ON AFFICHE LA PAGE QUI A LA CLE 
        If FicheMembre.ContainsKey(Key) Then
            RecMembre = FicheMembre.Item(Key)
            Affichage(RecMembre)
        End If

    End Sub
    '*************************************************************************************************************




    '******************************************************************************************************************

    Private Sub Button4_Click_1(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim Position As Integer
        Dim Key As String

        'QUAND LE (>) EST CLIQUER, ON PREND UN PAS EN AVANT


        Position = CInt(RecMembre.NoMembre.Substring(1))
        Position += 1
        Key = "M" & Format(Position, "000000")

        ' ON AFFICHE LA PAGE QUI A LA CLE 
        If FicheMembre.ContainsKey(Key) Then
            RecMembre = FicheMembre.Item(Key)
            Affichage(RecMembre)
        End If

    End Sub
    '*****************************************************************************************************************




    '*************************************************************************************************

    Private Sub Button5_Click_1(sender As System.Object, e As System.EventArgs) Handles Button5.Click

        'AFFICHER LA PREMIERE PAGE QUAND CELUI-CI EST CLIQUER (>>)
        RecMembre = FicheMembre.Last.Value
        Affichage(RecMembre)


    End Sub
    '*************************************************************************************************


    '*********************************************************************************************************************

    Private Sub ChargerCollections_Click(sender As System.Object, e As System.EventArgs) Handles ChargerCollections.Click

        '---------------------------------------------------------
        'IF IL Y A RIEN DANS PROVINCE THEN ON LE REMPLIS..
        If ProvinceComb.MaxLength < 2 Then
            lireFichier(Provin)
            Me.ProvinceComb.Items.Clear()
            For Each element In FicheProvince.Keys
                Me.ProvinceComb.Items.Add(element)
            Next element
        End If
        '---------------------------------------------------------

        '--------------------------------------------------------------
        'IF IL Y A RIEN DANS LANGAGUES THEN ON LE REMPLIS..
        If LangComb.MaxLength < 2 Then
            lireFichier(Lang)
            Me.LangComb.Items.Clear()
            For Each element In FicheLangue.Keys
                Me.LangComb.Items.Add(element)
            Next element
        End If
        '--------------------------------------------------------------

        '--------------------------------------------------------------------
        ''IF IL Y A RIEN DANS TYPE THEN ON LE REMPLIS..
        If TypeComb.MaxLength < 2 Then
            lireFichier(Type)
            Me.TypeComb.Items.Clear()
            For Each element In FicheTypeMembre.Keys
                Me.TypeComb.Items.Add(element)
            Next element
        End If
        '--------------------------------------------------------------------

        '-----------------------------------------------------
        'ON PREND LE PREMIER PAR DEFAUT
        lireFichier(Membre)
        RecMembre = FicheMembre.FirstOrDefault.Value
        Affichage(RecMembre)
        '-----------------------------------------------------

    End Sub
    '********************************************************************************************************************

    Function lireFichier(ByRef Traj As String)

        Dim nombre As Integer = 0
        Dim liseurdeFichier As System.IO.StreamReader
        Dim Array As New ArrayList



        '========================================================================
        'Si le /// appartient a la province then
        If Traj = Provin Then
            If System.IO.File.Exists(Traj) Then
                ' On lit le fichier ligne par ligne
                liseurdeFichier = My.Computer.FileSystem.OpenTextFileReader(Traj)
                While Not liseurdeFichier.EndOfStream
                    ProvinStringliseur = liseurdeFichier.ReadLine()
                    methode(ProvinStringliseur)

                    nombre += 1
                End While
                liseurdeFichier.Close()
            End If
        End If
        '=========================================================================





        '==========================================================================
        'Si le /// appartient au languages then
        If Traj = Lang Then
            If System.IO.File.Exists(Traj) Then
                ' On lit le fichier ligne par ligne
                liseurdeFichier = My.Computer.FileSystem.OpenTextFileReader(Traj)
                While Not liseurdeFichier.EndOfStream
                    LanguesStringliseur = liseurdeFichier.ReadLine()
                    methode(LanguesStringliseur)

                    nombre += 1
                End While
                liseurdeFichier.Close()
            End If
        End If
        '===========================================================================






        '========================================================================
        'Si le /// appartient au types then
        If Traj = Type Then
            If System.IO.File.Exists(Traj) Then
                ' On lit le fichier ligne par ligne
                liseurdeFichier = My.Computer.FileSystem.OpenTextFileReader(Traj)
                While Not liseurdeFichier.EndOfStream
                    TypesStringliseur = liseurdeFichier.ReadLine()
                    methode(TypesStringliseur)

                    nombre += 1
                End While
                liseurdeFichier.Close()
            End If
        End If
        '=========================================================================

        '========================================================================
        'Si le /// appartient au membre then
        If Traj = Membre Then
            If System.IO.File.Exists(Traj) Then
                ' On lit le fichier ligne par ligne
                liseurdeFichier = My.Computer.FileSystem.OpenTextFileReader(Traj)
                While Not liseurdeFichier.EndOfStream
                    MembresStringliseur = liseurdeFichier.ReadLine()
                    methode(MembresStringliseur)

                    nombre += 1
                End While
                liseurdeFichier.Close()
            End If
        End If
        '=========================================================================


        Return Array
    End Function

    Public Sub methode(ByRef stringLiseur As String)
        Dim Pages As New ArrayList
        Dim Key As String

        'ON VERIFIE SI LA PROVINCE ET ON LE MET A LA STRUCTURE
        If stringLiseur = ProvinStringliseur Then

            Pages = lISEUR(ProvinStringliseur, 2)


            '------------------------------ 
            Key = Pages(0)
            RecProvince.ProvCode = Pages(0)
            RecProvince.ProvDesc = Pages(1)
            '------------------------------


            '-----------------------------------------------
            If FicheProvince.ContainsKey(Key) Then
                MsgBox("ERREURE, lA CLE EXIST " & CStr(Key))


            Else
                FicheProvince.Add(Key, RecProvince)

            End If
            '------------------------------------------------




        End If



        '--------------------------------------------
        If stringLiseur = LanguesStringliseur Then
            Pages = lISEUR(stringLiseur, 2)
            Key = Pages(0)
            RecLangue.LangCode = Pages(0)
            RecLangue.LangDesc = Pages(1)
            '--------------------------------------

            If FicheLangue.ContainsKey(Key) Then
                MsgBox("ERREURE, lA CLE EXIST " & CStr(Key))
            Else
                FicheLangue.Add(Key, RecLangue)
            End If
        End If


        If stringLiseur = TypesStringliseur Then
            Pages = lISEUR(stringLiseur, 2)
            Key = Pages(0)
            RecTypeMembre.TypeMembre = Pages(0)
            RecTypeMembre.TypeMembreDesc = Pages(1)
            If FicheTypeMembre.ContainsKey(Key) Then
                MsgBox("ERREURE, lA CLE EXIST " & CStr(Key))
            Else
                FicheTypeMembre.Add(Key, RecTypeMembre)
            End If
        End If


        If stringLiseur = MembresStringliseur Then
            Pages = lISEUR(MembresStringliseur, 11)
            ' lES METTRES DANS LA STRCTURE
            Key = Pages(0)
            RecMembre.NoMembre = Pages(0)
            RecMembre.TypeMembre = Pages(1)
            RecMembre.LangCode = Pages(2)
            RecMembre.MembreNom = Pages(3)
            RecMembre.MembrePrenom = Pages(4)
            RecMembre.MembreAdresse = Pages(5)
            RecMembre.MembreVille = Pages(6)
            RecMembre.ProvCode = Pages(7)
            RecMembre.MembreCodePostal = Pages(8)
            RecMembre.MembreNoTel = Pages(9)
            RecMembre.MembreEMail = Pages(10)
            If FicheMembre.ContainsKey(Key) Then
                MsgBox("ERREURE, lA CLE EXIST " & CStr(Key))
            Else
                FicheMembre.Add(Key, RecMembre)
            End If
        End If

    End Sub



    Public Function lISEUR(ByRef Position As String, numero As Integer) As ArrayList
        Dim Array As New ArrayList
        Dim index As Integer = 1
        Dim compt As Integer = 0

        '--------------------------------------------------
        While compt <> numero
            index = InStr(Position, "|")
            If index <> 0 Then
                Array.Add(Position.Substring(0, index - 1))
                Position = Position.Remove(0, index)
            Else
                Array.Add(Position.Substring(0))
            End If
            compt += 1
        End While
        '--------------------------------------------------
        'ON RETOURNE LE ARRAY
        Return Array
    End Function

    Sub Affichage(ByRef RecMembre As Membres)

        '--------------------------------------------------
        'On place celui selectionner
        Me.NumeroMem.Text = RecMembre.NoMembre
        Me.NomMem.Text = RecMembre.MembreNom
        Me.PrenomMem.Text = RecMembre.MembrePrenom
        Me.AdrseMem.Text = RecMembre.MembreAdresse
        Me.VilleMem.Text = RecMembre.MembreVille
        Me.ProvinceComb.Text = RecMembre.ProvCode
        Me.CodePostalComb.Text = RecMembre.MembreCodePostal
        Me.NumeroTel.Text = RecMembre.MembreNoTel
        Me.AdrseCourriel.Text = RecMembre.MembreEMail
        Me.LangComb.Text = RecMembre.LangCode
        Me.TypeComb.Text = RecMembre.TypeMembre
        '---------------------------------------------------



        '----------------------------------------------------------------------
        'la description de la langue, la province et le type sera afficher
        Province_descrip.Text = FicheProvince.Item(ProvinceComb.Text).ProvDesc
        Lang_descrip.Text = FicheLangue.Item(LangComb.Text).LangDesc
        Type_descrip.Text = FicheTypeMembre.Item(TypeComb.Text).TypeMembreDesc
        '----------------------------------------------------------------------

    End Sub

    Private Sub Sauvegarder_Click(sender As System.Object, e As System.EventArgs) Handles Sauvegarder.Click
        save = False
        For Each element In FicheMembre.Keys
            Dim membre As New Membres
            membre = FicheMembre.Item(element)
            ficher.Add(element)
            ficFinal = ficher.ToArray
        Next element
        If UpdateAJour = True Then
            AjouterDansFichierSequentiel()
            save = True
        Else
            MsgBox("rien est modifier")
        End If
        UpdateAJour = False
    End Sub






    Private Sub AjouterDansFichierSequentiel()
        'IF LE FICHIER N'EXISTE PAS THEN, ON LE CREER NOUS MEME
        Dim rechercherfich As Boolean
        Dim Key As String
        Dim nombre = 0


        rechercherfich = My.Computer.FileSystem.FileExists("..\..\..\Sauvegarde_BB_Membres.txt")


        '------------------------------------------------------------------------------------
        If rechercherfich Then
            My.Computer.FileSystem.DeleteFile("..\..\..\Sauvegarde_BB_Membres.txt")

            For Each element In ficFinal
                Key = ficFinal.GetValue(nombre)
                chaine(Key)
                My.Computer.FileSystem.WriteAllText("..\..\..\Sauvegarde_BB_Membres.txt", enreg & vbCrLf, True)
                nombre += 1
            Next element

        Else
            MsgBox("Fichier non existant. Il sera créé.")
            My.Computer.FileSystem.WriteAllText("..\..\..\Sauvegarde_BB_Membres.txt", enreg & vbCrLf, False)
        End If
        '--------------------------------------------------------------------------------------


    End Sub


    Public Sub chaine(index As String)
        RecMembre = FicheMembre.Item(index)
        enreg = RecMembre.NoMembre & "|" &
        RecMembre.TypeMembre & "|" &
        RecMembre.LangCode & "|" &
        RecMembre.MembreNom & "|" &
        RecMembre.MembrePrenom & "|" &
        RecMembre.MembreAdresse & "|" &
        RecMembre.MembreVille & "|" &
        RecMembre.ProvCode & "|" &
        RecMembre.MembreCodePostal & "|" &
        RecMembre.MembreNoTel & "|" &
        RecMembre.MembreEMail
    End Sub
    Private Sub FermerApplication_Click(sender As System.Object, e As System.EventArgs) Handles FermerApplication.Click

        If UpdateAJour = True And save = True Then
            Me.Close()
        ElseIf UpdateAJour = False Then
            Me.Close()
        Else MsgBox("saver le modification")
        End If


    End Sub




End Class
