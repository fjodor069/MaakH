'-----------------------------------------------------------------------
' ConsoleMaakH.exe  / Module1.vb
'
' programma voor maken van boorprogramma's
'
'  voor Zayer / Unisign /Sacem
'
' Windows Console Application : Target Platform x86 (32-bit)
'
' source files  : Module1.vb / Holeclass.vb / ValueComparer.vb
' wordt aangeroepen door : Autocad Autolisp routine (MaakH.lsp) 
'
' invoerfiles: EXPORT.ASC  / BCYCLUS.TXT
' uitvoerfile: XXXX.H
'
' 29-11-2007 originele listing MAAKH.BAS in GWBasic/QBasic          FW
'            begonnen in Visual Studio 2005                         RvS
' 26-3-2012 omgezet naar Visual Studio 2010                         RvS
' 3-4-2012  lezen van .asc file gewijzigd , met textfieldparser     RvS
' 7-4-2012  begin met console app                                   RvS
' 8-4-2012  invoer routines                                         RvS
' 22-4-2012 uitvoer translatecycle                                  RvS
' 3-5-2012  2-spil boren uitgewerkt                                 RvS
'
' 7-8-2012 testen met Frans W.
'           uitvoer in 3 decimalen  
'           regelnummers toegevoegd
'           quote karakters verwijderd (bcyclus)
'           delta>270 toegevoegd voor 2 spillen
' 8-8-2012  Translatecycle uitvoer is nu goed
' 9-8-2012  invoer van meerdere bewerkingen mogelijk
' 10-8-2012 extra spaties toegevoegd; laatste regel weggehaald
' 13-8-2012 default waarde voor toerental                           RvS
'           extra uitvoer naar .TXT file
' 16-8-2012 toegevoegd: unisign gedeelte
' 2-10-2012 toegevoegd: sacem gedeelte  
' 4-10-2012 sorteervolgorde aangepast voor sacem
' 22-10-2012 MO gewijzigd in M0
' ----------------------------------------------------------------------
Option Explicit On
Option Strict On
Imports System
Imports System.IO
Imports System.Math
Imports System.Collections
Imports System.Globalization


Module Module1


    Dim astring As New String(Chr(61), 66)      'een string met karakter '='

    Enum machineType                            'machine typen:
        mZayer = 0                              'boorbank met 2 spillen
        mUnisign                                'boorbank met 1 spil
        mSacem                                  'kotterbank
    End Enum

    Dim mType As machineType


    'netwerk locatie waar boorprogramma files worden weggeschreven 
    Dim sPath1 As String() = {"S:\26_824W\BOORMACH\Zayer\",
                                 "S:\26_824W\BOORMACH\Unisign\",
                                 "S:\26_824W\BOORMACH\Sacem\"}

    Dim apparaatName As String() = {"ZAYER", "UNISIGN", "SACEM"}

    Const sPath2 As String = "D:\"                     'het path waar de invoerfile export.asc staat
    Const sPath3 As String = "D:\Eigen\voorbeeld boren\"
    Const sCyclus As String = "BCYCLUS.TXT"

    Const minDist As Double = 270.0         'minimale afstand tussen twee spillen van Zayer
    Const defToerental As Integer = 1000    'default toerental

    Dim cyclus As New ArrayList

    Dim sSpil As String() = {"Alle gaten met de Y-spil",
                             "Alle gaten met de V-spil",
                             "met beide spillen, voorkeur voor V-spil bij enkele gaten",
                             "met beide spillen, voorkeur voor Y-spil bij enkele gaten"}
    Dim iCycle As Integer

    Dim sProg As String         'naam boorprogramma
    Dim sFileName As String     'bestandsnaam zonder extensie

    Dim sTest As String = "Export.asc"      'het invoerbestand van autocad


    'variabelen voor opslaan en sorteren 
    Dim ssg, pbg, pg, tg1, tg2, bcg As New List(Of Holes)()

    'de uitvoer file
    Dim sw As New StringWriter              'het boorprogramma
    Dim st As New StringWriter              'de text file


    Dim ci As CultureInfo                   'numeric format info


    '-----main routine; programma begint hier
    Sub Main()

        'begin met informatie regels
        Console.WriteLine(astring)
        Console.WriteLine("| Programma : {0,30}                     | ", My.Application.Info.AssemblyName)
        Console.WriteLine("| versie    : {0,30}                     | ", My.Application.Info.Version.ToString)
        Console.WriteLine("| datum     : {0,30}                     | ", "22/10/2012")
        Console.WriteLine("| Company   : {0,30}                     | ", My.Application.Info.CompanyName)
        Console.WriteLine("|             {0,30}                     | ", Space(20))
        Console.WriteLine(astring)
        Console.WriteLine()

        mType = Kies_Apparaat()             'kies voor welk boor machine type een programma moet worden gemaakt

        Console.WriteLine("|             {0,30}           |", "De boorprogramma's worden opgeslagen in:")
        Console.WriteLine("|             {0,30}                     | ", Space(20))
        Console.WriteLine("|             {0,30}                     | ", sPath1(mType))
        Console.WriteLine()

        'lees beschikbare cycli voor machine


        If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then
            Read_cyclus(sPath1(mType), sCyclus)
        End If

        'opens export.asc and read input 
        '---Read_from_file(sPath2, sTest)
        Read_from_file(sPath2, sTest)


        '-------------------------------------
        Console.Write("Naam voor het boorprogramma (komt in 1e regel te staan)   : ")
        sProg = Console.ReadLine()

        Console.Write("Bestandsnaam voor het programma (zonder extensie)  : ")
        'hier nog controleren of de naam goed is ingevoerd.
        sFileName = Console.ReadLine()


        '-------------------
        Init_file()

        Console.WriteLine(astring)
        '-------------------------------
        Console.WriteLine("aantal steunstang gaten  : {0}  ", ssg.Count)
        Console.WriteLine("aantal pakketbout gaten  : {0}  ", pbg.Count)
        Console.WriteLine("aantal pijpgaten         : {0}  ", pg.Count)
        Console.WriteLine("aantal tapgaten (type 1) : {0}  ", tg1.Count)
        Console.WriteLine("aantal tapgaten (type 2) : {0}  ", tg2.Count)
        Console.WriteLine("aantal boutcirkelgaten   : {0}  ", bcg.Count)
        Console.WriteLine()

        Console.WriteLine("Druk op een toets....")
        While Not Console.KeyAvailable()
        End While
        Console.ReadKey()

        Console.Clear()             'clear screen

        iCycle = 0

        If ssg.Count > 0 Then Bewerking("Steunstanggaten   ", ssg)
        If pbg.Count > 0 Then Bewerking("Pakketboutgaten   ", pbg)
        If pg.Count > 0 Then Bewerking("Pijpgaten         ", pg)
        If tg1.Count > 0 Then Bewerking("Tapgaten          ", tg1)
        If tg2.Count > 0 Then Bewerking("Tapgaten 2e serie ", tg2)
        If bcg.Count > 0 Then Bewerking("Boutcirkelgaten   ", bcg)

        '-------------------------------
        Console.WriteLine(astring)

        'schrijf alle regels naar het boorprogramma
        'Finish_file(sPath3, sFileName)
        Finish_file(sPath1(mType), sFileName)


        'end of the program, wait for user key press
        Console.WriteLine("Druk op een toets voor einde ....")
        While Not Console.KeyAvailable()
        End While
        Console.ReadKey()

        'end of the program
        End


    End Sub

    '
    ' subroutine 1 ; stel opties in voor een bewerkingsgroep
    ' invoer van :
    ' - commentaarregel
    ' - cyclus voor de bewerking (een van de standaard cycli)
    ' - met 1 of 2 spillen boren
    ' - evt nog een tweede bewerking
    '
    ' uitvoer naar : boorprogramma (stringwriter)
    '
    Sub Bewerking(ByVal sName As String, ByVal alst As List(Of Holes))


        Dim sComment As String
        Dim iKeuze As Integer

        Dim iToerental As Integer

        Dim iSpil As Integer
        Dim bFinished As Boolean = False        ' als evt tweede bewerking gereed is
        Dim cki As ConsoleKeyInfo
        Dim ch As Char


        Console.WriteLine(sName & " : {0} gaten ", alst.Count)
        Console.Write("Commentaarregel (bijv. gaten dia 19 )  : ")
        sComment = Console.ReadLine()

        '---keuze uit beschikbare cycli maken
        'je kan meer dan een cyclus kiezen (bij Zayer or Unisign)
        Do
            If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then

                iKeuze = Kies_Cyclus()
            Else
                iKeuze = 1

            End If

            'commando M0 na elke bewerking in een cyclus aan het einde van de cyclus
            If iCycle > 0 Then
                If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then


                    sw.WriteLine("M0")
                Else
                    sw.WriteLine("STOP M0")
                End If
            Else
                iCycle = iCycle + 1
            End If

            sw.WriteLine("* - " & sComment)                 'schrijf comment in boorprogramma
            st.WriteLine()                                  'schrijf comment in textfile
            st.WriteLine("***** " & sComment & " *****:")


            '--keuze voor toerental
            iToerental = Kies_toerental()
            sw.WriteLine("TOOL CALL 0 Z S" & iToerental.ToString)

            If mType = machineType.mUnisign Then            'extra commando voor unisign
                sw.WriteLine("FN 19: PLC =+")
            End If
            '--keuze voor type spil
            If mType = machineType.mZayer Then

                iSpil = Kies_Spil()                         'zayer heeft keuze mogelijkheid voor 1 of 2 spillen
            Else
                iSpil = 1                                   'unisign en sacem altijd 1 spil
            End If

            ' schrijf een comment in de text file
            If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then
                st.WriteLine("- Cyclus: {0} / {1} ", cyclus(iKeuze - 1), sSpil(iSpil - 1))
            End If

            '--coordinaten sorteren, controleren en vertalen 
            Sorteergaten(alst)
            Controleergaten(alst)
            'maken van de programma regels voor boorprogramma
            'en commentaar naar textfile van rijen en aantallen gaten
            TranslateCycle(iSpil, alst, iKeuze)

            Console.WriteLine()

            '---vraag of er nog een bewerking nodig is (cyclus)

            Console.Write("Nog een bewerking op deze groep (J/N) ? : ")

            Do
                cki = Console.ReadKey(True)        'lees toets zonder echo

                ch = cki.KeyChar

            Loop Until (cki.Key = ConsoleKey.J Or cki.Key = ConsoleKey.N)
            If cki.Key = ConsoleKey.N Then bFinished = True

            Console.WriteLine()


            '------
        Loop While Not bFinished

        'Console.WriteLine("Druk op een toets....")
        'While Not Console.KeyAvailable()
        'End While
        'Console.ReadKey()
        'clear screen
        Console.Clear()


    End Sub
    '
    ' kies voor welk apparaat type een boorprogramma wordt gemaakt
    '
    ' uitvoer is de keuze als enum type
    Function Kies_Apparaat() As machineType


        Dim iKey As Integer
        Dim cki As ConsoleKeyInfo
        Dim ch As Char


        Do
            Console.Write("Kies een apparaat (1=Zayer, 2=Unisign, 3=Sacem) :  ")
            cki = Console.ReadKey(True)
            ch = cki.KeyChar

            If IsNumeric(ch) Then
                iKey = Int32.Parse(ch.ToString)
            Else
                Console.WriteLine("geef een numerieke waarde : {0} ", ch.ToString)

            End If
        Loop Until (iKey >= 1) And (iKey <= 3)       'keuze tussen 1 en 3


        Select Case iKey

            Case 1
                Kies_Apparaat = machineType.mZayer


            Case 2
                Kies_Apparaat = machineType.mUnisign

            Case 3
                Kies_Apparaat = machineType.mSacem

            Case Else
                Kies_Apparaat = machineType.mZayer


        End Select
        'Console.WriteLine(" keuze = {0} ---> {1} ", cki.Key.ToString, iKey)

        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Apparaat {0} {1}", iKey, apparaatName(iKey - 1))
        Console.WriteLine()

        Console.ForegroundColor = ConsoleColor.Gray

    End Function
    '
    ' geef aan welke cycli er gekozen kunnen worden 
    ' uit de arraylist cyclus()
    ' en geef de gekozen waarde terug als integer
    ' alleen voor zayer en unisign
    Function Kies_Cyclus() As Integer


        Dim s As String
        Dim ikey As Integer

        '---keuze uit beschikbare cycli maken
        '   je kan meer dan 1 cyclus kiezen per bewerkingsgroep

        'geef aan welke cycli gekozen kunnen worden
        '(staan in textfile BCYCLUS.TXT)
        Console.WriteLine("Beschikbare cycli: ")
        '-------------
        'For Each str As String In cyclus
        '    Console.WriteLine(str)
        'Next
        '-------------
        For i = 1 To cyclus.Count - 1
            Console.WriteLine("{0} ) {1} ", i, cyclus(i - 1))

        Next
        Console.WriteLine()

        ' read a numeric string from the console between 1 and count-1
        ' convert the string to a valid integer
        ' n.b. het kan ook een cijfer zijn met meer dan 1 karakter (digit)
        '
        ikey = 0
        Console.Write("Kies cyclus : ")
        Do

            s = Console.ReadLine()
            If Int32.TryParse(s, ikey) Then

                'Console.WriteLine(" keuze = {0} --> {1} ", s, ikey)
            Else

                Console.WriteLine(" geef een numerieke waarde : {0} ", s)
            End If

            'ga net zolang door tot je een juiste keuze hebt
        Loop Until (ikey > 0) And (ikey < cyclus.Count)

        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Keuze = {0} ", cyclus(ikey - 1).ToString)
        Console.ForegroundColor = ConsoleColor.Gray


        'Console.WriteLine("Druk op een toets....")
        'While Not Console.KeyAvailable()
        'End While
        'Console.ReadKey()
        ''clear screen
        'Console.Clear()


        Kies_Cyclus = ikey

    End Function

    Function Kies_toerental() As Integer

        Dim s As String
        Dim i As Integer
        Do
            Console.Write("Geef toerental (default = {0} rpm) : ", defToerental)
            s = Console.ReadLine()

            If Int32.TryParse(s, i) Then

                'Console.WriteLine(" keuze = {0} --> {1} ", s, ikey)
                If i < 1000 Then
                    Console.WriteLine(" geef waarde >= 1000 : {0}", s)
                End If
                If i > 3500 Then
                    Console.WriteLine(" geef waarde <= 3500 : {0}", s)
                End If

            Else
                'op enter gedrukt of een andere toets
                i = defToerental
                ' Console.WriteLine(" geef een numerieke waarde : {0} ", s)
            End If
            'ga door totdat de invoer goed is
        Loop Until (i >= 1000) And (i <= 3500)

        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Invoer = {0} rpm ", i.ToString)
        Console.ForegroundColor = ConsoleColor.Gray

        Kies_toerental = i


    End Function
    '
    ' kies welke boorspillen er gebruikt worden  
    '
    Function Kies_Spil() As Integer


        Dim i As Integer
        Dim ikey As Integer
        Dim cki As ConsoleKeyInfo
        Dim ch As Char

        i = 0
        For Each s As String In sSpil

            Console.WriteLine("{0}) {1} ", i + 1, s)
            i = i + 1

        Next
        Do

            Console.Write("Uw keuze : ")

            cki = Console.ReadKey(True)        'lees toets zonder echo

            ch = cki.KeyChar

            If IsNumeric(ch) Then

                ikey = Int32.Parse(ch.ToString)

                ' Console.WriteLine(" keuze = {0} --> {1} ", cki.Key.ToString, ikey)


            Else

                Console.WriteLine(" geef een numerieke waarde : {0} ", ch.ToString)
            End If


            'ga door totdat de invoer goed is (keuze iKey tussen 1 en 4)
        Loop Until (ikey >= 1) And (ikey < i + 1)

        Console.ForegroundColor = ConsoleColor.DarkGreen

        Console.WriteLine("Invoer = {0} {1} ", ikey, sSpil(ikey - 1))

        Console.ForegroundColor = ConsoleColor.Gray




        Kies_Spil = ikey

    End Function
    '-------------------------------------------------------------
    'schrijf de eerste regels van het boorprogramma
    '-------------------------------------------------------------
    Sub Init_file()


        sw.WriteLine("BEGIN PGM " & sProg & " MM")
        '--extra commentaar regel 
        sw.WriteLine("; PROGRAMMA VOOR {0}", apparaatName(mType))

        If mType = machineType.mUnisign Then    'extra regels voor unisign programma

            sw.WriteLine("CYCL DEF 7.0 NULPUNT")
            sw.WriteLine("CYCL DEF 7.1 X+0")
            sw.WriteLine("CYCL DEF 7.2 Y+0")

        End If


        'schrijf info naar text file

        st.WriteLine("PROGRAMMANAAM: " & sFileName)


    End Sub
    '-------------------------------------------------------------
    ' schrijf laatste regel boorprogramma
    ' maak nieuw sFileName.H bestand aan op schijf en schrijf alles weg 
    ' maak ook een sFileName.TXT bestand aan 
    '-------------------------------------------------------------
    Sub Finish_file(ByVal sPath As String, ByVal aName As String)

        If mType = machineType.mSacem Then
            sw.WriteLine("STOP M0")
        End If
        sw.WriteLine("END PGM " & sProg & " MM")

        ' add line numbers beginning with zero 
        Dim strReader As New StringReader(sw.ToString)
        Dim strWriter As New StringWriter()

        Dim aline As String
        Dim i As Integer
        i = 0
        Do

            aline = strReader.ReadLine()
            If (Not aline Is Nothing) Then strWriter.WriteLine(i & " " & aline)
            i = i + 1


        Loop Until aline Is Nothing


        'finally, write all output to the text file (het boorprogramma)
        Try
            Using outfile As New StreamWriter(sPath & aName & ".H")
                ' outfile.Write(sw.ToString())
                outfile.Write(strWriter.ToString)
                outfile.Close()
            End Using

        Catch ex As Exception

            Console.WriteLine("Finish_File: " & ex.Message)

        End Try

        'also, write all comment output to the text file (.TXT)
        Try
            Using outfile As New StreamWriter(sPath & aName & ".TXT")
                outfile.Write(st.ToString)
                outfile.Close()
            End Using

        Catch ex As Exception

            Console.WriteLine("Finish_File: " & ex.Message)

        End Try
        sw.Close()
        st.Close()

    End Sub
    'lees beschikbare cycli voor boormachine uit een text file BCYCLUS.TXT
    'deze staat in de folder sPath1
    'deze routine slechts eenmaal uitvoeren ; alleen voor Zayer en Unisign
    Sub Read_cyclus(ByVal sPath As String, ByVal sName As String)



        If Not File.Exists(sPath & sName) Then
            Console.WriteLine("Read_cyclus : {0} does not exist.", sName)
            'MessageBox.Show("File does not exist  ")
            Return
        End If

        Try

            Using sr As StreamReader = New StreamReader(sPath & sName)
                Dim line As String

                Do

                    line = sr.ReadLine()

                    ' remove unnecessary quote characters
                    line = Replace(line, """", String.Empty)

                    cyclus.Add(line)

                Loop Until line Is Nothing
                sr.Close()
            End Using

        Catch ex As Exception
            Console.WriteLine("Read_cyclus " & ex.Message)

        End Try


    End Sub


    'lees een *.asc bestand ; dit is het bestand wat gemaakt wordt door AutoLisp
    'bestand openen als comma delimited text file 
    'en alle coordinaten lezen
    'per layer aan de juiste collectie toevoegen
    'het formaat van de regels is < layernaam >, <x> , <y> 
    'de coordinaten worden gelezen als double floating point getal
    '
    Sub Read_from_file(ByVal sPath As String, ByVal sName As String)

        If Not File.Exists(sPath & sName) Then
            Console.WriteLine("Read_from_file: {0} does not exist.", sName)
            'MessageBox.Show("File does not exist  ")
            Return
        End If


        Dim delimiter As String = ","
        '-----------------
        Dim x, y As Double

        Dim gat As Holes

        Dim iTeller As Short


        Try
            'create an instance of textfieldparser to read from a comma delimited file
            Using sr As New FileIO.TextFieldParser(sPath & sName)
                sr.TextFieldType = FileIO.FieldType.Delimited
                sr.SetDelimiters(delimiter)

                Dim currentRow As String()

                While Not sr.EndOfData

                    Try
                        'read all fields on the current line and return them in an array of strings
                        currentRow = sr.ReadFields()

                        Dim ls As String

                        ls = currentRow(0)                  'de naam van de layer

                        'de coordinaten inlezen ; omzetten naar een culture om de decimale punt te lezen
                        Try
                            ci = CultureInfo.CreateSpecificCulture("en-US")
                            x = Double.Parse(currentRow(1), ci)
                            y = Double.Parse(currentRow(2), ci)

                        Catch ex As FormatException
                            Console.WriteLine("{0}: Unable to parse '{1}' '{2}'.",
                                              ci.Name, currentRow(1), currentRow(2))
                        End Try


                        '-------
                        'steeds een nieuw object maken voor elk nieuw gat
                        gat = New Holes(x, y)

                        'afronden op 1 cijfer
                        x = Fix(x) + CShort((x - Fix(x)) * 1000) / 1000
                        y = Fix(y) + CShort((y - Fix(y)) * 1000) / 1000

                        gat.x = x
                        gat.y = y



                        Select Case ls
                            Case "SSG"

                                ssg.Add(gat)

                            Case "PBG"

                                pbg.Add(gat)

                            Case "PG"

                                pg.Add(gat)

                            Case "TG1"

                                tg1.Add(gat)

                            Case "TG2"

                                tg2.Add(gat)

                            Case "BCG"

                                bcg.Add(gat)

                        End Select

                        iTeller = CShort(iTeller + 1)   'telt het totale aantal
                        '-----------------------------------


                    Catch ex As Exception
                        'let the user know what went wrong
                        Console.WriteLine("Read from file" & ex.Message)
                    End Try

                End While


                sr.Close()
            End Using



        Catch ex As Exception
            'let the user know what went wrong
            Console.WriteLine(ex.Message)

        End Try

    End Sub


    Sub Sorteergaten(ByVal alst As List(Of Holes))
        'sorteer de objecten op coordinaten van laag naar hoog
        'dus van -x,-y naar +x,+y eerst per x rij dan per y kolom
        'Opm. de arraylist wordt gesorteerd, maar niet de collection !!
        'Dim alst As ArrayList = New ArrayList(boorgroep)

        If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then


            Dim defComp As Comparer(Of Holes) = Comparer(Of Holes).Default

            alst.Sort()

        Else


            alst.Sort(New HolesyFirst())


        End If


    End Sub
    Sub Controleergaten(ByVal al As List(Of Holes))
        '
        'controleer of onderlinge afstand in y richting niet te klein is
        Dim diffx, diffy As Double

        'Dim al As ArrayList = ArrayList.Adapter(boorgroep)

        Dim i As Integer
        For i = 1 To al.Count - 1
            Dim boorgat1 As Holes = DirectCast(al(i), Holes)
            Dim boorgat2 As Holes = DirectCast(al(i - 1), Holes)
            diffy = System.Math.Abs(boorgat1.y - boorgat2.y)
            diffx = System.Math.Abs(boorgat1.x - boorgat2.x)

            If diffx < 0.1 Then
                If diffy < 0.1 Then
                    Console.WriteLine("!! Dubbel coordinaat !!  x : {0}  y : {1} ", boorgat1.x, boorgat1.y)
                    'MsgBox("dubbel coordinaat !! :   " & vbCrLf & _
                    '        "diffy :  " & CStr(diffy) & vbCrLf & _
                    '        al.Item(i).sInfo & vbCrLf & _
                    '        al.Item(i - 1).sInfo)
                    Exit For
                End If
            End If
        Next

    End Sub
    Sub TranslateCycle(ByVal q As Integer, ByVal al As List(Of Holes), ByVal iCyclusKeuze As Integer)
        'maak het boorprogramma voor de gesorteerde boorgroep
        'in de arraylist
        ' schrijf programma regels als strings naar object sw (stringwriter)
        ' deze wordt ineen keer toegevoegd aan boorprogramma 
        ' schrijf ook commentaar regels naar object st (textfile)
        '
        'iCycle = nummer van de groep
        'q= 1 alle gaten met Y-spil
        'q = 2 alle gaten met V-spil
        'q > 2 met twee spillen
        'iCyclusKeuze : 1 voorboren, etc.
        '
        Dim strCall As String

        If (mType = machineType.mZayer) Or (mType = machineType.mUnisign) Then


            strCall = "CALL PGM " & cyclus(iCyclusKeuze - 1).ToString
        Else
            strCall = "CYCL CALL M3 M7"

        End If

        'bepaal eerst aantal rijen en aantal gaten per rij (nodig voor text uitvoer file)
        Dim xr As New ArrayList     'lijst met alle x rijcoordinaten
        Dim xv As Double = 9999.9   'startwaarde
        Dim i As Integer = 0

        'bepaal aantal rijen en de rijcoordinaten
        For Each gat As Holes In al
            If gat.x <> xv Then

                xv = gat.x
                xr.Add(gat.x)
            End If
        Next


        'maak ruimte voor y-coordinaten (maak per rij een aparte arraylist)
        Dim yr(xr.Count) As ArrayList
        i = 0
        For Each value As Double In xr
            yr(i) = New ArrayList
            i = i + 1
        Next
        'plaats nu de y coordinaten per rij in de arraylist
        i = 0
        xv = DirectCast(xr(0), Double)
        For Each gat As Holes In al

            If gat.x <> xv Then

                xv = gat.x
                i = i + 1
            End If
            yr(i).Add(gat.y)
        Next
        'aanmaak regels voor textfile
        st.WriteLine()
        For i = 0 To xr.Count - 1       'alle rijen afwerken
            Dim j As Integer
            Dim xi As Double = DirectCast(xr(i), Double)
            Dim yi(yr(i).Count) As Double
            j = 0

            For Each yy As Double In yr(i)
                yi(j) = yy
                j = j + 1
            Next
            st.WriteLine("Rij: {0} met {1} gaten, x= {2}, eerste y={3}, laatste y={4}, ", (i + 1).ToString("###"),
                                                                                          yr(i).Count.ToString("####"),
                                                                                          xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                                          yi(0).ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                                          yi(yr(i).Count - 1).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
        Next
        st.WriteLine("----------------> Totaal aantal gaten: {0} ", al.Count)

        'aanmaak regels voor boorprogramma
        Select Case q

            Case 1  'alle gaten met de Y-spil 

                'wegschrijven met cultureinfo voor decimale punt

                Select Case mType

                    Case machineType.mZayer
                        For Each gat As Holes In al
                            sw.WriteLine("L  X{0}  Y{1} FMAX M71", gat.x.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                   gat.y.ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                            sw.WriteLine(strCall)
                        Next

                    Case machineType.mUnisign
                        For Each gat As Holes In al
                            sw.WriteLine("L  X{0}  Y{1} FMAX", gat.x.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                     gat.y.ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                            sw.WriteLine(strCall)
                        Next

                    Case machineType.mSacem
                        For Each gat As Holes In al
                            sw.WriteLine("L  X{0}  Y{1} FMAX", gat.x.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                     gat.y.ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                            sw.WriteLine(strCall)
                        Next

                    Case Else
                        For Each gat As Holes In al
                            sw.WriteLine("L  X{0}  Y{1} FMAX", gat.x.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                                     gat.y.ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                            sw.WriteLine(strCall)
                        Next


                End Select



            Case 2  'alle gaten met de V spil
                For Each gat As Holes In al
                    sw.WriteLine("L  X{0}  V{1} FMAX M72", gat.x.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                           gat.y.ToString("+##0.###;-##0.###", New CultureInfo("en-US")))

                    sw.WriteLine(strCall)
                Next


            Case Else '3 of 4
                'alle gaten met 2 spillen boren
                '-------------------------------


                'vervolgens alle rijen afwerken
                For i = 0 To xr.Count - 1
                    Dim a As Integer
                    Dim N1, N2 As Integer
                    Dim j As Integer
                    Dim res As Integer          ' rest deling
                    Dim delta As Double         ' delta afstand tussen 2 gaten
                    'de huidige rij met y coordinaten
                    Dim xi As Double = DirectCast(xr(i), Double)
                    Dim yi(yr(i).Count) As Double
                    j = 0

                    For Each yy As Double In yr(i)
                        yi(j) = yy
                        j = j + 1
                    Next
                    'dummy rij maken
                    Dim gg(yr(i).Count + 1) As Boolean
                    For j = 0 To yr(i).Count - 1
                        gg(j) = False           'markeer gaten als false = niet geboord

                    Next

                    a = yr(i).Count
                    If a = Fix(a / 2) * 2 Then

                        'rij met even aantal gaten met 2 spillen
                        ' sw.WriteLine("rij met even aantal gaten met 2 spillen")


                        a = Math.DivRem(a, 2, res)  'a = a/2

                        N1 = 0
                        N2 = N1 + a


                        While N2 < yr(i).Count

                            If gg(N2) = True Then
                                Exit While
                            End If

                            delta = Abs(yi(N2) - yi(N1))        'controleer afstand tussen twee spillen

                            If delta > minDist Then

                                ' boren met 2 spillen indien mogelijk
                                sw.WriteLine("L  X{0}  V{1}  Y{2} FMAX M73",
                                             xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(N2).ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(N1).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))

                                sw.WriteLine(strCall)
                                gg(N1) = True       'markeer als geboord
                                gg(N2) = True
                                N1 = N1 + 1
                                N2 = N2 + 1

                            Else

                                N2 = N2 + 1
                                '---
                            End If

                        End While

                        'resterende gaten boren met 1 spil
                        For j = 0 To yr(i).Count - 1


                            If gg(j) = False Then
                                If q = 3 Then

                                    sw.WriteLine("L  X{0}  V{1} FMAX M72",
                                             xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(j).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                                Else

                                    sw.WriteLine("L  X{0}  Y{1} FMAX M71",
                                             xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(j).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))

                                End If

                                sw.WriteLine(strCall)
                                gg(j) = True
                            End If
                        Next



                    Else
                        'rij met oneven aantal gaten met 2 spillen
                        'Debug.Print(" oneven rij {0} - aantal : {1} ", i, yr(i).Count)
                        'sw.WriteLine(" rij met oneven aantal gaten met 2 spillen")

                        a = Math.DivRem(a, 2, res)
                        N1 = 0
                        N2 = N1 + a


                        'print de verschil afstand tussen twee spillen
                        ' Debug.Print(" delta = " & Abs(yi(N2) - yi(N1)))
                        While N2 < yr(i).Count

                            If (gg(N1) = True) Or (gg(N2) = True) Or (a < 1) Then
                                Exit While
                            End If

                            delta = Abs(yi(N2) - yi(N1))        'controleer afstand tussen twee spillen

                            If delta > minDist Then

                                ' boren met 2 spillen indien mogelijk
                                sw.WriteLine("L  X{0}  V{1}  Y{2} FMAX M73",
                                             xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(N2).ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                             yi(N1).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))

                                sw.WriteLine(strCall)
                                gg(N1) = True       'markeer als geboord
                                gg(N2) = True
                                N1 = N1 + 1
                                N2 = N2 + 1

                            Else

                                N2 = N2 + 1
                                '---
                            End If



                        End While
                        'resterende gaten boren met 1 spil
                        For j = 0 To yr(i).Count - 1


                            If gg(j) = False Then
                                If q = 3 Then

                                    sw.WriteLine("L  X{0}  V{1} FMAX M72",
                                                 xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                 yi(j).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))
                                Else

                                    sw.WriteLine("L X{0} Y{1} FMAX M71",
                                                 xi.ToString("+##0.###;-##0.###", New CultureInfo("en-US")),
                                                 yi(j).ToString("+##0.###;-##0.###", New CultureInfo("en-US")))

                                End If

                                sw.WriteLine(strCall)
                                gg(j) = True
                            End If
                        Next



                    End If

                Next


        End Select



        '--afsluitende bewerking van een cyclus
        Select Case mType
            Case machineType.mZayer
                sw.WriteLine("L  Z+0  W+0 R0 FMAX M91")
            Case machineType.mUnisign
                sw.WriteLine("L  Z-30 R0 FMAX M91")
            Case machineType.mSacem
                sw.WriteLine("L  Z+150 R0 FMAX")
        End Select


    End Sub


End Module
