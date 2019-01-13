'------------------------------------------------------------------
' frmBoorForm.vb
' Boor1.exe
' Boorprogramma
' 29-11-2007 originele listing MAAKH.BAS in GWBasic/QBasic 
'            begonnen in Visual Studio 2005
' 26-3-2012 omgezet naar Visual Studio 2010
' 3-4-2012  lezen van .asc file gewijzigd , met textfieldparser
'------------------------------------------------------------------
Option Strict Off
Option Explicit On
Imports System
Imports System.Collections
Imports System.Text
Imports System.IO

Public Class frmBoorForm
    Inherits System.Windows.Forms.Form

    'all variables must be declared !!

    Private mNode As System.Windows.Forms.TreeNode

    Private mCurrentIndex As Short
    Private iCycle As Integer

    Dim ssg, pbg, pg, tg1, tg2, bcg As New ArrayList()
    Private strProg, strComment As String
    Dim sw As New StringWriter()


    Private Sub btnClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnClose.Click
        Me.Close()
        End
    End Sub

    Private Sub frmBoorForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load


        'zet waarden op nul e.d.
        Init()


        With Treeview1
            .ImageList = ImageList2
            .BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            .ShowRootLines = True
            .LabelEdit = True

        End With


        mNode = Treeview1.Nodes.Insert(0, "", "Steunstanggaten", "closed")
        mNode = Treeview1.Nodes.Insert(1, "", "Pakketboutgaten", "closed")
        mNode = Treeview1.Nodes.Insert(2, "", "Pijpgaten", "closed")
        mNode = Treeview1.Nodes.Insert(3, "", "Tapgaten (type 1)", "closed")
        mNode = Treeview1.Nodes.Insert(4, "", "Tapgaten (type 2)", "closed")
        mNode = Treeview1.Nodes.Insert(5, "", "Boutcirkelgaten", "closed")

        'window show, anders blijft het op de achtergrond     
        Me.Show()

        'open asc bestand; gelijk bij aanvang van het programma

        With dlgDialogOpen

            .Title = "ExportBestand openen"
            .Filter = "Ascii files (*.asc)|*.asc| All files (*.*)|*.*"
            .CheckFileExists = True
            .CheckPathExists = True
        End With

        dlgDialogOpen.ShowDialog()


        'verwerk het bestand
        'OpenTextFile (dlgDialog.FileName)
        Read_from_file(dlgDialogOpen.FileName)
        ' Ophalen((dlgDialogOpen.FileName))

        Sorteergaten(ssg)
        Controleergaten(ssg)
        Sorteergaten(pbg)
        Controleergaten(pbg)
        Sorteergaten(pg)
        Controleergaten(pg)
        Sorteergaten(tg1)
        Controleergaten(tg1)
        Sorteergaten(tg2)
        Controleergaten(tg2)
        Sorteergaten(bcg)
        Controleergaten(bcg)

        TranslateCycle(1, ssg)
        TranslateCycle(2, pg)

        'Dim gat As Holes
        'For Each gat In ssg
        '    ComboBox1.Items.Add(gat.sInfo)
        'Next
        'ComboBox1.SelectedIndex = 0

        Finish()

        Exit Sub

ErrHandler:
        'user cancelled fileopen dialog

        End
        Exit Sub


    End Sub
    Sub Init()
        'doe begin waarden 
        strProg = "programma.naam"
        sw.WriteLine("BEGIN PGM " & strProg & " MM")
        iCycle = 0

    End Sub
    Sub Finish()
        sw.WriteLine("END PGM " & strProg & "MM")
        TextBox1.AppendText(sw.ToString)

    End Sub
    Sub Read_from_file(ByVal aName As String)
        'lees een *.asc bestand ; dit is het bestand wat gemaakt wordt door AutoLisp
        'bestand openen als comma delimited text file 
        'en alle coordinaten lezen
        'per layer aan de juiste collectie toevoegen
        'het formaat van de regels is < layernaam >, <x> , <y> 
        Dim delimiter As String = ","
        '-----------------
        Dim x, y As Double
        Dim newNode As System.Windows.Forms.TreeNode
        Dim gat As Holes
        Dim stext As String
        Dim iTeller As Short


        Try
            'create an instance of textfieldparser to read from a comma delimited file
            Using sr As New FileIO.TextFieldParser(aName)
                sr.TextFieldType = FileIO.FieldType.Delimited
                sr.SetDelimiters(delimiter)

                Dim currentRow As String()

                While Not sr.EndOfData

                    Try
                        'read all fields on the current line and return them in an array of strings
                        currentRow = sr.ReadFields()

                        Dim ls As String

                        ls = currentRow(0)
                        x = CDbl(currentRow(1))
                        y = CDbl(currentRow(2))
                        '-----------------------------------
                        'steeds een nieuw object maken voor elk nieuw gat
                        gat = New Holes(x, y)


                        x = Fix(x) + CShort((x - Fix(x)) * 1000) / 1000
                        y = Fix(y) + CShort((y - Fix(y)) * 1000) / 1000

                        gat.x = x
                        gat.y = y


                        'de tekst bij de node 
                        stext = ls & " X:" & CStr(x) & " Y:" & CStr(y)
                        Select Case ls
                            Case "SSG"
                                'ssg = ssg + 1
                                'ADD(KEY,TEXT,IMAGEKEY)
                                newNode = Treeview1.Nodes(0).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")
                                newNode.Tag = "SSG"
                                ssg.Add(gat)
                                '                    ssg1.Add(gat)
                                ' ComboBox1.Items.Add(gat.sInfo)

                            Case "PBG"

                                'pbg = pbg + 1
                                newNode = Treeview1.Nodes(1).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")
                                newNode.Tag = "PBG"
                                pbg.Add(gat)

                            Case "PG"
                                'pg = pg + 1
                                newNode = Treeview1.Nodes(2).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")
                                newNode.Tag = "PG"
                                pg.Add(gat)

                            Case "TG1"
                                'tg1 = tg1 + 1
                                newNode = Treeview1.Nodes(3).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")

                                newNode.Tag = "TG1"
                                tg1.Add(gat)

                            Case "TG2"
                                'tg2 = tg2 + 1
                                newNode = Treeview1.Nodes(4).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")
                                newNode.Tag = "TG2"
                                tg2.Add(gat)

                            Case "BCG"
                                'bcg = bcg + 1
                                newNode = Treeview1.Nodes(5).Nodes.Add(iTeller & " ID", _
                                                                        stext, "book")
                                newNode.Tag = "BCG"
                                bcg.Add(gat)

                        End Select

                        iTeller = iTeller + 1
                        '-----------------------------------


                    Catch ex As Exception

                    End Try

                End While


                sr.Close()
            End Using



        Catch ex As Exception
            'let the user know what went wrong
            Console.WriteLine(ex.Message)

        End Try



    End Sub
   

    Sub Sorteergaten(ByVal alst As ArrayList)
        'sorteer de objecten op coordinaten van laag naar hoog
        'dus van -x,-y naar +x,+y eerst per x rij dan per y kolom
        'Opm. de arraylist wordt gesorteerd, maar niet de collection !!
        'Dim alst As ArrayList = New ArrayList(boorgroep)

        alst.Sort(New ValueComparer())


    End Sub
    Sub Controleergaten(ByVal al As ArrayList)
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
                    MsgBox("dubbel coordinaat !! :   " & vbCrLf & _
                            "diffy :  " & CStr(diffy) & vbCrLf & _
                            al.Item(i).sInfo & vbCrLf & _
                            al.Item(i - 1).sInfo)
                    Exit For
                End If
            End If
        Next

    End Sub
    Sub TranslateCycle(ByVal q As Integer, ByVal al As ArrayList)
        'maak het boorprogramma voor de gesorteerde boorgroep
        'in de arraylist
        ' schrijf programma regels als strings naar object sw (stringwriter)
        ' deze wordt ineen keer toegevoegd aan textbox1 met .appendtext()
        '
        'iCycle = nummer van de groep
        'q= 1 alle gaten met Y-spil
        'q = 2 alle gaten met V-spil
        'q > 2 met twee spillen
        Dim strCall As String
        Dim cyclus As Integer

        strCall = "CALL PGM " & CStr(cyclus)
        strComment = "commentaar.regel"
        If iCycle > 0 Then
            sw.WriteLine("M0 ")

        End If
        iCycle = iCycle + 1
        sw.WriteLine(strComment)

        'voor q = 1
        Dim gat As Holes
        For Each gat In al
            sw.WriteLine("L X {0} Y {1} FMAX M71", gat.x, gat.y)
            sw.WriteLine(strCall)
        Next

        sw.WriteLine("L Z+0 W+0 R0 FMAX M91")



    End Sub


    Private Sub TreeView1_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles Treeview1.NodeMouseClick
        Dim Node As System.Windows.Forms.TreeNode = eventArgs.Node

        With StatusBar1
            .Items.Clear()
            .Items.Add(Node.Text)
            .Items.Add(Node.GetNodeCount(True) & " items")

            .Items.Item(1).AutoSize = True

        End With

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        AboutBox1.ShowDialog()

    End Sub




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' 
        'demo voor afronden van een getal
        Dim x, y As Double

        x = 3.1234
        y = 5.5678
        MsgBox("before X :  " & x & " Y :  " & y)
        x = Fix(x) + CShort((x - Fix(x)) * 1000) / 1000
        y = Fix(y) + CShort((y - Fix(y)) * 1000) / 1000
        MsgBox("after X :  " & x & " Y :  " & y)


    End Sub
End Class