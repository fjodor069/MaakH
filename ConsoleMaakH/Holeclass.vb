Imports System.Collections.Generic

Public Class Holes
    Implements IComparable(Of Holes)

    Protected xv As Double
    Protected yv As Double
    Private strInfo As String

    'deze method wordt gebruikt om twee objecten te vergelijken bv om te sorteren
    'wordt aangeroepen door class valuecomparer
    Public Function CompareTo(ByVal other As Holes) As Integer _
        Implements IComparable(Of Holes).CompareTo


        'sorteer eerst op x dan op y
        'dus als beide gelijke x dan op y sorteren
        If Me.x.Equals(other.x) = True Then
            Return y.CompareTo(other.y)
        Else
            'als niet gelijke x dan gewoon op x sorteren
            Return x.CompareTo(other.x)
        End If




    End Function


    Public Sub New(ByVal x As Double, ByVal y As Double)
        xv = x
        yv = y
    End Sub
    Public Property x() As Double
        Get
            x = xv
        End Get
        Set(ByVal value As Double)
            xv = value
        End Set
    End Property
    Public Property y() As Double
        Get
            y = yv
        End Get
        Set(ByVal value As Double)
            yv = value
        End Set
    End Property
    ReadOnly Property sInfo() As String
        Get
            sInfo = " X:" & CStr(x) & " Y:" & CStr(y)
        End Get

    End Property

End Class


Public Class HolesyFirst
    Inherits Comparer(Of Holes)

    Public Overrides Function Compare(ByVal H1 As Holes, ByVal H2 As Holes) As Integer

        'sorteer eerst op y dan op x
        'dus als beide gelijke y dan op x sorteren
        If H1.y.Equals(H2.y) = True Then
            Return H1.x.CompareTo(H2.x)
        Else
            'als niet gelijke y dan gewoon op y sorteren
            Return H1.y.CompareTo(H2.y)
        End If


    End Function

End Class
