Option Explicit On
Option Compare Binary
Option Strict On


Imports System.Collections
'dit is een class die gebruikt wordt om b.v een arraylist te sorteren
'die bestaat uit objecten, die zelf specifieke CompareTo routine bevatten
'dwz de objecten kunnen zichzelf met elkaar vergelijken
'de class is hier specifiek voor het object Holes gemaakt
Public Class ValueComparer

    Implements IComparer

    ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
       Implements IComparer.Compare

        Dim A As Holes = DirectCast(x, Holes)
        Dim B As Holes = DirectCast(y, Holes)


        Return CInt(A.CompareTo(B))
    End Function 'IComparer.Compare

End Class
