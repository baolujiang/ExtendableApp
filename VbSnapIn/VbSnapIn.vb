Imports CommonSnappableTypes

<CompanyInfo(CompanyName:="Chuck's Software", CompanyUrl:="www.chucksoftware.com")>
Public Class VbSnapIn
    Implements IAppFunctionality

    Public Sub DoIt() Implements IAppFunctionality.DoIt
        Console.WriteLine("You have just used the VB snap in!")
    End Sub
End Class
