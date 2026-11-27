AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensMain.vb"

Public Class RuleManager

    Public Shared Sub Run(type As String)
        FlensGlobals.ThisApp = ThisApplication
        FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

        Select Case type
            Case "Flenzen"
                FlensRunner.Handle(BuildInput())
            Case "Anders"
                FlensGlobals.Show("Anders wat!")
            Case Else
                FlensGlobals.Show("Onbekend type: " & type)
        End Select
    End Sub

    Private Shared Function BuildInput() As Input
        Return New Input With {
            .InputLength = Parameter("InputLengte"),
            .InputWidth = Parameter("InputBreedte"),
            .InputFlangeHeight = Parameter("InputFlensHoogte"),
            .InputFlangeWidth = Parameter("InputFlensBreedte"),
            .InputHoleDiameter = Parameter("InputGatDia")
        }
    End Function

End Class

Sub Main()
    RuleManager.Run(Parameter("Type"))
End Sub