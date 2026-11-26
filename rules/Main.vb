'Imports Inventor
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensMain.vb"

Sub Main()
    'paden fixen voor OOP progammering
    FlensGlobals.Init(ThisApplication, AddressOf Logger.Info, AddressOf Logger.Error)
   dim type = Parameter("Type")
   select case type
   case "Flenzen"
    new MainFlenzen(model)
   case "Anders"
    FlensGlobals.Show("Anders wat!")
    End select 
End Sub
