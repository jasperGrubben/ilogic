Imports Inventor

AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensMain.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensVault.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"


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
