Attribute VB_Name = "FabricCover"
Option Explicit
' Заполняет поля без индекса
Public Function CreateCoverWithOutID(curRegister As C_RegisterInfo) As C_CoverInfo

    Dim cover As C_CoverInfo
    Set cover = New C_CoverInfo
    
    cover.NameEnterprise = GetNameEnterprise()
    cover.OkpoEnterprise = GetOkpoEnterprise()
    
    cover.years = curRegister.getPeriod()
    
    cover.sheetCount = curRegister.GetSheetCount
    cover.lastChange = Format(Now(), "dd.mm.yyyy hh:nn")
    
    Set cover.innerRegistry = curRegister.getCollection()
        
    Set CreateCoverWithOutID = cover
End Function
' Заполняет поля с индексом
Public Function CreateCoverWithID() As C_CoverInfo

    Dim cover As C_CoverInfo
    Set cover = CreateCoverWithOutID()
    
    cover.index = Int(Form_Cover.tb_Index.value)
        
    Set CreateCoverWithID = cover
End Function

Private Sub FillFormFromReport(report As C_CoverInfo)
  Form_Cover.tb_NameEnterprise = cover.NameEnterprise
  Form_Cover.tb_OkpoEnterprise = cover.OkpoEnterprise
  Form_Cover.tb_Years = cover.years
  Form_Cover.tb_sheetCount = cover.sheetCount
End Sub
