Attribute VB_Name = "FabricCover"
Option Explicit
' Заполняет поля без индекса
Public Function CreateCoverWithOutID(curRegister As C_RegisterInfo) As C_CoverInfo

    Dim cover As C_CoverInfo
    Set cover = New C_CoverInfo
    
    cover.NameEnterprise = GetNameEnterprise()
    cover.OkpoEnterprise = GetOkpoEnterprise()
    cover.numberInBase = GetNumberInBase()
    
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
