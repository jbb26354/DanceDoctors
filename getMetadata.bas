Function tblGetMetadata()

Dim aTableDef As DAO.TableDef
Dim AField As DAO.Field
Dim TheTargetTable As DAO.Recordset

DoCmd.SetWarnings False
DoCmd.RunSQL ("delete * from Metadata;")
Set TheTargetTable = CurrentDb.OpenRecordset("Metadata")

With TheTargetTable
  .AddNew
    For Each aTableDef In CurrentDb.TableDefs
      If (aTableDef.Name <> "AccessDataTypes" _
          And Left(aTableDef.Name, 4) <> "MSys" _
          And aTableDef.Name <> "Metadata" _
          And Left(aTableDef.Name, 1) <> "~" _
          And Left(aTableDef.Name, 2) <> "ZZ" _
          ) = True Then
        For Each AField In aTableDef.Fields
          .AddNew
            !TheTable = aTableDef.Name
            !TheField = AField.Name
            Select Case AField.Type
              Case 1
                !TheDataType = "YesNo"
              Case 2
                !TheDataType = "Byte"
              Case 3
                !TheDataType = "Integer(16)"
              Case 4
                !TheDataType = "Long(32)"
              Case 5
                !TheDataType = "Currency"
              Case 7
                !TheDataType = "Double(64)"
              Case 8
                !TheDataType = "DateTime"
              Case 10
                !TheDataType = "Text"
              Case 11
                !TheDataType = "OLEObject"
              Case 12
                !TheDataType = "Memo"
              Case 20
                !TheDataType = "Decimal"
              Case Else
                !TheDataType = CStr(AField.Type)
            End Select
            !TheSize = AField.Size
            !IsRequired = AField.Properties("Required").Value
            !AllowsZeroLength = AField.Properties("AllowZeroLength").Value
            !TheDefaultValue = AField.Properties("DefaultValue").Value
            On Error Resume Next
              ' DO NOT use the ordinal integer value for the Description property call below.
              ' Always use the name. Or it won't work.
              !TheDescription = AField.Properties("Description").Value
            If Err.Number = 3265 Then
              !TheDescription = "Undefined"
            End If
            On Error GoTo 0
            !TheTimestamp = Date & " " & Time
          .Update
        Next
      End If
    Next
End With

TheTargetTable.Close
DoCmd.OpenTable ("Metadata")

End Function
