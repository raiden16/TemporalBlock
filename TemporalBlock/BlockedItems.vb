Public Class BlockedItems

    Private bSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private bSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oItem As SAPbobsCOM.Items
    Dim oItemWI As SAPbobsCOM.ItemWarehouseInfo

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New(ByVal SBOApplication As SAPbouiCOM.Application, ByVal SBOCompany As SAPbobsCOM.Company)
        MyBase.New()
        bSBOApplication = SBOApplication
        bSBOCompany = SBOCompany
    End Sub

    Public Function Search()

        Dim stQueryH, Item, WareHouse As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim DueDate, CurrDate As Date
        Dim stQueryH2 As String
        Dim oRecSetH2 As SAPbobsCOM.Recordset

        oRecSetH = bSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = bSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""ItemCode"",""WhsCode"" from OITW where ""Locked""='Y'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    CurrDate = Now.Date
                    Item = oRecSetH.Fields.Item("ItemCode").Value
                    WareHouse = oRecSetH.Fields.Item("WhsCode").Value

                    stQueryH2 = "Select Top 1 * from ""@TEMPORALBLOCK"" T1 where T1.""U_TypeM""='Bloqueo' and T1.""U_Item""='" & Item & "' and T1.""U_WhsCode""='" & WareHouse & "' order by to_integer(""Code"") desc"
                    oRecSetH2.DoQuery(stQueryH2)

                    If oRecSetH2.RecordCount > 0 Then

                        oRecSetH2.MoveFirst()

                        For l = 0 To oRecSetH2.RecordCount - 1

                            DueDate = oRecSetH2.Fields.Item("U_DocDueDate").Value

                            If CurrDate > DueDate Then

                                UnlockItem(Item, WareHouse)

                            End If

                            If l < oRecSetH2.RecordCount - 1 Then
                                oRecSetH2.MoveNext()
                            End If

                        Next

                    End If

                    If i < oRecSetH.RecordCount - 1 Then
                        oRecSetH.MoveNext()
                    End If

                Next

            End If

        Catch ex As Exception

            bSBOApplication.MessageBox("BlockedItems Search. busqueda de articulos bloqueados en almacen. " & ex.Message)

        End Try

    End Function


    Public Function UnlockItem(ByVal Item As String, ByVal WhsHouse As String)

        oItem = bSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        Dim stQueryH, Line As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim llError As Long
        Dim lsError As String

        oRecSetH = bSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            oItem.GetByKey(Item)
            oItemWI = oItem.WhsInfo

            stQueryH = "Select ""RowNum"" from 
                        (Select ROW_NUMBER() OVER (ORDER BY ""WhsCode"")-1 AS ""RowNum"",""WhsCode"" from OITW where ""ItemCode""='" & Item & "') T0
                        where T0.""WhsCode""='" & WhsHouse & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                Line = oRecSetH.Fields.Item("RowNum").Value

            End If

            oItemWI.SetCurrentLine(Line)
            oItemWI.Locked = 0              '0="N"   1="Y"

            If oItem.Update() <> 0 Then

                bSBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            End If

        Catch ex As Exception

            bSBOApplication.MessageBox("BlockedItems UnlockItem. desbloquear articulo de almacen. " & ex.Message)

        End Try

    End Function


End Class
