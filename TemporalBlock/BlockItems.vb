Public Class BlockItems

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Sub BlockingItem(ByVal Item As String, ByVal WhsHouse As String, ByVal Id As String, ByVal User As String, ByVal Desde As String, ByVal Hasta As String, ByVal Motivo As String)

        Dim stQueryH, Line As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim llError As Long
        Dim lsError As String
        Dim oItem As SAPbobsCOM.Items
        Dim oItemWI As SAPbobsCOM.ItemWarehouseInfo
        Dim TypeM, CreateDate As String

        oItem = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        Try

            oItem.GetByKey(Item)
            oItemWI = oItem.WhsInfo

            stQueryH = "Select ""RowNum"" from (Select ROW_NUMBER() OVER (ORDER BY ""WhsCode"")-1 AS ""RowNum"",""WhsCode"" from OITW where ""ItemCode""=rtrim('" & Item & "')) T0 where T0.""WhsCode""='" & WhsHouse & "'"
            oRecSetH.DoQuery(stQueryH)


            If oRecSetH.RecordCount > 0 Then

                Line = oRecSetH.Fields.Item("RowNum").Value

            End If

            oItemWI.SetCurrentLine(Line)

            If Id = "1" Then
                oItemWI.Locked = 1               '0="N"   1="Y"
                TypeM = "Bloqueo"
            Else
                oItemWI.Locked = 0              '0="N"   1="Y"
                TypeM = "Desbloqueo"
            End If

            If oItem.Update() = 0 Then

                CreateDate = Now.Year & "/" & Now.Month & "/" & Now.Day

                UpdateTable(TypeM, User, CreateDate, Item, WhsHouse, Desde, Hasta, Motivo)

            Else
                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("BlockItems BlockingItem. bloqueando el articulo del almacen. " & ex.Message)

        End Try

    End Sub


    Public Sub UpdateTable(ByVal TypeM As String, ByVal User As String, ByVal CreateDate As String, ByVal Item As String, ByVal WhsCode As String, ByVal DocDate As String, ByVal DocDueDate As String, ByVal Reason As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim stQueryH2 As String
        Dim oRecSetH2 As SAPbobsCOM.Recordset
        Dim DocEntry As String
        Dim fechaI, fechaf As String

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select Count(""Code"") as ""Consecutivo"" from ""@TEMPORALBLOCK"""
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                DocEntry = oRecSetH.Fields.Item("Consecutivo").Value + 1

            End If

            If DocDate = "" Then
                fechaI = ""
            Else
                fechaI = DocDate.Substring(6, 4) & "/" & DocDate.Substring(3, 2) & "/" & DocDate.Substring(0, 2)
            End If

            If DocDate = "" Then
                fechaf = ""
            Else
                fechaf = DocDueDate.Substring(6, 4) & "/" & DocDueDate.Substring(3, 2) & "/" & DocDueDate.Substring(0, 2)
            End If

            stQueryH2 = "insert into ""DESARROLLAR1"".""@TEMPORALBLOCK"" values ('" & DocEntry & "', '" & DocEntry & "','" & TypeM & "','" & User & "','" & CreateDate & "',rtrim('" & Item & "'),'" & WhsCode & "','" & fechaI & "','" & fechaf & "','" & Reason & "');"
            oRecSetH2.DoQuery(stQueryH2)

        Catch ex As Exception

            SBOApplication.MessageBox("BlockItems UpdateTable. actualizar tabla de log. " & ex.Message)

        End Try

    End Sub


End Class
