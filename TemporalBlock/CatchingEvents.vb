Friend Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim Item As String
    Dim User As String


    Public Sub New()

        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

        FirstTask()

    End Sub


    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        End Try

    End Sub


    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()

        Try
            SBOCompany = SBOApplication.Company.GetDICompany
            User = SBOCompany.UserName
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
            'Finally
        End Try

    End Sub


    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()

        Dim loRecSet As SAPbobsCOM.Recordset

        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try

    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()

        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            lofilter.AddEx(150) '// FORMA DATOS MAESTROS DE ARTCULOS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("tekLockItem") '// FORMA BLOQUEAR
            lofilter.AddEx("tekUnlockItem") '// FORMA DESBLOQUEAR
            lofilter.AddEx(150) '// FORMA DATOS MAESTROS DE ARTCULOS

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    Private Sub FirstTask()

        Dim oBI As BlockedItems

        Try
            oBI = New BlockedItems(SBOApplication, SBOCompany)
            oBI.Search()

        Catch ex As Exception
            SBOApplication.MessageBox("FisrtTask: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        ''SBOApplication.MessageBox("Action: " & pVal.Before_Action & "  Type: " & pVal.FormTypeEx)
        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then
        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx

                    Case 150                     '////// FORMA Datos Maestros de Articulo
                        frmItemControllerAfter(FormUID, pVal)

                    Case "tekLockItem"                     '////// FORMA BLOQUEAR
                        frmLockContAf(FormUID, pVal)

                    Case "tekUnlockItem"                     '////// FORMA DESBLOQUEAR
                        frmUnlockContAf(FormUID, pVal)

                End Select
            End If
        End If

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmItemControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim oBI As BItem
        Dim oBlock As FrmtekBlock
        Dim coForm As SAPbouiCOM.Form
        Dim stTabla, DocCur As String
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oDatatable As SAPbouiCOM.DBDataSource
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType
                            '///// Carga de formas
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    oBI = New BItem
                    oBI.addFormItems(FormUID)

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        '--- Boton bloquear del dato maestros del articulo
                        Case "btLck"

                            stTabla = "OITM"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            Item = oDatatable.GetValue("ItemCode", 0)

                            If Item Is Nothing Or Item = "" Then

                                SBOApplication.MessageBox("Por favor selecciona un articulo.")

                            Else

                                oBlock = New FrmtekBlock
                                oBlock.openForm(csDirectory)

                            End If

                        '--- Boton desbloquear del dato maestros del articulo
                        Case "btUlck"

                            stTabla = "OITM"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            Item = oDatatable.GetValue("ItemCode", 0)

                            If Item Is Nothing Or Item = "" Then

                                SBOApplication.MessageBox("Por favor selecciona un articulo.")

                            Else

                                oBlock = New FrmtekBlock
                                oBlock.openFormU(csDirectory)

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            oBI = Nothing
        End Try
    End Sub


    Private Sub frmLockContAf(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim oBkI As BlockItems
        Dim coForm As SAPbouiCOM.Form
        Dim Desde, Hasta As String
        Dim Almacen, Motivo, Id As String
        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim Resultado As Integer
        Dim DUser, DDate, DDDate, DReason As String

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case "5"
                            oBkI = New BlockItems
                            coForm = SBOApplication.Forms.Item(FormUID)
                            Desde = coForm.DataSources.UserDataSources.Item("dsDate").Value
                            Hasta = coForm.DataSources.UserDataSources.Item("dsDDate").Value
                            Almacen = coForm.DataSources.UserDataSources.Item("dsWhs").Value
                            Motivo = coForm.DataSources.UserDataSources.Item("dsMotv").Value
                            Id = "1"

                            If Desde = "" Then

                                SBOApplication.MessageBox("Por favor coloca la fecha inicial del bloqueo.")

                            ElseIf Hasta = "" Then

                                SBOApplication.MessageBox("Por favor coloca la fecha final del bloqueo.")

                            ElseIf Almacen = "-" Then

                                SBOApplication.MessageBox("Por favor coloca el almacen donde se bloqueara el articulo.")

                            ElseIf Motivo = "" Then

                                SBOApplication.MessageBox("Por favor coloca el motivo del bloqueo.")

                            ElseIf Desde <> "" And Hasta <> "" And Almacen <> "-" And Motivo <> "" Then

                                stQueryH = "Select ""ItemCode"",""WhsCode"" from OITW where ""Locked""='Y' and ""ItemCode""=rtrim('" & Item & "') and ""WhsCode""='" & Almacen & "'"
                                oRecSetH.DoQuery(stQueryH)

                                If oRecSetH.RecordCount > 0 Then

                                    oRecSetH.MoveFirst()

                                    stQueryH2 = "Select Top 1 * from ""@TEMPORALBLOCK"" T1 where T1.""U_TypeM""='Bloqueo' and T1.""U_Item""=rtrim('" & Item & "') and T1.""U_WhsCode""='" & Almacen & "' order by to_integer(""Code"") desc"
                                    oRecSetH2.DoQuery(stQueryH2)

                                    If oRecSetH2.RecordCount > 0 Then

                                        DUser = oRecSetH2.Fields.Item("U_User").Value
                                        DDate = oRecSetH2.Fields.Item("U_DocDate").Value
                                        DDDate = oRecSetH2.Fields.Item("U_DocDueDate").Value
                                        DReason = oRecSetH2.Fields.Item("U_Reason").Value

                                        SBOApplication.MessageBox("Este articulo ya fue bloqueado por el usuario " & DUser & " desde " & DDate & " hasta " & DDDate & " por el siguiente motivo: " & DReason)

                                    Else

                                        SBOApplication.MessageBox("Este Articulo esta bloqueado actualmente.")

                                    End If

                                Else

                                    oBkI.BlockingItem(Item, Almacen, Id, User, Desde, Hasta, Motivo)

                                    coForm.DataSources.UserDataSources.Item("dsDate").Value = Nothing
                                    coForm.DataSources.UserDataSources.Item("dsDDate").Value = Nothing
                                    coForm.DataSources.UserDataSources.Item("dsWhs").Value = Nothing
                                    coForm.DataSources.UserDataSources.Item("dsMotv").Value = Nothing

                                End If

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally

        End Try

    End Sub


    Private Sub frmUnlockContAf(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim oBkI As BlockItems
        Dim coForm As SAPbouiCOM.Form
        Dim Desde, Hasta As String
        Dim Almacen, Motivo, Id As String
        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim DUser, DDate, DDDate, DReason As String
        Dim Resultado As Integer

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case "3"
                            oBkI = New BlockItems
                            coForm = SBOApplication.Forms.Item(FormUID)
                            Almacen = coForm.DataSources.UserDataSources.Item("dsWhsU").Value
                            Motivo = coForm.DataSources.UserDataSources.Item("dsMotvU").Value
                            Desde = ""
                            Hasta = ""
                            Id = "2"

                            If Almacen = "-" Then

                                SBOApplication.MessageBox("Por favor coloca el almacen donde se bloqueara el articulo.")

                            ElseIf Motivo = "" Then

                                SBOApplication.MessageBox("Por favor coloca el motivo del bloqueo.")

                            ElseIf Almacen <> "-" And Motivo <> "" Then

                                stQueryH = "Select ""ItemCode"",""WhsCode"" from OITW where ""Locked""='N' and ""ItemCode""=rtrim('" & Item & "') and ""WhsCode""='" & Almacen & "'"
                                oRecSetH.DoQuery(stQueryH)

                                If oRecSetH.RecordCount > 0 Then

                                    SBOApplication.MessageBox("Este Articulo actualmente esta desbloqueado.")

                                Else

                                    stQueryH2 = "Select Top 1 * from ""@TEMPORALBLOCK"" T1 where T1.""U_TypeM""='Bloqueo' and T1.""U_Item""=rtrim('" & Item & "') and T1.""U_WhsCode""='" & Almacen & "' order by to_integer(""Code"") desc"
                                    oRecSetH2.DoQuery(stQueryH2)

                                    If oRecSetH2.RecordCount > 0 Then

                                        DUser = oRecSetH2.Fields.Item("U_User").Value
                                        DDate = oRecSetH2.Fields.Item("U_DocDate").Value
                                        DDDate = oRecSetH2.Fields.Item("U_DocDueDate").Value
                                        DReason = oRecSetH2.Fields.Item("U_Reason").Value

                                        Resultado = SBOApplication.MessageBox("Este articulo fue bloqueado por el usuario " & DUser & " desde " & DDate & " hasta " & DDDate & " por el siguiente motivo: " & DReason & ". ¿Estas seguro que deseas desbloquearlo?", 1, "Continue", "Cancel", "")

                                        If (Resultado = 1) Then

                                            oBkI.BlockingItem(Item, Almacen, Id, User, Desde, Hasta, Motivo)

                                            coForm.DataSources.UserDataSources.Item("dsWhsU").Value = Nothing
                                            coForm.DataSources.UserDataSources.Item("dsMotvU").Value = Nothing

                                        End If

                                    Else

                                        oBkI.BlockingItem(Item, Almacen, Id, User, Desde, Hasta, Motivo)

                                        coForm.DataSources.UserDataSources.Item("dsWhsU").Value = Nothing
                                        coForm.DataSources.UserDataSources.Item("dsMotvU").Value = Nothing

                                    End If

                                End If

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally

        End Try

    End Sub


End Class
