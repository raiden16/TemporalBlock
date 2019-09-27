Imports System.Drawing

Public Class FrmtekBlock

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String)
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim key As String
        'Dim Monto As Integer

        Try

            key = "1"
            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            csFormUID = "tekLockItem"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID, key)

            cargarComboAlmacenes(key)

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            Return Monto

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openFormU(ByVal psDirectory As String)
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim key As String
        'Dim Monto As Integer

        Try

            key = "2"
            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            csFormUID = "tekUnlockItem"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID, key)

            cargarComboAlmacenes(key)

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            Return Monto

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String, ByVal key As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources(key)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources(ByVal key As String) As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources(key)
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources(ByVal key As String) As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid
        Dim oCombo As SAPbouiCOM.ComboBox

        Try
            bindUserDataSources = 0

            If key = "1" Then

                loDS = coForm.DataSources.UserDataSources.Add("dsDate", SAPbouiCOM.BoDataType.dt_DATE)
                loText = coForm.Items.Item("1").Specific    'identifico mi caja de texto
                loText.DataBind.SetBound(True, "", "dsDate")    ' uno mi userdatasources a mi caja de fecha

                loDS = coForm.DataSources.UserDataSources.Add("dsDDate", SAPbouiCOM.BoDataType.dt_DATE)
                loText = coForm.Items.Item("2").Specific    'identifico mi caja de texto
                loText.DataBind.SetBound(True, "", "dsDDate")    ' uno mi userdatasources a mi caja de fecha

                loDS = coForm.DataSources.UserDataSources.Add("dsWhs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                oCombo = coForm.Items.Item("3").Specific  'identifico mi combobox
                oCombo.DataBind.SetBound(True, "", "dsWhs")   ' uno mi userdatasources a mi combobox

                loDS = coForm.DataSources.UserDataSources.Add("dsMotv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                loText = coForm.Items.Item("4").Specific  'identifico mi caja de texto
                loText.DataBind.SetBound(True, "", "dsMotv")   ' uno mi userdatasources a mi caja de texto

            ElseIf key = "2" Then

                loDS = coForm.DataSources.UserDataSources.Add("dsMotvU", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                loText = coForm.Items.Item("1").Specific  'identifico mi caja de texto
                loText.DataBind.SetBound(True, "", "dsMotvU")   ' uno mi userdatasources a mi caja de texto

                loDS = coForm.DataSources.UserDataSources.Add("dsWhsU", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                oCombo = coForm.Items.Item("2").Specific  'identifico mi combobox
                oCombo.DataBind.SetBound(True, "", "dsWhsU")   ' uno mi userdatasources a mi combobox

            End If

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '---- Carga de Porcentajes
    Public Function cargarComboAlmacenes(ByVal key As String)

        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecSet As SAPbobsCOM.Recordset

        Try
            cargarComboAlmacenes = 0
            '--- referencia de combo 
            If key = "1" Then
                oCombo = coForm.Items.Item("3").Specific
            Else
                oCombo = coForm.Items.Item("2").Specific
            End If
            coForm.Freeze(True)
            '---- SI YA SE TIENEN VALORES, SE ELIMMINAN DEL COMBO
            If oCombo.ValidValues.Count > 0 Then
                Do
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Loop While oCombo.ValidValues.Count > 0
            End If
            '--- realizar consulta
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet.DoQuery("select '-' as ""WhsCode"",null as ""WhsName"" from dummy union all Select ""WhsCode"",""WhsName"" from ""OWHS"" order by ""WhsCode""")
            '---- cargamos resultado
            oRecSet.MoveFirst()
            Do While oRecSet.EoF = False
                oCombo.ValidValues.Add(oRecSet.Fields.Item(0).Value, oRecSet.Fields.Item(1).Value)
                oRecSet.MoveNext()
            Loop
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            coForm.Freeze(False)


        Catch ex As Exception
            coForm.Freeze(False)
            MsgBox("FrmTratamientoPedidos. cargarComboPorcentaje: " & ex.Message)
        Finally
            oCombo = Nothing
            oRecSet = Nothing
        End Try
    End Function


End Class
