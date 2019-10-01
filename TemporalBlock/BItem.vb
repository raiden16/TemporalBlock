Public Class BItem

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private Directorio As String

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Directorio = oCatchingEvents.csDirectory
    End Sub

    '//----- AGREGA ELEMENTOS A LA FORMA
    Public Sub addFormItems(ByVal FormUID As String)
        Dim loItem As SAPbouiCOM.Item
        Dim loButton As SAPbouiCOM.Button
        Dim lsItemRef As String

        Try
            '//AGREGA BOTON BLOQUEAR EN DATOS MAESTROS DE ARTICULOS
            coForm = cSBOApplication.Forms.Item(FormUID)
            lsItemRef = "2"
            loItem = coForm.Items.Add("btLck", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 5
            loItem.Top = coForm.Items.Item(lsItemRef).Top
            loItem.Width = coForm.Items.Item(lsItemRef).Width + 40
            loItem.Height = coForm.Items.Item(lsItemRef).Height
            loButton = loItem.Specific
            loButton.Caption = "Bloquear"

            '//AGREGA BOTON DESBLOQUEAR EN DATOS MAESTROS DE ARTICULOS
            coForm = cSBOApplication.Forms.Item(FormUID)
            lsItemRef = "btLck"
            loItem = coForm.Items.Add("btUlck", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 5
            loItem.Top = coForm.Items.Item(lsItemRef).Top
            loItem.Width = coForm.Items.Item(lsItemRef).Width
            loItem.Height = coForm.Items.Item(lsItemRef).Height
            loButton = loItem.Specific
            loButton.Caption = "Desbloquear"

            '//AGREGA BOTON DESBLOQUEAR EN DATOS MAESTROS DE ARTICULOS
            coForm = cSBOApplication.Forms.Item(FormUID)
            lsItemRef = "btUlck"
            loItem = coForm.Items.Add("btLog", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + 5
            loItem.Top = coForm.Items.Item(lsItemRef).Top
            loItem.Width = coForm.Items.Item(lsItemRef).Width
            loItem.Height = coForm.Items.Item(lsItemRef).Height
            loButton = loItem.Specific
            loButton.Caption = "Log_Bloqueo"

        Catch ex As Exception
            cSBOApplication.MessageBox("addFormItems. agregar elementos a la forma. " & ex.Message)
        Finally
            coForm = Nothing
            loItem = Nothing
            loButton = Nothing
        End Try
    End Sub


End Class
