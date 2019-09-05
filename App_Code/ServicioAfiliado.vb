Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.IO
Imports System.Xml.Serialization

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.DataSet
Imports System.Drawing
Imports System.Text

<WebService(Namespace:="www.comfacasanare.com.co")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    Private CsDatos As New Datos
    Private CsError As New ErrorWs

    <WebMethod()> _
    Public Function ConsultarAfiliado(TipoDocumento_CC_CE As String, IdentificacionAfiliado As String) As XmlDocument

        Dim InfoAfiliado As New ArrayList
        Dim InfoEmpresa As New ArrayList
        Dim ClaseParentescoAfil As String

        Dim posicion As Integer
        Dim cadena As String

        Try

            ClaseParentescoAfil = ""

            If Trim(IdentificacionAfiliado) = "" Then
                CsError.MensajeError = "Campo NumeroIdentificacion entro en blanco"
                Return XmlError(CsError)
            End If

            If Trim(TipoDocumento_CC_CE) = "" Then
                CsError.MensajeError = "Campo TipoDocumento_CC_CE entro en blanco (debe escribir CC o CE)"
                Return XmlError(CsError)
            End If

		
            '********************************************************************************
            'analizar caracteres claves
            cadena = TipoDocumento_CC_CE.ToUpper & " " & IdentificacionAfiliado.ToUpper
            posicion = cadena.LastIndexOf("'")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("?")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("%")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("<>")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("<")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf(">")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("*")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("$")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("#")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("SELECT")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("INSERT")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("INTO")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("VALUES")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("DELETE")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("UPDATE")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("WHERE")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("=")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf("(")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If
            posicion = cadena.LastIndexOf(")")
            If posicion >= 0 Then
                CsError.MensajeError = "En los campos no es posible escribir caracteres especiales"
                Return XmlError(CsError)
            End If

              
            '********************************************************************************

            CadenaDeConexion = ConfigurationManager.ConnectionStrings.Item("Cadena").ConnectionString
            RutaSalida = ConfigurationManager.ConnectionStrings.Item("RutaSalida").ConnectionString

            FechaIngreso = Format(Now, "yyyyMMdd")
            HoraIngreso = Format(Now, "HHmmss")
		
		

            If Not CsDatos.Conectar_BD(CadenaDeConexion) Then
                CsError.MensajeError = "Error con conexión a la Base de Datos"
                Return XmlError(CsError)
            End If
	
		  CsDatos.IniciarTransaccion()

            ConsultaAfiliado(TipoDocumento_CC_CE, IdentificacionAfiliado, InfoAfiliado)

            If Trim(IdentificacionAfiliado) = "" Then
                CsError.MensajeError = "Campo NumeroIdentificacion no está relacionado en la Base de Datos"
                Return XmlError(CsError)
            End If

            ConsultaBeneficiarios(TipoDocumento_CC_CE, IdentificacionAfiliado, InfoAfiliado)

            ConsultaEmpresa(TipoDocumento_CC_CE, IdentificacionAfiliado, InfoEmpresa, ClaseParentescoAfil)


CsDatos.InterrumpirTransaccion()


            CsDatos.CerrarBaseDatos()



            ConsultarAfiliado = CrearXmlSalida(InfoAfiliado, InfoEmpresa, ClaseParentescoAfil)

        Catch ex As Exception
            '>>>
            CsDatos.InterrumpirTransaccion()
            CsDatos.CerrarBaseDatos()
            '>>>
            CsError.MensajeError = ex.Message
            Return XmlError(CsError)
        End Try

    End Function

    Private Sub ConsultaAfiliado(TipoDoc As String, ByRef NumeroIdentificacion As String, ByRef InfoAfiliado As ArrayList)

        Dim ClsInformacionAfiliado As InformacionAfiliado

        Dim DataSetAfiliado As New DataSet
        Dim Fila As DataRow
        Dim SqlConsulta As String

        Dim TipoDocumento As String
        Dim estadoCivil As String
        Dim estado As String
        Dim FechaDate As Date
        Dim fechaFormat As String
        

        Try

            Select Case TipoDoc.ToUpper
                Case "CC"
                    TipoDoc = "1"
                Case "CE"
                    TipoDoc = "4"
                Case Else
                    TipoDoc = "9"
            End Select

            'SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "'"
            SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "' AND coddoc = '" & TipoDoc & "'"
            
            
            
            DataSetAfiliado = CsDatos.Consultar(SqlConsulta)
            

            If DataSetAfiliado.Tables(0).Rows.Count = 0 Then
                'NO SE ENCONTRO DOCUMENTO EN NINGUNA TABLA, ERROR
                NumeroIdentificacion = ""
                Exit Sub
            End If

            For Each Fila In DataSetAfiliado.Tables(0).Rows

                ClsInformacionAfiliado = New InformacionAfiliado()


                ClsInformacionAfiliado.IdBeneficiario = Fila("cedtra") & ""

                TipoDocumento = ""
                TipoDocumento = Fila("coddoc") & ""

                Select Case TipoDocumento
                    Case "1"
                        TipoDocumento = "1"
                    Case "2"
                        TipoDocumento = "3"
                    Case "3"
                        TipoDocumento = "2"
                    Case "4"
                        TipoDocumento = "4"
                    Case "5"
                        TipoDocumento = "7"
                    Case "6"
                        TipoDocumento = "5"
                    Case Else
                        TipoDocumento = ""
                End Select

                ClsInformacionAfiliado.TidBeneficiario = TipoDocumento
                ClsInformacionAfiliado.AlfaDocumento = ""
                ClsInformacionAfiliado.PrimerApellido = Fila("priape") & ""
                ClsInformacionAfiliado.SegundoApellido = Fila("segape") & ""
                ClsInformacionAfiliado.PrimerNombre = Fila("prinom") & ""
                ClsInformacionAfiliado.SegundoNombre = Fila("segnom") & ""

                Try
                    FechaDate = Fila("fecnac") & ""
                    fechaFormat = Format(FechaDate, "yyyyMMdd")
                    ClsInformacionAfiliado.FechaNacimiento = fechaFormat
                Catch
                    ClsInformacionAfiliado.FechaNacimiento = Fila("fecnac") & ""
                End Try

                estadoCivil = ""
                estadoCivil = Fila("estciv") & ""
                Select Case Fila("estciv") & ""
                    Case "1"
                        estadoCivil = "SO"
                    Case "2"
                        estadoCivil = "CA"
                    Case "3"
                        estadoCivil = "VI"
                    Case "5"
                        estadoCivil = "UL"
                    Case "6"
                        estadoCivil = "SE"
                    Case Else
                        estadoCivil = ""
                End Select

                ClsInformacionAfiliado.EstadoCivil = estadoCivil

                ClsInformacionAfiliado.Genero = Fila("sexo") & ""

                estado = ""
                estado = Fila("estado") & ""
                Select Case estado
                    Case "A"
                        estado = "0"
                    Case "I"
                        estado = "1"
                    Case Else
                        estado = ""
                End Select
                ClsInformacionAfiliado.Estado = estado
                ClsInformacionAfiliado.Categoria = Fila("codcat") & ""
                'PREGUNTAR ?
                ClsInformacionAfiliado.ClaseParentesco = ""

                ClsInformacionAfiliado.NumeroTarjeta = "0"
                ClsInformacionAfiliado.Ciudad = Fila("codzon") & ""
                ClsInformacionAfiliado.Direccion = Fila("direccion") & ""
                ClsInformacionAfiliado.Telefono = Fila("telefono") & ""
                ClsInformacionAfiliado.Celular = Fila("telefono") & ""
                ClsInformacionAfiliado.CorreoElectronico = ""
                ClsInformacionAfiliado.IdTrabajador = Fila("cedtra") & ""
                ClsInformacionAfiliado.TidTrabajador = TipoDocumento
                ClsInformacionAfiliado.NombreCompletoTrabajador = Fila("priape") & " " & Fila("segape") & " " & Fila("prinom") & " " & Fila("segnom") & ""

                InfoAfiliado.Add(ClsInformacionAfiliado)
                Exit For

            Next

        Catch ex As Exception
            Throw (New Exception(ex.Message))
        End Try
    End Sub

    Private Sub ConsultaBeneficiarios(TipoDoc As String, ByRef NumeroIdentificacion As String, ByRef InfoAfiliado As ArrayList)

        Dim ClsInformacionAfiliado As InformacionAfiliado

        Dim DataSetBeneficiarios As New DataSet
        Dim DataSetAfiliado As New DataSet
        Dim DataSetAux As New DataSet
        Dim Fila, FilaAux As DataRow
        Dim SqlConsulta, estadoCivil As String

        Dim codcat, codzon, direccion, telefono As String
        Dim TipoDocumento, Parentesco As String
        Dim TipoDocumentoTra, estado As String
        Dim FechaDate As Date
        Dim fechaFormat As String

        Try
            codcat = ""
            codzon = ""
            direccion = ""
            telefono = ""
            TipoDocumentoTra = ""

            Select Case TipoDoc.ToUpper
                Case "CC"
                    TipoDoc = "1"
                Case "CE"
                    TipoDoc = "4"
                Case Else
                    TipoDoc = "1"
            End Select

            'SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "'"
            SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "' AND coddoc = '" & TipoDoc & "'"
            DataSetAfiliado = CsDatos.Consultar(SqlConsulta)

            For Each Fila In DataSetAfiliado.Tables(0).Rows

                codcat = Fila("codcat") & ""
                codzon = Fila("codzon") & ""
                direccion = Fila("direccion") & ""
                telefono = Fila("telefono") & ""

                TipoDocumentoTra = ""
                TipoDocumentoTra = Fila("coddoc") & ""
                Select Case TipoDocumentoTra
                    Case "1"
                        TipoDocumentoTra = "1"
                    Case "2"
                        TipoDocumentoTra = "3"
                    Case "3"
                        TipoDocumentoTra = "2"
                    Case "4"
                        TipoDocumentoTra = "4"
                    Case "5"
                        TipoDocumentoTra = "7"
                    Case "6"
                        TipoDocumentoTra = "5"
                    Case Else
                        TipoDocumentoTra = ""
                End Select

                Exit For
            Next

            SqlConsulta = "SELECT * FROM subsi23 WHERE cedtra = '" & NumeroIdentificacion & "'"
            DataSetBeneficiarios = CsDatos.Consultar(SqlConsulta)

            For Each Fila In DataSetBeneficiarios.Tables(0).Rows

                SqlConsulta = "SELECT * FROM subsi22 WHERE codben = " & Val(Fila("codben") & "")
                DataSetAux.Tables.Clear()
                DataSetAux = CsDatos.Consultar(SqlConsulta)

                For Each FilaAux In DataSetAux.Tables(0).Rows

                    ClsInformacionAfiliado = New InformacionAfiliado()

                    ClsInformacionAfiliado.IdBeneficiario = FilaAux("documento") & ""

                    TipoDocumento = ""
                    TipoDocumento = FilaAux("coddoc") & ""
                    Select Case TipoDocumento
                        Case "1"
                            TipoDocumento = "1"
                        Case "2"
                            TipoDocumento = "3"
                        Case "3"
                            TipoDocumento = "2"
                        Case "4"
                            TipoDocumento = "4"
                        Case "5"
                            TipoDocumento = "7"
                        Case "6"
                            TipoDocumento = "5"
                        Case Else
                            TipoDocumento = ""
                    End Select

                    ClsInformacionAfiliado.TidBeneficiario = TipoDocumento
                    ClsInformacionAfiliado.AlfaDocumento = ""
                    ClsInformacionAfiliado.PrimerApellido = FilaAux("priape") & ""
                    ClsInformacionAfiliado.SegundoApellido = FilaAux("segape") & ""
                    ClsInformacionAfiliado.PrimerNombre = FilaAux("prinom") & ""
                    ClsInformacionAfiliado.SegundoNombre = FilaAux("segnom") & ""

                    Try
                        FechaDate = FilaAux("fecnac") & ""
                        fechaFormat = Format(FechaDate, "yyyyMMdd")
                        ClsInformacionAfiliado.FechaNacimiento = fechaFormat
                    Catch
                        ClsInformacionAfiliado.FechaNacimiento = FilaAux("fecnac") & ""
                    End Try

                    estadoCivil = ""
                    ClsInformacionAfiliado.EstadoCivil = estadoCivil

                    ClsInformacionAfiliado.Genero = FilaAux("sexo") & ""

                    estado = ""
                    estado = FilaAux("estado") & ""
                    Select Case estado
                        Case "A"
                            estado = "0"
                        Case "I"
                            estado = "1"
                        Case Else
                            estado = ""
                    End Select
                    ClsInformacionAfiliado.Estado = estado
                    ClsInformacionAfiliado.Categoria = codcat

                    Parentesco = ""
                    Parentesco = FilaAux("parent") & ""
                    Select Case Parentesco
                        Case "1"
                            Parentesco = "HI"
                        Case "2"
                            Parentesco = "HT"
                        Case "3"
                        Case Else
                            Parentesco = ""
                    End Select
                    ClsInformacionAfiliado.ClaseParentesco = Parentesco

                    ClsInformacionAfiliado.NumeroTarjeta = "0"
                    ClsInformacionAfiliado.Ciudad = codzon
                    ClsInformacionAfiliado.Direccion = direccion
                    ClsInformacionAfiliado.Telefono = telefono
                    ClsInformacionAfiliado.Celular = telefono
                    ClsInformacionAfiliado.CorreoElectronico = ""
                    ClsInformacionAfiliado.IdTrabajador = NumeroIdentificacion
                    ClsInformacionAfiliado.TidTrabajador = TipoDocumentoTra
                    ClsInformacionAfiliado.NombreCompletoTrabajador = FilaAux("priape") & " " & FilaAux("segape") & " " & FilaAux("prinom") & " " & FilaAux("segnom") & ""

                    InfoAfiliado.Add(ClsInformacionAfiliado)

                Next
            Next

        Catch ex As Exception
            Throw (New Exception(ex.Message))
        End Try
    End Sub

    Private Sub ConsultaEmpresa(TipoDoc As String, ByRef NumeroIdentificacion As String, ByRef InfoEmpresa As ArrayList, ByRef ClaseParentescoAfil As String)

        Dim ClsInformacionEmpresa As InformacionEmpresa
        Dim SqlConsulta As String
        Dim DataSetEmpresa As New DataSet
        Dim DataSetAfiliado As New DataSet
        Dim DataSetValores As New DataSet
        Dim Fila As DataRow
        Dim nit, salario, fecsal, fecset As String
        Dim fecpag, valcon, valor, cant As String
        Dim TipoDocumento, estado As String
        Dim fechaDate As Date
        Dim fechaFormat As String

        Try
            nit = ""
            salario = ""
            fecsal = ""
            fecset = ""

            Select Case TipoDoc.ToUpper
                Case "CC"
                    TipoDoc = "1"
                Case "CE"
                    TipoDoc = "4"
                Case Else
                    TipoDoc = "1"
            End Select

            'datos adicionales del afiliado para la empresa
            'SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "'"
            SqlConsulta = "SELECT * FROM subsi15 WHERE cedtra = '" & NumeroIdentificacion & "' AND coddoc = '" & TipoDoc & "'"
            DataSetAfiliado = CsDatos.Consultar(SqlConsulta)

            For Each Fila In DataSetAfiliado.Tables(0).Rows
                nit = Fila("nit") & ""
                salario = Fila("salario") & ""
                fecsal = Fila("fecsal") & ""
                fecset = Fila("fecest") & ""
                Exit For
            Next

            'ultimos pagos de la empresa
            SqlConsulta = "select max(fecpag) as fecpag,valcon from subsi11 where nit  = '" & nit & "'"
            DataSetValores = CsDatos.Consultar(SqlConsulta)

            fecpag = ""
            valcon = ""

            For Each Fila In DataSetValores.Tables(0).Rows
                fecpag = Fila("fecpag") & ""
                valcon = Fila("valcon") & ""
                Exit For
            Next

            'ultimos pagos de la empresa
            DataSetValores.Tables.Clear()
            'SqlConsulta = "select max(periodo),valor from subsi09 where nit  = '" & nit & "'"
            SqlConsulta = "select A.periodo,A.valor from empresa.subsi09 A where A.periodo = " & _
                          "(select max(B.periodo) from empresa.subsi09 B where B.nit =  '" & nit & "') " & _
                          "AND A.nit =  '" & nit & "' LIMIT 1;"
            DataSetValores = CsDatos.Consultar(SqlConsulta)

            valor = ""

            For Each Fila In DataSetValores.Tables(0).Rows
                valor = Fila("valor") & ""
                Exit For
            Next

            'cantidad de afiliados de la empresa
            DataSetValores.Tables.Clear()
            SqlConsulta = "select count(*) as cant from subsi15 where nit  = '" & nit & "'"
            DataSetValores = CsDatos.Consultar(SqlConsulta)

            cant = ""

            For Each Fila In DataSetValores.Tables(0).Rows
                cant = Fila("cant") & ""
                Exit For
            Next

            'datos adicionales de la empresa
            SqlConsulta = "SELECT * FROM subsi02 WHERE nit = '" & nit & "'"
            DataSetEmpresa = CsDatos.Consultar(SqlConsulta)

            For Each Fila In DataSetEmpresa.Tables(0).Rows

                ClsInformacionEmpresa = New InformacionEmpresa()

                ClsInformacionEmpresa.IdEmpresa = nit

                TipoDocumento = ""
                TipoDocumento = Fila("coddoc") & ""
                Select Case TipoDocumento
                    Case "1"
                        TipoDocumento = "1"
                    Case "2"
                        TipoDocumento = "3"
                    Case "3"
                        TipoDocumento = "2"
                    Case "4"
                        TipoDocumento = "4"
                    Case "5"
                        TipoDocumento = "7"
                    Case "6"
                        TipoDocumento = "5"
                    Case Else
                        TipoDocumento = ""
                End Select

                ClsInformacionEmpresa.TidEmpresa = TipoDocumento
                ClsInformacionEmpresa.RazonSocialEmp = Fila("razsoc") & ""
                estado = Fila("estado") & ""
                Select Case estado
                    Case "A"
                        estado = "0"
                    Case "I"
                        estado = "1"
                    Case Else
                        estado = estado
                End Select
                ClsInformacionEmpresa.EstadoEmp = estado
                ClsInformacionEmpresa.CodigoDependencia = "0"
                ClsInformacionEmpresa.RazonSocialDep = ""

                Try
                    fechaDate = Fila("fecapr") & ""
                    fechaFormat = Format(fechaDate, "yyyyMMdd")
                    ClsInformacionEmpresa.FechaInicioEmp = fechaFormat
                Catch
                    ClsInformacionEmpresa.FechaInicioEmp = Fila("fecapr") & ""
                End Try

                ClsInformacionEmpresa.Salario = salario
                ClsInformacionEmpresa.Cargo = ""

                Try
                    fechaDate = fecsal
                    fechaFormat = Format(fechaDate, "yyyyMMdd")
                    ClsInformacionEmpresa.FechaInicioTR = fechaFormat
                Catch
                    ClsInformacionEmpresa.FechaInicioTR = fecsal
                End Try

                Try
                    fechaDate = fecset
                    fechaFormat = Format(fechaDate, "yyyyMMdd")
                    ClsInformacionEmpresa.FechafinTR = fechaFormat
                Catch
                    ClsInformacionEmpresa.FechafinTR = fecset
                End Try

                Try
                    fechaDate = fecpag
                    fechaFormat = Format(fechaDate, "yyyyMMdd")
                    ClsInformacionEmpresa.PUltimoAporte = fechaFormat
                Catch
                    ClsInformacionEmpresa.PUltimoAporte = fecpag
                End Try


                ClsInformacionEmpresa.VUltimoAporte = valcon
                ClsInformacionEmpresa.VUltimosubsidio = valor
                ClsInformacionEmpresa.CantTrAfiliados = cant
                ClsInformacionEmpresa.Ciudad = Fila("codciu") & ""
                ClsInformacionEmpresa.Direccion = Fila("direccion") & ""
                ClsInformacionEmpresa.Telefono = Fila("telefono") & ""
                ClsInformacionEmpresa.Extension = "0"

                ClaseParentescoAfil = Fila("calemp") & ""
                Select Case ClaseParentescoAfil
                    Case "E"
                        ClaseParentescoAfil = "TD"
                    Case "F"
                        ClaseParentescoAfil = "TI"
                End Select

                InfoEmpresa.Add(ClsInformacionEmpresa)
                Exit For

            Next

        Catch ex As Exception
            Throw (New Exception(ex.Message))
        End Try
    End Sub

    Private Function CrearXmlSalida(ByVal ArregloAfiliadosBeneficiarios As ArrayList, ArregloEmpresas As ArrayList, ByVal ClaseParentescoAfil As String) As XmlDocument

        Dim xmlDocumento As XmlDocument
        'Dim xmlNodoSoapBody As XmlNode
        Dim xmlNodoAfiliadoResult, xmlNodoAfiliadoResponse As XmlNode
        Dim xmlNodoInfo, xmlNodoDatos As XmlNode
        Dim ClaseInfo As InformacionAfiliado
        Dim ClaseInfoEmp As InformacionEmpresa
        Dim xmldecl As XmlDeclaration

        Try

            xmlDocumento = New XmlDocument()

            'xmlNodoSoapBody = xmlDocumento.CreateNode(XmlNodeType.Element, "Soap:Body", "")
            'xmlDocumento.AppendChild(xmlNodoSoapBody)

            xmlNodoAfiliadoResponse = xmlDocumento.CreateNode(XmlNodeType.Element, "TrabajadorInfoResponse", "")
            xmlDocumento.AppendChild(xmlNodoAfiliadoResponse)

            xmlNodoAfiliadoResult = xmlDocumento.CreateNode(XmlNodeType.Element, "DsGrupoFamiliar", "")

            For I As Integer = 0 To ArregloAfiliadosBeneficiarios.Count - 1

                ClaseInfo = ArregloAfiliadosBeneficiarios.Item(I)

                If I = 0 Then

                    xmlNodoInfo = xmlDocumento.CreateNode(XmlNodeType.Element, "Afiliado", "")

                    xmlNodoDatos = xmlDocumento.CreateElement("IdBeneficiario")
                    xmlNodoDatos.InnerText = ClaseInfo.IdBeneficiario
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("TidBeneficiario")
                    xmlNodoDatos.InnerText = ClaseInfo.TidBeneficiario
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("AlfaDocumento")
                    xmlNodoDatos.InnerText = ClaseInfo.AlfaDocumento
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("PrimerApellido")
                    xmlNodoDatos.InnerText = ClaseInfo.PrimerApellido
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("SegundoApellido")
                    xmlNodoDatos.InnerText = ClaseInfo.SegundoApellido
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("PrimerNombre")
                    xmlNodoDatos.InnerText = ClaseInfo.PrimerNombre
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("SegundoNombre")
                    xmlNodoDatos.InnerText = ClaseInfo.SegundoNombre
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("FechaNacimiento")
                    xmlNodoDatos.InnerText = ClaseInfo.FechaNacimiento
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("EstadoCivil")
                    xmlNodoDatos.InnerText = ClaseInfo.EstadoCivil
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Genero")
                    xmlNodoDatos.InnerText = ClaseInfo.Genero
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Estado")
                    xmlNodoDatos.InnerText = ClaseInfo.Estado
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Categoria")
                    xmlNodoDatos.InnerText = ClaseInfo.Categoria
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("ClaseParentesco")
                    xmlNodoDatos.InnerText = ClaseParentescoAfil
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("NumeroTarjeta")
                    xmlNodoDatos.InnerText = ClaseInfo.NumeroTarjeta
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Ciudad")
                    xmlNodoDatos.InnerText = ClaseInfo.Ciudad
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Direccion")
                    xmlNodoDatos.InnerText = ClaseInfo.Direccion
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Telefono")
                    xmlNodoDatos.InnerText = ClaseInfo.Telefono
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Celular")
                    xmlNodoDatos.InnerText = ClaseInfo.Celular
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("CorreoElectronico")
                    xmlNodoDatos.InnerText = ClaseInfo.CorreoElectronico
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("IdTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.IdTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("TidTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.TidTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("NombreCompletoTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.NombreCompletoTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoAfiliadoResult.AppendChild(xmlNodoInfo)

                End If

                If I > 0 Then

                    xmlNodoInfo = xmlDocumento.CreateNode(XmlNodeType.Element, "Benficiario", "")

                    xmlNodoDatos = xmlDocumento.CreateElement("IdBeneficiario")
                    xmlNodoDatos.InnerText = ClaseInfo.IdBeneficiario
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("TidBeneficiario")
                    xmlNodoDatos.InnerText = ClaseInfo.TidBeneficiario
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("AlfaDocumento")
                    xmlNodoDatos.InnerText = ClaseInfo.AlfaDocumento
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("PrimerApellido")
                    xmlNodoDatos.InnerText = ClaseInfo.PrimerApellido
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("SegundoApellido")
                    xmlNodoDatos.InnerText = ClaseInfo.SegundoApellido
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("PrimerNombre")
                    xmlNodoDatos.InnerText = ClaseInfo.PrimerNombre
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("SegundoNombre")
                    xmlNodoDatos.InnerText = ClaseInfo.SegundoNombre
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("FechaNacimiento")
                    xmlNodoDatos.InnerText = ClaseInfo.FechaNacimiento
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("EstadoCivil")
                    xmlNodoDatos.InnerText = ClaseInfo.EstadoCivil
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Genero")
                    xmlNodoDatos.InnerText = ClaseInfo.Genero
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Estado")
                    xmlNodoDatos.InnerText = ClaseInfo.Estado
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Categoria")
                    xmlNodoDatos.InnerText = ClaseInfo.Categoria
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("ClaseParentesco")
                    xmlNodoDatos.InnerText = ClaseInfo.ClaseParentesco
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("NumeroTarjeta")
                    xmlNodoDatos.InnerText = ClaseInfo.NumeroTarjeta
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Ciudad")
                    xmlNodoDatos.InnerText = ClaseInfo.Ciudad
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Direccion")
                    xmlNodoDatos.InnerText = ClaseInfo.Direccion
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Telefono")
                    xmlNodoDatos.InnerText = ClaseInfo.Telefono
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("Celular")
                    xmlNodoDatos.InnerText = ClaseInfo.Celular
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("CorreoElectronico")
                    xmlNodoDatos.InnerText = ClaseInfo.CorreoElectronico
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("IdTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.IdTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("TidTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.TidTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoDatos = xmlDocumento.CreateElement("NombreCompletoTrabajador")
                    xmlNodoDatos.InnerText = ClaseInfo.NombreCompletoTrabajador
                    xmlNodoInfo.AppendChild(xmlNodoDatos)

                    xmlNodoAfiliadoResult.AppendChild(xmlNodoInfo)
                End If

            Next

            For I As Integer = 0 To ArregloEmpresas.Count - 1

                ClaseInfoEmp = ArregloEmpresas.Item(I)

                xmlNodoInfo = xmlDocumento.CreateNode(XmlNodeType.Element, "Empresa", "")

                xmlNodoDatos = xmlDocumento.CreateElement("IdEmpresa")
                xmlNodoDatos.InnerText = ClaseInfoEmp.IdEmpresa
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("TidEmpresa")
                xmlNodoDatos.InnerText = ClaseInfoEmp.TidEmpresa
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("RazonSocialEmp")
                xmlNodoDatos.InnerText = ClaseInfoEmp.RazonSocialEmp
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("EstadoEmp")
                xmlNodoDatos.InnerText = ClaseInfoEmp.EstadoEmp
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("CodigoDependencia")
                xmlNodoDatos.InnerText = ClaseInfoEmp.CodigoDependencia
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("RazonSocialDep")
                xmlNodoDatos.InnerText = ClaseInfoEmp.RazonSocialDep
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("FechaInicioEmp")
                xmlNodoDatos.InnerText = ClaseInfoEmp.FechaInicioEmp
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Salario")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Salario
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Cargo")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Cargo
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("FechaInicioTR")
                xmlNodoDatos.InnerText = ClaseInfoEmp.FechaInicioTR
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("FechafinTR")
                xmlNodoDatos.InnerText = ClaseInfoEmp.FechafinTR
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("PUltimoAporte")
                xmlNodoDatos.InnerText = ClaseInfoEmp.PUltimoAporte
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("VUltimoAporte")
                xmlNodoDatos.InnerText = ClaseInfoEmp.VUltimoAporte
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("VUltimosubsidio")
                xmlNodoDatos.InnerText = ClaseInfoEmp.VUltimosubsidio
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("CantTrAfiliados")
                xmlNodoDatos.InnerText = ClaseInfoEmp.CantTrAfiliados
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Ciudad")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Ciudad
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Direccion")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Direccion
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Telefono")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Telefono
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoDatos = xmlDocumento.CreateElement("Extension")
                xmlNodoDatos.InnerText = ClaseInfoEmp.Extension
                xmlNodoInfo.AppendChild(xmlNodoDatos)

                xmlNodoAfiliadoResult.AppendChild(xmlNodoInfo)

            Next

            xmlNodoAfiliadoResponse.AppendChild(xmlNodoAfiliadoResult)

            'xmlNodoSoapBody.AppendChild(xmlNodoAfiliadoResponse)

            xmldecl = xmlDocumento.CreateXmlDeclaration("1.0", Nothing, Nothing)
            xmldecl.Encoding = "UTF-8"
            xmldecl.Standalone = "yes"

            Dim root As XmlElement = xmlDocumento.DocumentElement
            xmlDocumento.InsertBefore(xmldecl, root)

            CrearXmlSalida = xmlDocumento 'xmlDocumento.OuterXml

            '***xmlDocumento.Save(RutaSalida & "Consulta_" & CodCaja & "_" & FechaIngreso & "_" & HoraIngreso & ".xml")

        Catch ex As Exception
            CsError.MensajeError = ex.Message
            Return XmlError(CsError)
        End Try
    End Function


    <WebMethod()> _
    Private Function XmlError(CsError As ErrorWs) As XmlDocument

        Dim xmlDocumento As XmlDocument
        Dim xmlNodo As XmlNode

        Try

            xmlDocumento = New XmlDocument()

            xmlNodo = xmlDocumento.CreateElement("ErrorWs")
            xmlDocumento.AppendChild(xmlNodo)

            xmlNodo = xmlDocumento.CreateElement("MensajeError")
            xmlNodo.InnerText = CsError.MensajeError 
            xmlDocumento.DocumentElement.AppendChild(xmlNodo)

            XmlError = xmlDocumento 'xmlDocumento.OuterXml

        Catch ex As Exception
            Throw (New Exception(ex.Message))
        End Try

    End Function

End Class