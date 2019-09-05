Imports Microsoft.VisualBasic
Imports MySql.Data.MySqlClient
Imports System.Data

Public Class Datos

    Public Sub New()

    End Sub

    Function Conectar_BD(ByVal CadenaConexion As String) As Boolean
        Try

            If myConnection.ConnectionString = "" Then
                myConnection.ConnectionString = CadenaConexion
            End If

            If myConnection.State = ConnectionState.Open Then
                Return True
            End If

            myConnection.Open()

            Conectar_BD = True

        Catch ex As Exception
            Conectar_BD = False
            Throw (New Exception(ex.Message & "-" & "Error al conectarse a la base de datos"))
        End Try
    End Function

    Public Function CerrarBaseDatos() As Boolean
        Try
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
                Return True
            Else
                Return True
            End If

        Catch ex As Exception
            CerrarBaseDatos = False
        End Try
    End Function


    '>>>
    Public Function IniciarTransaccion() As Boolean
        Try
            If myConnection.State = ConnectionState.Open Then
                myTransaccion = myConnection.BeginTransaction()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            IniciarTransaccion = False
        End Try
    End Function

    Public Function FinalizarTransaccion() As Boolean
        Try
            If myConnection.State = ConnectionState.Open Then
                myTransaccion.Commit()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            FinalizarTransaccion = False
        End Try
    End Function

    Public Function InterrumpirTransaccion() As Boolean
        Try
            If myConnection.State = ConnectionState.Open Then
                myTransaccion.Rollback()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            InterrumpirTransaccion = False
        End Try
    End Function
    '>>>

    Public Function Consultar(ByVal Consulta As String) As DataSet
        Dim MiDataSet As New DataSet
        Dim Obj_Dadp As MySqlDataAdapter
        Try
            Obj_Dadp = New MySqlDataAdapter(Consulta, myConnection)
            Obj_Dadp.SelectCommand.Transaction = myTransaccion
            Obj_Dadp.SelectCommand.CommandTimeout = 240
            Obj_Dadp.Fill(MiDataSet)
            Return MiDataSet
            '>>>
            Obj_Dadp.Dispose()
            '>>>
        Catch ex As Exception
            Consultar = Nothing
            Throw (New Exception(ex.Message))
        End Try
    End Function

End Class
