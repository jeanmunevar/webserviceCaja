
Imports System.Data.OleDb
Imports System.Data

Imports MySql.Data.MySqlClient

Module Module1

    Public CadenaDeConexion, RutaSalida, CodCaja As String
    Public FechaIngreso, HoraIngreso As String

    Public myConnection As New MySqlConnection
    Public myTransaccion As MySqlTransaction

End Module
