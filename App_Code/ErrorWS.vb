Imports Microsoft.VisualBasic

Public Class ErrorWs

    Private _MensajeError As String

    Sub New()
        Me._MensajeError = String.Empty
    End Sub

    Public Property MensajeError() As String
        Get
            Return _MensajeError
        End Get
        Set(ByVal value As String)
            _MensajeError = value
        End Set
    End Property

End Class
