Imports System
Imports Inventor

''' <summary>
''' Representa la información detallada de una ocurrencia o pieza simétrica (espejada)
''' identificada en el ensamblaje por SearchSimetriasEngine mediante análisis geométrico de contornos.
''' Compatible con SearchSimetriasDialog y todas las vistas de revisión de RuleManager V2.
''' </summary>
Public Class SimetriaItemInfo

    ''' <summary>
    ''' Indica si el usuario seleccionó aplicar el cambio de iProperties a este componente en el diálogo.
    ''' </summary>
    Public Property AplicarCambio As Boolean = True

    ''' <summary>
    ''' Ocurrencia del componente en el ensamblaje.
    ''' </summary>
    Public Property Ocurrencia As ComponentOccurrence
    Public Property Occurrence As ComponentOccurrence
        Get
            Return Ocurrencia
        End Get
        Set(value As ComponentOccurrence)
            Ocurrencia = value
        End Set
    End Property

    ''' <summary>
    ''' Nombre de la ocurrencia en el árbol del ensamblaje.
    ''' </summary>
    Public Property Nombre As String = ""
    Public Property Name As String
        Get
            Return Nombre
        End Get
        Set(value As String)
            Nombre = value
        End Set
    End Property

    ''' <summary>
    ''' Alias de nombre de ocurrencia esperado por SearchSimetriasDialog.
    ''' </summary>
    Public Property OccurrenceName As String
        Get
            If Not String.IsNullOrEmpty(Nombre) Then Return Nombre
            If Ocurrencia IsNot Nothing Then Return Ocurrencia.Name
            Return ""
        End Get
        Set(value As String)
            Nombre = value
        End Set
    End Property

    ''' <summary>
    ''' Nombre para mostrar del archivo de pieza o subensamblaje.
    ''' </summary>
    Public Property NombreDocumento As String = ""
    Public Property DisplayName As String
        Get
            Return NombreDocumento
        End Get
        Set(value As String)
            NombreDocumento = value
        End Set
    End Property

    ''' <summary>
    ''' Ruta absoluta del archivo en disco.
    ''' </summary>
    Public Property RutaArchivo As String = ""
    Public Property FullFileName As String
        Get
            Return RutaArchivo
        End Get
        Set(value As String)
            RutaArchivo = value
        End Set
    End Property

    ''' <summary>
    ''' Documento de la pieza simétrica.
    ''' </summary>
    Public Property Documento As Document
    Public Property Document As Document
        Get
            Return Documento
        End Get
        Set(value As Document)
            Documento = value
        End Set
    End Property

    ''' <summary>
    ''' Documento original del cual se deriva o del cual es reflejo geométrico.
    ''' </summary>
    Public Property DocumentoOriginal As Document
    Public Property OriginalDocument As Document
        Get
            Return DocumentoOriginal
        End Get
        Set(value As Document)
            DocumentoOriginal = value
        End Set
    End Property

    ''' <summary>
    ''' Part Number original restaurado para agrupar con la derecha en listas de materiales (BOM).
    ''' </summary>
    Public Property PartNumberOriginal As String = ""
    Public Property OriginalPartNumber As String
        Get
            Return PartNumberOriginal
        End Get
        Set(value As String)
            PartNumberOriginal = value
        End Set
    End Property

    ''' <summary>
    ''' Alias de Part Number original sugerido esperado por SearchSimetriasDialog.
    ''' </summary>
    Public Property PartNumberOriginalSugerido As String
        Get
            Return PartNumberOriginal
        End Get
        Set(value As String)
            PartNumberOriginal = value
        End Set
    End Property

    ''' <summary>
    ''' Part Number actual de la pieza antes de normalizar.
    ''' </summary>
    Public Property PartNumberActual As String = ""
    Public Property CurrentPartNumber As String
        Get
            Return PartNumberActual
        End Get
        Set(value As String)
            PartNumberActual = value
        End Set
    End Property

    ''' <summary>
    ''' Método geométrico o estructural mediante el cual fue detectada la simetría.
    ''' </summary>
    Public Property MetodoDeteccion As String = ""
    Public Property DetectionMethod As String
        Get
            Return MetodoDeteccion
        End Get
        Set(value As String)
            MetodoDeteccion = value
        End Set
    End Property

    ''' <summary>
    ''' Texto del método de detección para SearchSimetriasDialog.
    ''' </summary>
    Public Property MetodoDeteccionTexto As String
        Get
            Return MetodoDeteccion
        End Get
        Set(value As String)
            MetodoDeteccion = value
        End Set
    End Property

    ''' <summary>
    ''' Detalle adicional del método de detección para SearchSimetriasDialog.
    ''' </summary>
    Public Property DetalleMetodo As String
        Get
            Return MetodoDeteccion
        End Get
        Set(value As String)
            MetodoDeteccion = value
        End Set
    End Property

    ''' <summary>
    ''' Indica si la pieza es quiral (carece de simetría interna propia y requiere pieza opuesta).
    ''' </summary>
    Public Property EsQuiral As Boolean = True
    Public Property IsChiral As Boolean
        Get
            Return EsQuiral
        End Get
        Set(value As Boolean)
            EsQuiral = value
        End Set
    End Property

    ''' <summary>
    ''' Indica si ya fue procesada la asignación de propiedades iProperty.
    ''' </summary>
    Public Property Procesado As Boolean = False
    Public Property IsProcessed As Boolean
        Get
            Return Procesado
        End Get
        Set(value As Boolean)
            Procesado = value
        End Set
    End Property

    ''' <summary>
    ''' Descripción normalizada para BOM (SIMETRIA).
    ''' </summary>
    Public Property Descripcion As String = "SIMETRIA"
    Public Property Description As String
        Get
            Return Descripcion
        End Get
        Set(value As String)
            Descripcion = value
        End Set
    End Property

    ''' <summary>
    ''' Alias de nueva descripción esperado por SearchSimetriasDialog.
    ''' </summary>
    Public Property DescripcionNueva As String
        Get
            Return Descripcion
        End Get
        Set(value As String)
            Descripcion = value
        End Set
    End Property

    Public Sub New()
        AplicarCambio = True
        Descripcion = "SIMETRIA"
    End Sub

    Public Sub New(occ As ComponentOccurrence)
        Me.New()
        Ocurrencia = occ
        If occ IsNot Nothing Then
            Nombre = occ.Name
            Try
                If occ.Definition IsNot Nothing AndAlso occ.Definition.Document IsNot Nothing Then
                    Documento = DirectCast(occ.Definition.Document, Document)
                    NombreDocumento = Documento.DisplayName
                    RutaArchivo = Documento.FullFileName
                End If
            Catch
            End Try
        End If
    End Sub

    Public Sub New(occ As ComponentOccurrence, metodo As String, origDoc As Document, origPN As String)
        Me.New(occ)
        MetodoDeteccion = metodo
        DocumentoOriginal = origDoc
        PartNumberOriginal = origPN
    End Sub

    Public Overrides Function ToString() As String
        If Not String.IsNullOrEmpty(PartNumberOriginal) Then
            Return Nombre & " [Original PN: " & PartNumberOriginal & "] - " & MetodoDeteccion
        End If
        Return Nombre & " - " & MetodoDeteccion
    End Function

End Class
