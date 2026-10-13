Imports System.IO
Imports System.Windows.Forms
Imports System.Xml.Serialization

''' <summary>
''' Configuración global del Add-In RuleManager V2.
''' Serializada automáticamente en XML bajo LocalApplicationData.
''' </summary>
Public Class AddinSettings

    ' ============================================
    ' 1. GENERAL / INICIO
    ' ============================================
    Public Property CrearCarpetasProyecto As Boolean = True
    Public Property SincronizarLibreriasAuto As Boolean = True
    Public Property HabilitarSilentOperation As Boolean = True
    Public Property RutaLogs As String = "C:\Temp\RuleManager.log"

    ' ============================================
    ' 2. ENSAMBLAJE (IAM)
    ' ============================================
    Public Property OmitirContentCenterCOM As Boolean = True
    Public Property DetectarSimetriasAuto As Boolean = True
    Public Property ActualizarCantidadUsada As Boolean = True
    Public Property RutaPlantillasReglas As String = "D:\Bibliotecas\Ilogic\Plantillas para reglas"
    Public Property AbrirCarpetaAlTerminar As Boolean = True

    ' ============================================
    ' 3. PLANOS (IDW)
    ' ============================================
    Public Property FormatoPdfPredeterminado As String = "B3"
    Public Property OrdenarHojasAuto As Boolean = True
    Public Property PlantillaLdmExcel As String = "REQUERIMIENTO DE MATERIALESV5.xlsx"

    ' ============================================
    ' 4. PIEZA (IPT) & SELECCIÓN BASE
    ' ============================================
    Public Property CalcularCuboCorteAuto As Boolean = True
    Public Property ExtraerParametrosChapa As Boolean = True
    Public Property PlantillaPiezaDefault As String = "PIEZA DB 2.ipt"
    Public Property ActivarSeleccionBaseEnPlantilla As Boolean = True

    ' Catálogos persistentes inyectados
    Public Property EspesoresChapaInyectados As New List(Of String)
    Public Property EspesoresPlanchaInyectados As New List(Of String)
    Public Property DiametrosRedondoInyectados As New List(Of String)
    Public Property DimensionesCuadradoInyectados As New List(Of String)
    Public Property MetalesInyectados As New List(Of String)
    Public Property PlasticosInyectados As New List(Of String)

    ' ============================================
    ' 4.1 MEMORIA DE ÚLTIMA CREACIÓN DE PIEZA
    ' ============================================
    Public Property PrecargarGuardarDatosPieza As Boolean = True
    Public Property UltimoTipoPieza As String = "CHAPA"

    ' Datos memorizados para CHAPA
    Public Property UltimaChapaUnidadesMm As Boolean = True
    Public Property UltimaChapaEspesor As String = "3mm"
    Public Property UltimaChapaMaterial As String = "Acero ASTM A36"

    ' Datos memorizados para PLANCHA
    Public Property UltimaPlanchaEspesor As String = "10mm"
    Public Property UltimaPlanchaMaterial As String = "APM - NATURAL"

    ' Datos memorizados para CUADRADO
    Public Property UltimoCuadradoSubTipo As String = "TREFILADO"
    Public Property UltimoCuadradoDimension As String = "1/2"" (12.70mm)"
    Public Property UltimoCuadradoMaterial As String = "Acero SAE 1045"

    ' Datos memorizados para REDONDO
    Public Property UltimoRedondoSubTipo As String = "TREFILADO"
    Public Property UltimoRedondoDimension As String = "25mm"
    Public Property UltimoRedondoMaterial As String = "AISI 304"

    ' Datos memorizados para GENERICO y GUIA
    Public Property UltimoGenericoMaterial As String = "Genérico"
    Public Property UltimaGuiaMaterial As String = "Genérico"

    ' ============================================
    ' 5. PROYECTO & ARCHIVOS
    ' ============================================
    Public Property ExcluirOldVersions As Boolean = True
    Public Property PreservarCarpetasEspeciales As Boolean = True
    Public Property ProtegerRoutedSystems As Boolean = True
    Public Property EnrutarMulticorteDxf As Boolean = True

    ' ============================================
    ' 6. ODOO ERP
    ' ============================================
    Public Property TipoBomPredeterminado As String = "Kit"
    Public Property PrefijoExternalId As String = "PRJ"

    ' ============================================
    ' UBICACIÓN DEL ARCHIVO DE CONFIGURACIÓN
    ' ============================================
    Private Shared ReadOnly Property SettingsPath As String
        Get
            Dim configDir As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RuleManagerV2")

            If Not Directory.Exists(configDir) Then
                Directory.CreateDirectory(configDir)
            End If

            Return Path.Combine(configDir, "RuleManagerV2.config.xml")
        End Get
    End Property

    ' ============================================
    ' INSTANCIA GLOBAL (SINGLETON)
    ' ============================================
    Private Shared _current As AddinSettings = Nothing

    Public Shared ReadOnly Property Current As AddinSettings
        Get
            If _current Is Nothing Then
                _current = Load()
            End If
            Return _current
        End Get
    End Property

    ' ============================================
    ' CARGAR / GUARDAR
    ' ============================================
    Public Shared Function Load() As AddinSettings
        Try
            If File.Exists(SettingsPath) Then
                Dim xs As New XmlSerializer(GetType(AddinSettings))
                Using sr As New StreamReader(SettingsPath)
                    Dim cfg = DirectCast(xs.Deserialize(sr), AddinSettings)
                    If cfg IsNot Nothing Then Return cfg
                End Using
            End If
        Catch
            ' Si el archivo está corrupto o desactualizado, arrancar con defaults
        End Try
        Return New AddinSettings()
    End Function

    Public Shared Sub Save()
        Try
            Dim xs As New XmlSerializer(GetType(AddinSettings))
            Using sw As New StreamWriter(SettingsPath)
                xs.Serialize(sw, Current)
            End Using
        Catch ex As Exception
            MessageBox.Show(
                "No se pudo guardar la configuración:" & vbCrLf & ex.Message,
                "RuleManager - Opciones",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        End Try
    End Sub

End Class
