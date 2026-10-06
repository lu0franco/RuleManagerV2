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

    ' ============================================
    ' 3. PLANOS (IDW)
    ' ============================================
    Public Property FormatoPdfPredeterminado As String = "B3"
    Public Property OrdenarHojasAuto As Boolean = True
    Public Property PlantillaLdmExcel As String = "REQUERIMIENTO DE MATERIALESV5.xlsx"

    ' ============================================
    ' 4. PIEZA (IPT)
    ' ============================================
    Public Property CalcularCuboCorteAuto As Boolean = True
    Public Property ExtraerParametrosChapa As Boolean = True

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
