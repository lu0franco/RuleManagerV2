Imports System.IO
Imports System.Windows.Forms
Imports System.Xml.Serialization

Public Class AddinSettings

    ' ============================================
    ' OPCIONES DE COMPORTAMIENTO
    ' (agregar nuevas propiedades aqui; el XML
    '  las guarda automaticamente)
    ' ============================================

    ' Crear la estructura de carpetas del proyecto
    ' (01 - DIBUJOS, 02 - PLANOS, etc.) al activar
    ' o cambiar de proyecto en Inventor
    Public Property CrearCarpetasProyecto As Boolean = True


    ' ============================================
    ' UBICACION DEL ARCHIVO DE CONFIGURACION
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
    ' INSTANCIA GLOBAL (singleton simple)
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

    Private Shared Function Load() As AddinSettings
        Try
            If File.Exists(SettingsPath) Then
                Dim xs As New XmlSerializer(GetType(AddinSettings))
                Using sr As New StreamReader(SettingsPath)
                    Dim cfg = DirectCast(xs.Deserialize(sr), AddinSettings)
                    If cfg IsNot Nothing Then Return cfg
                End Using
            End If
        Catch
            ' Si el archivo esta corrupto, arrancar con defaults
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
                "No se pudo guardar la configuracion:" & vbCrLf & ex.Message,
                "RuleManager - Opciones",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        End Try
    End Sub

End Class
