Imports Inventor
Imports System.IO

Imports IOPath = System.IO.Path

''' <summary>
''' Orquesta el proceso completo de renombrado de proyecto:
''' 1. Reemplaza referencias en IDWs (ProjectRenamer)
''' 2. Espera que Inventor suelte los locks
''' 3. Renombra físicamente los archivos .idw (IdwPhysicalRenamer)
''' </summary>
Public Class ProjectRenameWorkflow

    Private ReadOnly _app As Inventor.Application
    Private ReadOnly _asmDoc As AssemblyDocument

    Public Sub New(app As Inventor.Application, asmDoc As AssemblyDocument)
        _app = app
        _asmDoc = asmDoc
    End Sub

    ''' <summary>
    ''' Ejecuta el flujo completo de renombrado.
    ''' </summary>
    ''' <returns>Ruta de la carpeta del proyecto nuevo</returns>
    Public Function Execute() As String

        ' === GUARDAR DATOS ANTES de que ProjectRenamer cierre todo ===
        Dim oldAsmName As String = IOPath.GetFileName(_asmDoc.FullFileName)

        ' === PASO 1: Renombrar proyecto y reemplazar referencias ===
        Dim renamer As New ProjectRenamer(_app, _asmDoc)
        Dim targetFolder As String = renamer.Execute()

        ' === PASO 2: Recolectar datos para el renombrado físico ===
        Dim newProjectName As String = IOPath.GetFileName(targetFolder)

        ' === PASO 3: Esperar a que Inventor suelte los locks ===
        For i As Integer = 1 To 3
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            System.Threading.Thread.Sleep(2000)  ' 2 segundos x 3 = 6 segundos total
        Next

        ' === PASO 4: Renombrar físicamente los IDWs ===
        Dim idwRenamer As New IdwPhysicalRenamer()
        idwRenamer.Execute(targetFolder, newProjectName, oldAsmName)

        Return targetFolder

    End Function

End Class