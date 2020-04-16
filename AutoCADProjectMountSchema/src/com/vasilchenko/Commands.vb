Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko
    Public Class Commands

        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedTrans")>
        Public Shared Function acedTrans(ByVal point As Double(), ByVal fromRb As IntPtr, ByVal toRb As IntPtr, ByVal disp As Integer, ByVal result As Double()) As Integer
        End Function

        <CommandMethod("ASU_Project_Mount_Builder_DEBUG", CommandFlags.Session)>
        Public Shared Sub Builder()

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            If Application.GetSystemVariable("MIRRTEXT") = "1" Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("MIRRTEXT variable set to 0")
                Application.SetSystemVariable("MIRRTEXT", 0)
            End If

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufLocationSelector
                Try
                    objForm.ShowDialog()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using

        End Sub

    End Class
End Namespace