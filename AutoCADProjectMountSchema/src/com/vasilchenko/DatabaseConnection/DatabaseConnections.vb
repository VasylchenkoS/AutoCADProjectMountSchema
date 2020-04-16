Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Autodesk.Electrical.Project


Namespace com.vasilchenko.DatabaseConnection
    Module DatabaseConnections
        Friend Function GetSqlDataTable(strSqlQuery As String, strSqlConnectionParameters As String) As DataTable
            Using objSqlDbConnection As New SqlConnection(strSqlConnectionParameters)
                Try
                    objSqlDbConnection.Open()
                    Dim objDataTable As New DataTable
                    Dim objSqlDbCommand = New SqlCommand(strSqlQuery, objSqlDbConnection)
                    Dim maxCode = objSqlDbCommand.ExecuteScalar
                    If maxCode Is Nothing Then
                        Return Nothing
                    Else
                        objDataTable.Load(objSqlDbCommand.ExecuteReader)
                        Return objDataTable
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return Nothing
                End Try
            End Using
        End Function

        Friend Function GetOleDataTable(strSqlQuery As String) As DataTable
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            If Not File.Exists(strConstProjectDatabasePath) Then
                MsgBox(
                    "Source file not found. File way: " & strConstProjectDatabasePath &
                    " Please open some project file and repeat.", vbCritical, "File Not Found")
                Throw New ArgumentNullException($"File Not Found", innerException:=New Exception())
            End If

            Dim strSqlConnectionParameters As String

            strSqlConnectionParameters = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strConstProjectDatabasePath & ";"
            Using objOleDbConnection As New OleDbConnection(strSqlConnectionParameters)
                Try
                    objOleDbConnection.Open()
                    Dim objOleDbCommand = New OleDbCommand(strSqlQuery, objOleDbConnection)
                    Dim objDataTable As New DataTable
                    objDataTable.Load(objOleDbCommand.ExecuteReader)

                    Return objDataTable
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New ArgumentNullException
                End Try
            End Using
        End Function
    End Module
End Namespace
