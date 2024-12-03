Imports Microsoft.Data.SqlClient

Public Class SQLParameterTest

    Dim m_bRequestRights As Boolean = False
    Dim m_sTempSql As String = ""
    Dim m_dbkey As String = "12"

    Public Function RequestsForUser(
                                   Optional ByVal sSearchName As String = "",
                                   Optional ByVal sLevel As String = "",
                                   Optional ByVal sAdditionalWhere As String = "") As String

        Dim sDynamicWhere As String = "WHERE 1 = 1 "

        If sSearchName <> "" Then
            sDynamicWhere &= "AND R.Name like '%" & sSearchName.Replace("'", "''") & "%' "
        End If

        If sLevel <> "" Then
            sDynamicWhere &= "AND R.Level = '" & sLevel & "' "
        End If

        If m_bRequestRights Then
            m_sTempSql =
                "SELECT R.DBKEY, R.NAME, R.LEVEL1, R.NAME + '(' + R.LEVEL1 + ')' AS ListBoxTest " &
                "FROM Request R " &
                sDynamicWhere &
                sAdditionalWhere & " " &
                "ORDER BY R.NAME"
        Else
            m_sTempSql =
                "SELECT R.DBKEY, R.NAME, R.LEVEL1, R.NAME + '(' + R.LEVEL1 + ')' AS ListBoxTest " &
                "FROM Request R INNER JOIN USERACCESSREQ UR ON R.dbkey = UR.RECORD_DBKEY " &
                sDynamicWhere &
                "AND UR.USER_DBKEY = " & m_dbkey & " " &
                sAdditionalWhere & " " &
                "ORDER BY R.NAME"
        End If

        Return m_sTempSql

    End Function


    Public Function RequestsForUserWithParameters(
                                   Optional ByVal sSearchName As String = "",
                                   Optional ByVal sLevel As String = "",
                                   Optional ByVal sAdditionalWhere As String = "") As String

        Dim parameters As New List(Of SqlParameter)

        Dim sDynamicWhere As String = "WHERE 1 = 1 "

        If sSearchName <> "" Then
            sSearchName = sSearchName.Replace("'", "''")
            sDynamicWhere &= "AND R.Name like '%@sSearchName%' "
            parameters.Add(New SqlParameter("@sSearchName", sSearchName))
        End If

        If sLevel <> "" Then
            sDynamicWhere &= "AND R.Level = '@sLevel' "
            parameters.Add(New SqlParameter("@sLevel", sLevel))
        End If

        If m_bRequestRights Then
            m_sTempSql =
                "SELECT R.DBKEY, R.NAME, R.LEVEL1, R.NAME + '(' + R.LEVEL1 + ')' AS ListBoxTest " &
                "FROM Request R " &
                sDynamicWhere &
                sAdditionalWhere & " " &
                "ORDER BY R.NAME"
        Else
            m_sTempSql =
                "SELECT R.DBKEY, R.NAME, R.LEVEL1, R.NAME + '(' + R.LEVEL1 + ')' AS ListBoxTest " &
                "FROM Request R INNER JOIN USERACCESSREQ UR ON R.dbkey = UR.RECORD_DBKEY " &
                sDynamicWhere &
                "AND UR.USER_DBKEY = " & m_dbkey & " " &
                sAdditionalWhere & " " &
                "ORDER BY R.NAME"
        End If

        Return m_sTempSql

    End Function



End Class
