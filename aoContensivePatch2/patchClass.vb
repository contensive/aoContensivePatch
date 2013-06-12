
Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoContensivePatch2
    '

    Public Class patchClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            '
            Const patchVersion = "4.1.003"
            '
            Dim s As String = ""
            Dim lastPatchVersion As String = ""
            Dim buildVersion As String = ""
            Dim sectionCid As Integer = 0
            Dim ruleCID As Integer = 0
            Dim mustLoadCDef As Boolean
            Dim sql As String = ""
            '
            Try
                Call CP.Utils.AppendLogFile("Running Patch for application " & CP.Site.Name)
                lastPatchVersion = CP.Site.GetProperty("PatchRevision", "0")
                If lastPatchVersion = "0" Then
                    lastPatchVersion = "1.0.000"
                End If
                '
                buildVersion = CP.Site.GetProperty("BuildVersion", "0")
                If buildVersion <= "0" Then
                    Call CP.Utils.AppendLogFile("Bad Build version [" & buildVersion & "]")
                Else
                    '
                    '---------------------------------------------------------------------------------
                    ' fix Dynamic Menu Section Rules.SectionID to be lookup into site sections
                    '---------------------------------------------------------------------------------
                    '
                    If (lastPatchVersion < "1.1.025") And (buildVersion < "4.1.487") Then
                        sectionCid = CP.Content.GetID("Site Sections")
                        ruleCID = CP.Content.GetID("Dynamic Menu Section Rules")
                        Call CP.Db.ExecuteSQL("update ccfields set lookupContentId=" & sectionCid & " where name='sectionid' and contentid=" & ruleCID)
                        mustLoadCDef = True
                    End If
                    '
                    '---------------------------------------------------------------------------------
                    ' update shopping cart property to update customers from orders
                    '---------------------------------------------------------------------------------
                    '
                    If (lastPatchVersion < "4.1.002") And (CP.Site.Name.ToLower <> "aatb") Then
                        Call CP.Site.SetProperty("orderUpdateMemberContactFromBilling", "1")
                    End If
                    '
                    '---------------------------------------------------------------------------------
                    ' 4.1.003 - remove old order reports
                    '---------------------------------------------------------------------------------
                    '
                    If (lastPatchVersion < "4.1.003") Then
                        sql = "delete from ccMenuEntries where name = 'order details report'"
                        Call CP.Db.ExecuteSQL(sql)
                        sql = "delete from ccMenuEntries where name = 'orders report'"
                        Call CP.Db.ExecuteSQL(sql)
                    End If
                    '
                    '---------------------------------------------------------------------------------
                    ' Clean up
                    '---------------------------------------------------------------------------------
                    '
                    If mustLoadCDef Then
                        Call CP.Utils.ExecuteAddon("Contensive Patch LoadContentDefinitions")
                    End If
                    '
                    '---------------------------------------------------------------------------------
                    ' Done
                    '---------------------------------------------------------------------------------
                    '
                    Call CP.Site.SetProperty("PatchRevision", patchVersion)
                End If

            Catch ex As Exception

            End Try
            Return s
        End Function
    End Class
End Namespace
