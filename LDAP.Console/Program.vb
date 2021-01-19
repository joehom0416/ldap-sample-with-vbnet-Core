Imports System
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Net
Imports System.Net.NetworkInformation

Module Program
    Sub Main(args As String())
        Dim username As String = Environment.UserName
        Console.WriteLine("User Name:" & username)
        Console.WriteLine("=====================")
        Console.Write("Enter LDAP path (put empty system will get for you):")
        Dim ldapPath As String = Console.ReadLine()
        Console.WriteLine("=====================")
        If IsUserExistInDomain(username, ldapPath) Then
            Console.WriteLine("Result:Success.")
        Else
            Console.WriteLine("Result:Failed")
        End If

        Console.WriteLine("Verify password of " & username)
        Console.WriteLine("=====================")
        Console.Write("Password:")
        Dim pwd As String = GetInput()
        Console.WriteLine()
        If VerifyUserPassword(username, pwd, ldapPath) Then
            Console.WriteLine("Result:Success.")
        Else
            Console.WriteLine("Result:Failed")
        End If


        Console.ReadKey()
    End Sub
    ''' <summary>
    ''' Get User Input
    ''' </summary>
    ''' <returns></returns>
    Public Function GetInput() As String
        Dim pwd As String = ""
        While True
            Dim i As ConsoleKeyInfo = Console.ReadKey(True)
            If i.Key = ConsoleKey.Enter Then
                Exit While
            ElseIf i.Key = ConsoleKey.Backspace Then
                If pwd.Length > 0 Then
                    pwd = pwd.Remove(pwd.Length - 1, 1)
                    Console.Write(vbBack & " " & vbBack)
                End If
            Else
                pwd &= (i.KeyChar)
                Console.Write("*")
            End If
        End While
        Return pwd
    End Function

    Function IsUserExistInDomain(ByVal username As String, ldapPath As String) As Boolean
        Try
            Dim root As DirectoryEntry
            If (String.IsNullOrEmpty(ldapPath)) Then
                root = New DirectoryEntry()
            Else
                root = New DirectoryEntry(ldapPath)
            End If


            'Dim collections = root.Properties()
            'For Each value As PropertyValueCollection In collections
            '    Console.WriteLine(value.PropertyName)
            '    Console.WriteLine(value.Value)
            '    Console.WriteLine(value.Count)
            'Next


            Dim searcher As New DirectoryServices.DirectorySearcher(root)
            'Console.WriteLine("LDAP Path: " & searcher.SearchRoot.Path)

            searcher.PageSize = Integer.MaxValue
            searcher.SearchScope = DirectoryServices.SearchScope.Subtree
            searcher.ClientTimeout = New TimeSpan(0, 0, 10)
            ' searcher.Filter = String.Format("(&(objectCategory=Person)(anr={0}))", username)
            searcher.Filter = String.Format("(&(SAMAccountName={0})(!(userAccountControl:1.2.840.113556.1.4.803:=2)))", username)

            Dim result = searcher.FindOne()
            If result Is Nothing Then
                Return False
            Else

                Return True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace.ToString())
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Verify User Name with Directory Searcher
    ''' </summary>
    ''' <param name="username">User Name</param>
    ''' <param name="password">Password</param>
    ''' <returns></returns>
    Function VerifyUserPassword(ByVal username As String, ByVal password As String, ldapPath As String) As Boolean

        Dim root As DirectoryEntry
        If (String.IsNullOrEmpty(ldapPath)) Then
            root = New DirectoryEntry()
        Else
            root = New DirectoryEntry(ldapPath)
        End If
        Dim searcher As New DirectoryServices.DirectorySearcher(root)
        searcher.SearchRoot.Username = username
        searcher.SearchRoot.Password = password
        searcher.PageSize = Integer.MaxValue

        searcher.SearchScope = DirectoryServices.SearchScope.Subtree
        searcher.ClientTimeout = New TimeSpan(0, 0, 10)
        ' searcher.Filter = String.Format("(&(objectCategory=Person)(anr={0}))", username)
        searcher.Filter = String.Format("(&(SAMAccountName={0})(!(userAccountControl:1.2.840.113556.1.4.803:=2)))", username)
        Try
            Dim result = searcher.FindOne()
            If result Is Nothing Then
                Return False
            Else

                Return True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
            Console.WriteLine(ex.StackTrace.ToString())
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Get Domain 
    ''' </summary>
    ''' <returns></returns>
    Function GetFQDN() As String
        Dim domainName As String = IPGlobalProperties.GetIPGlobalProperties().DomainName
        Dim hostName As String = Dns.GetHostName()
        Dim d As Domain = Domain.GetDomain(New DirectoryContext(DirectoryContextType.Domain, domainName))
        Return d.InfrastructureRoleOwner.Name
    End Function
End Module
