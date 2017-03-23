Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Configuration
Imports System.Web.Security

Public Class AccessMembershipProvider
    Inherits System.Web.Security.MembershipProvider

    '---for database access use---
    Public connStr As String
    Private comm As New OleDb.OleDbCommand
    Private _requiresQuestionAndAnswer As Boolean
    Private _minRequiredPasswordLength As Integer
    Private _enablePasswordReset As Boolean

    Private Function ConfigDBsourcePath() As String
        ' ===== try to retrieve DB source path from web.config ========================================
        If My.Computer.FileSystem.FileExists(ConfigurationManager.AppSettings("MembershipDBSourcePath")) Then
            Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConfigurationManager.AppSettings("MembershipDBSourcePath") & ";Persist Security Info=False"
        Else
            Return connStr
        End If
    End Function

    Public Overrides Sub Initialize(ByVal name As String, ByVal config As System.Collections.Specialized.NameValueCollection)
        '===retrives the attribute values set in
        'web.config and assign to local variables===
        If config("requiresQuestionAndAnswer") = "true" Then _
        _requiresQuestionAndAnswer = True

        If config("EnablePasswordReset") = "true" Then _
        _enablePasswordReset = True

        connStr = config("connectionString")

        MyBase.Initialize(name, config)
    End Sub

    Public Overrides ReadOnly Property RequiresQuestionAndAnswer() As Boolean
        Get
            If _requiresQuestionAndAnswer = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Overrides Function ValidateUser(ByVal username As String, ByVal password As String) As Boolean
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Select * From Membership WHERE " & "username=@username AND password=@password"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            comm.Parameters.AddWithValue("@username", username)
            comm.Parameters.AddWithValue("@password", password)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            If reader.HasRows Then
                ' set user.lastlogindate
                Return True
            Else
                Return False
            End If
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
            Return False
        End Try
    End Function

    Public Overrides Property ApplicationName() As String
        Get
            Return ""
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Overrides Function ChangePassword(ByVal username As String, ByVal oldPassword As String, ByVal newPassword As String) As Boolean
        Return False
    End Function

    Public Overrides Function ChangePasswordQuestionAndAnswer(ByVal username As String, ByVal password As String, ByVal newPasswordQuestion As String, ByVal newPasswordAnswer As String) As Boolean
        Return False
    End Function

    Public Overrides Function CreateUser(ByVal username As String, ByVal password As String, ByVal email As String, ByVal passwordQuestion As String, ByVal passwordAnswer As String, ByVal isApproved As Boolean, ByVal providerUserKey As Object, ByRef status As System.Web.Security.MembershipCreateStatus) As System.Web.Security.MembershipUser
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Insert into Membership Values " & "(@username , @password, @email, @qst, @answ)"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            comm.Parameters.AddWithValue("@username", username)
            comm.Parameters.AddWithValue("@password", password)
            comm.Parameters.AddWithValue("@email", email)
            comm.Parameters.AddWithValue("@qst", passwordQuestion)
            comm.Parameters.AddWithValue("@answ", passwordAnswer)
            comm.ExecuteNonQuery()
            conn.Close()
            status = MembershipCreateStatus.Success
            Return New MembershipUser("AccessMembershipProvider", username, username, email, passwordQuestion, String.Empty, True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now)
        Catch ex As Exception
            Console.Write(ex.ToString)
            status = MembershipCreateStatus.ProviderError
            Return New MembershipUser("AccessMembershipProvider", String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False, True, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now)
        End Try
    End Function

    Public Overrides Function DeleteUser(ByVal username As String, ByVal deleteAllRelatedData As Boolean) As Boolean
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Delete from Membership where username = @username"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            comm.Parameters.AddWithValue("@username", username)
            comm.ExecuteNonQuery()

            If deleteAllRelatedData Then
                sql = "Delete from Membership_Roles where username = @username"
                Dim comm2 As New OleDb.OleDbCommand(sql, conn)
                comm2.Parameters.AddWithValue("@username", username)
                comm2.ExecuteNonQuery()

                sql = "Delete from Membership_Groups where username = @username"
                Dim comm3 As New OleDb.OleDbCommand(sql, conn)
                comm3.Parameters.AddWithValue("@username", username)
                comm3.ExecuteNonQuery()
            End If

            conn.Close()
            Return True
        Catch ex As Exception
            Console.Write(ex.ToString)
            Return False
        End Try
    End Function

    Public Overrides Sub UpdateUser(ByVal user As System.Web.Security.MembershipUser)
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Update Membership set email = @email, nome = @nome, cognome = @cognome where username ='" & user.UserName & "'"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            comm.Parameters.AddWithValue("@email", user.Email)
            comm.Parameters.AddWithValue("@nome", user.Comment.Split(" ")(0))
            comm.Parameters.AddWithValue("@cognome", user.Comment.Split(" ")(1))

            'user.IsApproved , user.LastLoginDate, user.LastActivityDate

            comm.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
    End Sub

    Public Overrides ReadOnly Property EnablePasswordReset() As Boolean
        Get
            Return _enablePasswordReset
        End Get
    End Property

    Public Overrides ReadOnly Property EnablePasswordRetrieval() As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides Function FindUsersByEmail(ByVal emailToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim mUCollection As New MembershipUserCollection
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Select * from Membership where email ='" & emailToMatch & "'"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            Dim i As Integer = 1
            For Each m In reader
                If i > pageIndex * pageSize And i <= (pageIndex + 1) * pageSize Then
                    mUCollection.Add(New MembershipUser("AccessMembershipProvider", m("username") & "", m("nome") & " " & m("cognome"), m("email") & "", m("passwordQuestion") & "", String.Empty, True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now))
                End If
                i = i + 1
            Next
            totalRecords = i - 1
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return mUCollection
    End Function

    Public Overrides Function FindUsersByName(ByVal usernameToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim mUCollection As New MembershipUserCollection
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Select * from Membership where username ='" & usernameToMatch & "'"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            Dim i As Integer = 1
            For Each m In reader
                If i > pageIndex * pageSize And i <= (pageIndex + 1) * pageSize Then
                    mUCollection.Add(New MembershipUser("AccessMembershipProvider", m("username") & "", m("nome") & " " & m("cognome"), m("email") & "", m("passwordQuestion") & "", String.Empty, True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now))
                End If
                i = i + 1
            Next
            totalRecords = i - 1
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return mUCollection
    End Function

    Public Overrides Function GetAllUsers(ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim mUCollection As New MembershipUserCollection
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Select * from Membership"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            Dim i As Integer = 1
            For Each m In reader
                If i > pageIndex * pageSize And i <= (pageIndex + 1) * pageSize Then
                    mUCollection.Add(New MembershipUser("AccessMembershipProvider", m("username") & "", m("nome") & " " & m("cognome"), m("email") & "", m("passwordQuestion") & "", String.Empty, True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now))
                End If
                i = i + 1
            Next
            totalRecords = i - 1
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return mUCollection
    End Function

    Public Overrides Function GetNumberOfUsersOnline() As Integer
        Return 0
    End Function

    Public Overrides Function GetPassword(ByVal username As String, ByVal answer As String) As String
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Dim _password As String = ""
        Try
            conn.Open()
            Dim sql As String = "Select password from Membership where username ='" & username & "' and passwordAnswer = '" & answer & "'"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            For Each m In reader
                _password = m("password")
            Next
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return _password
    End Function

    Public Overloads Overrides Function GetUser(ByVal providerUserKey As Object, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser
        Return New MembershipUser("AccessMembershipProvider", "fakeUser", "", New Object, "", "", True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now)
    End Function

    Public Overloads Overrides Function GetUser(ByVal username As String, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser
        Dim _username As String = ""
        Dim _comment As String = ""
        Dim _email As String = ""
        Dim _passwordQuestion As String = ""
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        'Try
        conn.Open()
        Dim sql As String = "Select * from Membership where username ='" & username & "'"
        Dim comm As New OleDb.OleDbCommand(sql, conn)
        Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
        For Each m In reader
            _username = m("username")
            _comment = m("nome") & " " & m("cognome") & ""
            _email = m("email") & ""
            _passwordQuestion = m("passwordQuestion") & ""
        Next
        conn.Close()
        Return New MembershipUser("AccessMembershipProvider", _username, New Object, _email, _passwordQuestion, _comment, True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now)
        'Catch ex As Exception
        '    Console.Write(ex.ToString)
        '    Return New MembershipUser("AccessMembershipProvider", "unnamedUser", "", "", "", "", True, False, Date.Now, Date.Now, Date.Now, Date.Now, Date.Now)
        'End Try
    End Function

    Public Overrides Function GetUserNameByEmail(ByVal email As String) As String
        Dim _username As String = ""
        Dim conn As New OleDb.OleDbConnection(ConfigDBsourcePath)
        Try
            conn.Open()
            Dim sql As String = "Select top 1 username from Membership where email ='" & email & "'"
            Dim comm As New OleDb.OleDbCommand(sql, conn)
            Dim reader As OleDb.OleDbDataReader = comm.ExecuteReader
            For Each m In reader
                _username = m("username")
            Next
            conn.Close()
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return _username
    End Function

    Public Overrides ReadOnly Property MaxInvalidPasswordAttempts() As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property MinRequiredNonAlphanumericCharacters() As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property MinRequiredPasswordLength() As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property PasswordAttemptWindow() As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property PasswordFormat() As System.Web.Security.MembershipPasswordFormat
        Get
            Return MembershipPasswordFormat.Clear
        End Get
    End Property

    Public Overrides ReadOnly Property PasswordStrengthRegularExpression() As String
        Get
            Return ""
        End Get
    End Property

    Public Overrides ReadOnly Property RequiresUniqueEmail() As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides Function ResetPassword(ByVal username As String, ByVal answer As String) As String
        Return ""
    End Function

    Public Overrides Function UnlockUser(ByVal userName As String) As Boolean
        Return False
    End Function

    Public Function getMembershipRoles() As String()
        Dim a As String()
        ReDim a(2)
        a(0) = ""
        Return a
    End Function

End Class
