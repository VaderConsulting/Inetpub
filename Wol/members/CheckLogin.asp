<!--#include file="Database.asp"-->
<%
  Dim strEMail, strPassword
  strUser = Request("Username")
  strPassword = Request("Password")

  Dim rsUsers
  set rsUsers = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM Users WHERE Username = '" & strUser & "';"
  rsUsers.Open strSQL, objConn

  If rsUsers.EOF Then                                               ' User not found
    Session("User") = Request("User")
    If Request("SecondTry") = "True" then                      ' User's had two goes
      Response.Redirect "Register.asp"               ' - must register
    Else                                            ' Username wrong; password wrong
      Response.Redirect "Default.asp"              ' - allow another go
    End If
  Else                                     'One or more users found - check password
    While Not rsUsers.EOF
      If UCase(rsUsers("Password")) = UCase(strPassword) Then     ' password matched
        Dim strName, strValue
        For Each strField in rsUsers.Fields
          strName = strField.Name                       ' populate session variables
          strValue = strField.value
          Session(strName) = strValue
        Next
        Session("blnValidUser") = True                          ' update the database
		Session.Timeout = 300
        Response.Redirect "menu_for_registered_users.asp"       ' successful login
      Else 
        rsUsers.MoveNext
      End If
    Wend
    Session("User") = Request("User")          ' if we get this far then...
                                       ' ...password doesn't match any of DB entries

    If Request("SecondTry") = "True" then                      ' User's had two goes
      Response.Redirect "Register.asp"                         ' - must reregister
    Else                                            ' Username right; password wrong
      Response.Redirect "Default.asp?SecondTry=True&WrongPW=True" 
                                                                 '- allow another go
    End If
  End If
%>
