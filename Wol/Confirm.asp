<!--#include file="Connections/WolLan.asp" -->
<%
  Dim strUser, strPassword
  strUser = Request.Form("Textfield")
  strPassword = Request.Form("Textfield2")

  Dim rsUsers
  set rsUsers = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM Users WHERE userName = '" & strUser & "';"
  rsUsers.Open strSQL, objConn

  If rsUsers.EOF Then                                               ' User not found
    Session("User") = strUser
    If Request("SecondTry") = "True" then                      ' User's had two goes
      Response.Redirect "index.asp"               ' - must register
    Else                                            ' Username wrong; password wrong
      Response.Redirect "Admin/Update.asp"              ' - allow another go
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
'        Response.Redirect "menu_for_registered_users.asp"       ' successful login
        Response.Redirect "Admin/Update.asp"       ' successful login
      Else 
        rsUsers.MoveNext
      End If
    Wend
    Session("User") = strUser          ' if we get this far then...
                                       ' ...password doesn't match any of DB entries

    If Request("SecondTry") = "True" then                      ' User's had two goes
      Response.Redirect "index.asp.asp"                         ' - must reregister
    Else                                            ' Username right; password wrong
      Response.Redirect "Admin.asp" 
                                                                 '- allow another go
    End If
  End If
%>
