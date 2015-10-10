Attribute VB_Name = "todoist_Auth"
'---------------------------------------------------------------------------------------
' Project   : Excel_TODOist
' Module    : todoist_Auth
' Author    : Mauricio Souza (mauriciojxs@yahoo.com.br)
' Date      : 2015-09-21
' License   : MIT (http://www.opensource.org/licenses/mit-license.php
' Purpose   : Example of TODOist authentication to obtain a Token
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure  : login_example
' Purpose    : Shows an example of a request to TODOist using an Authenticator
'
' Parameters : []
' Return     : []
'---------------------------------------------------------------------------------------
Sub login_example()

    'set TRUE to enable debbuging in the immediate window
    WebHelpers.EnableLogging = True

    Dim Auth As New TodoistAuthenticator
    Dim CLIENT_ID As String
    Dim CLIENT_SECRET As String
    Dim REDIRECT_URL As String

    'set your ClientID, ClientSecret and RedirectURL obtained at https://developer.todoist.com/appconsole.html
    'you can set a ficticious URL, like https://com.yourappname.redirecturl, but use the same one in the App Console
    'and here in the code
    CLIENT_ID = "your ID here"
    CLIENT_SECRET = "your Secret here"
    REDIRECT_URL = "your redirect URL here"
    Call Auth.Setup(CLIENT_ID, CLIENT_SECRET, REDIRECT_URL)

    'set the scope you want for the access, as defined in https://developer.todoist.com/index.html#oauth
    Auth.Scope = "data:read_write"

    Dim Client As New WebClient

    'API Base URL provided by TODOist
    Client.BaseUrl = "https://todoist.com/API/v6/"

    'Define the authenticator to be used by the client
    Set Client.Authenticator = Auth


    Dim Request As New WebRequest

    'this is commom to all requests to TODOist
    Request.Method = WebMethod.HttpPost
    Request.Format = WebFormat.FormUrlEncoded
    Request.Resource = "sync"
    Request.AddQuerystringParam "seq_no", 0 'or the last one you received
    Request.AddQuerystringParam "seq_no_global", 0 'or the last one you received

    'this is the specific to the type of request you want
    Request.AddQuerystringParam "resource_types", "[""projects""]"

    Dim Response As WebResponse

    'execute
    Set Response = Client.Execute(Request)

    'do whatever you want with the Response

End Sub