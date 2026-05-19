Imports System.Web
Imports System.Web.Services
Imports Microsoft.SqlServer
Imports OMS.Common

Public Class DownloadProcess
    Implements System.Web.IHttpHandler
    ''' <summary>
    ''' Download 処理 (iFrame による別ページ)
    ''' Webページにはiframe をセットする
    '''             <!-- ダウンロードを裏で実行するための隠し iframe -->
    '''<iframe id = "downloadFrame" style="display:none;"></iframe>
    ''' 
    '''  呼び出し元は 次のように url に Download File リストのファイル名とzipの時のファイル名をパラメータにセットする 
    ''' </summary>
    ''' <param name="context"></param>
    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        ' 2026/5/19
        ' 2. URLにファイル名をパラメータとして付与する（日本語などの場合は UrlEncode を推奨）
        'Dim url As String = $"DownloadProcess.ashx?file={HttpUtility.UrlEncode(fileName)}&list={HttpUtility.UrlEncode(listName)}"
        'Dim script As String = $"document.getElementById('downloadFrame').src = '{url}';"
        'ClientScript.RegisterStartupScript(Me.GetType(), "downloadScript", script, True)

        ' URLパラメータから zip ファイル名
        ' ファイルリスト ファイル名を取得
        Dim fileName = context.Request.QueryString("file")
        Dim listName = context.Request.QueryString("list")

        If (String.IsNullOrEmpty(fileName)) Then
            fileName = ""
        End If
        Dim Response = context.Response
        Dim Server = context.Server
        Dim fileList = Utils.LoadFileList(listName)
        ' 使用したリストのファイルを削除する
        IO.File.Delete(listName)
        Utils.FilesTransfer(Response, Server, fileList, fileName)
    End Sub
    ''' <summary>
    ''' property
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class