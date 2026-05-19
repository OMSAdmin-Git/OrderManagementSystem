Namespace OMS.Common
    ''' <summary>
    ''' シンプルなファイルベースのロガー。
    ''' - スレッドセーフ
    ''' - 出力先ディレクトリが無ければ自動作成
    ''' - 失敗時は Console.WriteLine にフォールバック
    ''' 既存の New(path) と Write(message) は完全互換。
    ''' </summary>
    Public Class Logger

#Region "フィールド/コンストラクタ"
        Private ReadOnly _logFilePath As String
        Private Shared ReadOnly SyncRoot As New Object()

        Public Sub New(path As String)
            If String.IsNullOrWhiteSpace(path) Then
                Throw New ArgumentException("log file path is empty.", NameOf(path))
            End If
            _logFilePath = path
        End Sub
#End Region

#Region "公開API（既存互換 + オーバーロード）"
        ''' <summary>
        ''' メッセージを INFO レベルとして出力（既存互換）。
        ''' </summary>
        Public Sub Write(message As String)
            WriteInternal("INFO", message, Nothing)
        End Sub

        ''' <summary>
        ''' メッセージを指定レベルで出力（任意利用）。
        ''' 例: Write("WARN", "設定値が空です")
        ''' </summary>
        Public Sub Write(level As String, message As String)
            WriteInternal(level, message, Nothing)
        End Sub

        ''' <summary>
        ''' 例外付きで出力（任意利用）。
        ''' </summary>
        Public Sub Write(level As String, message As String, ex As Exception)
            WriteInternal(level, message, ex)
        End Sub
#End Region

#Region "内部実装"
        Private Sub WriteInternal(level As String, message As String, ex As Exception)
            Dim ts As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim lv As String = If(String.IsNullOrWhiteSpace(level), "INFO", level.Trim().ToUpperInvariant())

            Dim line As String
            If ex Is Nothing Then
                line = $"{ts} [{lv}] {message}"
            Else
                ' 例外情報を簡潔に（必要に応じて StackTrace まで）
                line = $"{ts} [{lv}] {message} | {ex.GetType().Name}: {ex.Message}"
            End If

            ' 実務: ファイル出力（失敗時は Console へフォールバック）
            Try
                EnsureLogDirectory()
                SyncLock SyncRoot
                    System.IO.File.AppendAllText(_logFilePath, line & Environment.NewLine)
                End SyncLock
            Catch
                ' 出力に失敗した場合はコンソールへ
                Console.WriteLine(line)
            End Try
        End Sub

        ''' <summary>
        ''' ログ出力先のディレクトリを保証。
        ''' </summary>
        Private Sub EnsureLogDirectory()
            Try
                Dim dir = System.IO.Path.GetDirectoryName(_logFilePath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
                    System.IO.Directory.CreateDirectory(dir)
                End If
            Catch
                ' ディレクトリ作成失敗時は呼び元でのファイル出力 Try/Catch に任せる
            End Try
        End Sub
#End Region

    End Class
End Namespace