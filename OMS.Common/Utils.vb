Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Data.Odbc
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.ExtendedProperties
Imports DocumentFormat.OpenXml.Math
Imports DocumentFormat.OpenXml.Vml.Office
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Ionic.Zip
Imports Microsoft.SqlServer
Imports Microsoft.VisualBasic.FileIO
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Common

    '===============================
    ' Utils
    '===============================
    Public Module Utils

#Region "DB 接続 / 検索条件"

        ''' <summary>
        ''' 接続文字列取得
        ''' </summary>
        Public Function GetConnectionString() As String
            Dim cs = ConfigurationManager.ConnectionStrings("OMSConnection")
            If cs Is Nothing OrElse String.IsNullOrWhiteSpace(cs.ConnectionString) Then
                Throw New InvalidOperationException("接続文字列 'OMSConnection' が見つかりません。Web.config を確認してください。")
            End If
            Return cs.ConnectionString
        End Function

        ''' <summary>検索条件</summary>
        Public Enum LikeMode
            Contains   ' 部分一致（%term%）
            StartsWith ' 前方一致（term%）
            EndsWith   ' 後方一致（%term）
            Exact      ' 完全一致（term）
        End Enum

        ''' <summary>
        ''' Oracle LIKE 用にパターンを作成（%と_をエスケープし、指定モードでワイルドカードを付与）
        ''' NULL/空白は Nothing を返して WHERE で無視できるようにする。
        ''' </summary>
        Public Function BuildLikePattern(input As String, mode As LikeMode) As String
            If String.IsNullOrWhiteSpace(input) Then
                Return Nothing
            End If
            Dim s As String = input.Trim()

            ' エスケープ（\% や \_）
            s = s.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")

            Select Case mode
                Case LikeMode.Contains
                    Return "%" & s & "%"
                Case LikeMode.StartsWith
                    Return s & "%"
                Case LikeMode.EndsWith
                    Return "%" & s
                Case LikeMode.Exact
                    Return s
                Case Else
                    Return "%" & s & "%"
            End Select
        End Function

#End Region

#Region "パス / ファイル・フォルダ操作"

        ''' <summary>
        ''' パス解決（仮想/相対/絶対）→ 物理パス
        ''' </summary>
        Public Function ResolvePath(server As HttpServerUtility, raw As String) As String
            If String.IsNullOrWhiteSpace(raw) Then Return raw
            If Path.IsPathRooted(raw) Then Return NormalizeSep(raw)
            If raw = "~" Then Return server.MapPath("~/")

            If raw.StartsWith("~") Then
                Dim trimmed = raw.TrimStart("~"c).TrimStart("/"c, "\"c)
                Return server.MapPath("~/" & trimmed)
            End If

            If raw.StartsWith("/") OrElse raw.StartsWith("\") Then
                Return server.MapPath("~/" & raw.TrimStart("/"c, "\"c))
            End If

            Return server.MapPath("~/" & raw.TrimStart("/"c, "\"c))
        End Function

        ''' <summary>ディレクトリが無い場合に新規作成</summary>
        Public Sub EnsureDirectory(path As String)
            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
        End Sub

        ''' <summary>
        ''' フォルダ名として安全な文字列に変換（禁止文字除去・長さ制限・空は "unknown"）
        ''' </summary>
        Public Function SafeFolderName(input As String, Optional maxLength As Integer = 64) As String
            If String.IsNullOrWhiteSpace(input) Then Return "unknown"

            Dim s As String = input.Trim()

            ' AD形式 "DOMAIN\user" → ユーザー名抽出（任意）
            Dim lastSlash = s.LastIndexOf("\"c)
            If lastSlash >= 0 AndAlso lastSlash < s.Length - 1 Then
                s = s.Substring(lastSlash + 1)
            End If

            ' 禁止文字の除去（\ / : * ? " < > |、制御文字、末尾の '.'）
            s = Regex.Replace(s, "[\\/:*?""<>|\u0000-\u001F]", "")
            s = s.Trim().Trim("."c)

            If String.IsNullOrEmpty(s) Then s = "unknown"
            If s.Length > maxLength Then s = s.Substring(0, maxLength)
            Return s
        End Function

        ''' <summary>パスセパレータ正規化</summary>
        Private Function NormalizeSep(path As String) As String
            Return path.Replace("/", "\")
        End Function
        ''' <summary>
        ''' Work base folder 取得
        ''' </summary>
        ''' <returns></returns>
        Public Function GetWorkFolderRoot() As String
            Dim rawWork As String = ConfigurationManager.AppSettings("WorkFolderRoot")
            If String.IsNullOrWhiteSpace(rawWork) Then
                Throw New ConfigurationErrorsException("appSettings['WorkFolderRoot'] が未定義です。Web.config を確認してください。")
            End If
            Return rawWork
        End Function
        ''' <summary>
        ''' Work path 取得 (ない場合は作成する)
        ''' </summary>
        ''' <returns></returns>
        Public Function GetWorkPath() As String

            Dim rawWork As String = GetWorkFolderRoot()
            Dim workFolder = rawWork & "FileTempFolder"
            If (Not IO.File.Exists(workFolder)) Then
                EnsureDirectory(workFolder)
            End If
            Return workFolder
        End Function
        ''' <summary>
        ''' Work base folder 取得
        ''' </summary>
        ''' <returns></returns>
        Public Function GetDoneFolderRoot() As String
            Dim rawWork As String = ConfigurationManager.AppSettings("CompletedFolderRoot")
            If String.IsNullOrWhiteSpace(rawWork) Then
                Throw New ConfigurationErrorsException("appSettings['CompletedFolderRoot'] が未定義です。Web.config を確認してください。")
            End If
            Return rawWork
        End Function

#End Region

#Region "値/文字列 正規化・整形"

        ''' <summary>日付を yyyy/MM/dd 形式に変換</summary>
        Public Function FormatDate(dt As DateTime) As String
            Return dt.ToString("yyyy/MM/dd")
        End Function

        ''' <summary>文字列整形：トリム＋大文字化</summary>
        Public Function NormalizeString(value As String) As String
            If String.IsNullOrEmpty(value) Then Return String.Empty
            Return value.Trim().ToUpper()
        End Function

        ''' <summary>入力文字列を Y/N に正規化（null/空/不正値は "N"）</summary>
        Public Function NormalizeYN(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return "N"
            Dim s = value.Trim().ToUpperInvariant()
            If s = "Y" OrElse s = "N" Then Return s
            Return "N"
        End Function

        ''' <summary>文字列を最大長までトリムして返す（NULLは DBNull.Value）</summary>
        Public Function SafeVarchar(s As String, maxLen As Integer) As Object
            Dim dummy As Boolean
            Return SafeVarchar(s, maxLen, dummy)
        End Function
        ''' <summary>文字列を最大長までトリムして返す（NULLは DBNull.Value）</summary>
        Public Function SafeVarchar(s As String, maxLen As Integer, ByRef isTruncated As Boolean) As Object
            'If s Is Nothing Then Return DBNull.Value
            'Dim t = s.Trim()
            'If t.Length <= maxLen Then Return t
            'Return t.Substring(0, maxLen)
            isTruncated = False

            'UTF-8対応
            If s Is Nothing Then
                Return DBNull.Value
            End If

            Dim t As String = s.Trim()

            ' UTF-8 のバイト数を取得
            Dim utf8 As Encoding = Encoding.UTF8

            ' そのままで収まる場合
            If utf8.GetByteCount(t) <= maxLen Then
                Return t
            End If

            'これより先の処理はトリム処理
            isTruncated = True

            ' UTF-8 のバイト長を超えない位置まで切り詰め
            Dim result As New StringBuilder()
            Dim currentBytes As Integer = 0

            For Each ch As Char In t

                Dim charBytes As Integer = utf8.GetByteCount(ch.ToString())

                ' 追加すると上限超過する場合は終了
                If currentBytes + charBytes > maxLen Then
                    Exit For
                End If

                result.Append(ch)
                currentBytes += charBytes

            Next

            Return result.ToString()
        End Function

        ''' <summary>文字列を最大長までトリムして返す（NULLは DBNull.Value）</summary>
        Public Function SafeVarcharLength(s As String, maxLen As Integer) As Object
            Dim dummy As Boolean
            Return SafeVarcharLength(s, maxLen, dummy)
        End Function


        ''' <summary>文字列を最大長までトリムして返す（NULLは DBNull.Value）</summary>
        Public Function SafeVarcharLength(s As String, maxLen As Integer, ByRef isTruncated As Boolean) As Object
            isTruncated = False
            If s Is Nothing Then Return DBNull.Value
            Dim t = s.Trim()
            If t.Length <= maxLen Then Return t
            isTruncated = True
            Return t.Substring(0, maxLen)
        End Function



        ''' <summary>空/空白なら Nothing、そうでなければ Trim 済み文字列</summary>
        Public Function NullIfWhite(ByVal s As String) As String
            If s Is Nothing Then Return Nothing
            s = s.Trim()
            Return If(s.Length = 0, Nothing, s)
        End Function

#End Region

#Region "区切り/囲み/改行/文字コード マッピング"

        ''' <summary>区切り文字名を区切り文字に変換</summary>
        Public Function MapDelimiter(code As String) As String
            If String.IsNullOrWhiteSpace(code) Then Return ","
            Select Case code.Trim().ToUpperInvariant()
                Case Constants.DELIMITER_COMMA, "COMMA" : Return ","
                Case Constants.DELIMITER_TAB, "TAB" : Return vbTab
                Case Constants.DELIMITER_SEMICOLON, "SEMICOLON" : Return ";"
                Case Constants.DELIMITER_PIPE, "PIPE" : Return "|"
                Case Constants.DELIMITER_SPACE, "SPACE" : Return " "
                Case Constants.DELIMITER_COLON, "COLON" : Return ":"
                Case Else : Return ","
            End Select
        End Function

        ''' <summary>囲み文字名を囲み文字に変換</summary>
        Public Function MapQuote(code As String) As String
            If String.IsNullOrWhiteSpace(code) Then Return """" ' 既定ダブルクォート
            Select Case code.Trim().ToUpperInvariant()
                Case Constants.QUOTE_D, "D_QUOTE" : Return """" ' ダブルクォート
                Case Constants.QUOTE_S, "S_QUOTE" : Return "'"   ' シングルクォート
                Case Constants.QUOTE_NONE, "NONE" : Return ""    ' 囲みなし
                Case Else : Return """"                          ' 既定はダブルクォート
            End Select
        End Function

        ''' <summary>
        ''' 改行コード名（"CRLF"/"LF"/"CR"）または実体（vbCrLf/vbLf/vbCr）から改行文字列を返す。
        ''' </summary>
        Public Function MapNewline(code As String) As String
            If String.IsNullOrEmpty(code) Then Return vbCrLf
            Dim key = code.Trim().ToUpperInvariant()

            Select Case key
                Case Constants.NEWLINE_CRLF, "CRLF" : Return vbCrLf
                Case Constants.NEWLINE_LF, "LF" : Return vbLf
                Case Constants.NEWLINE_CR, "CR" : Return vbCr
            End Select

            If code = vbCrLf OrElse code = vbLf OrElse code = vbCr Then
                Return code
            End If

            Return vbCrLf
        End Function

        ''' <summary>文字コード名を Encoding に変換（簡易マッピング＋同義語許容）</summary>
        Public Function MapEncoding(code As String) As Encoding
            If String.IsNullOrWhiteSpace(code) Then
                Return New UTF8Encoding(False) ' 既定はUTF-8 (BOMなし)
            End If

            Dim key = code.Trim().ToUpperInvariant()
            Select Case key
                Case Constants.ENCODING_SJIS, "SHIFT_JIS", "SJIS"
                    Return Encoding.GetEncoding("Shift_JIS")
                Case Constants.ENCODING_UTF8, "UTF-8", "UTF8"
                    Return New UTF8Encoding(False) ' BOMなし
                Case Constants.ENCODING_UTF8_BOM, "UTF-8-BOM", "UTF8-BOM", "UTF8BOM"
                    Return New UTF8Encoding(True)  ' BOMあり
                Case Else
                    Try
                        Return Encoding.GetEncoding(code)
                    Catch
                        Return New UTF8Encoding(False)
                    End Try
            End Select
        End Function

#End Region

#Region "CSV 整形 / 出力"

        '===============================
        ' CSV 整形ヘルパ
        '===============================
        ''' <summary>
        ''' CSV1セル分の文字列を整形。
        ''' - 囲み文字あり：囲み文字を二重化し、必要時に全体を囲む
        ''' - 囲みなし：原則そのまま（区切り/改行/前後スペース等は囲み推奨）
        ''' </summary>
        Public Function CsvFormat(value As String, delimiter As String, enclosure As String) As String
            If value Is Nothing Then value = ""

            If enclosure = "" Then
                Return value
            End If

            Dim needsEnclose As Boolean =
                value.Contains(delimiter) OrElse
                value.Contains(vbCr) OrElse value.Contains(vbLf) OrElse
                value.StartsWith(" ") OrElse value.EndsWith(" ") OrElse
                value.Contains(enclosure)

            Dim escaped As String = value.Replace(enclosure, enclosure & enclosure)
            If needsEnclose Then
                Return enclosure & escaped & enclosure
            Else
                Return escaped
            End If
        End Function

        ''' <summary>
        ''' ASCII フォールバック用ファイル名生成（ASCII以外/空白は '_'）
        ''' </summary>
        Private Function ToAsciiFallback(fileName As String) As String
            If String.IsNullOrEmpty(fileName) Then Return "export.csv"
            Dim sb As New StringBuilder(fileName.Length)
            For Each ch As Char In fileName
                If AscW(ch) < 128 AndAlso ch <> " "c Then
                    sb.Append(ch)
                Else
                    sb.Append("_"c)
                End If
            Next
            Return sb.ToString()
        End Function

        '============================================
        ' CSV 出力（HTTP レスポンスへストリーミング）
        '============================================
        ''' <summary>
        ''' SELECT結果をCSVとしてHTTPレスポンスにストリーミング出力（定数コード版）。
        ''' 呼び出し側は Constants のコード（例: DELIMITER_COMMA, QUOTE_D, NEWLINE_CRLF, ENCODING_UTF8_BOM）を渡す。
        ''' </summary>
        Public Sub ExportQueryToCsvResponse(
            ByVal sql As String,
            ByVal delimiterCode As String,
            ByVal enclosureCode As String,
            ByVal headerYN As String,
            ByVal lineEndingCode As String,
            ByVal charsetCode As String,
            ByVal processDate As DateTime,
            Optional ByVal fileBaseName As String = "EXPORT",
            Optional ByVal connectionStringName As String = "OMSConnection",
            Optional ByVal parameters As IEnumerable(Of OdbcParameter) = Nothing,
            Optional ByVal endResponse As Boolean = True,
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        )
            ' コード → 実値へ解決
            Dim delimiter As String = MapDelimiter(delimiterCode)
            Dim enclosure As String = MapQuote(enclosureCode)
            Dim newline As String = MapNewline(lineEndingCode)
            Dim enc As Encoding = MapEncoding(charsetCode)

            ' コアへ委譲
            ExportQueryToCsvResponseCore(
                sql:=sql,
                delimiter:=delimiter,
                enclosure:=enclosure,
                includeHeader:=String.Equals(headerYN, "Y", StringComparison.OrdinalIgnoreCase),
                newline:=newline,
                enc:=enc,
                processDate:=processDate,
                fileBaseName:=fileBaseName,
                connectionStringName:=connectionStringName,
                parameters:=parameters,
                endResponse:=endResponse
            )
        End Sub

        '--- 共通コア：実値で処理
        Private Sub ExportQueryToCsvResponseCore(
            ByVal sql As String,
            ByVal delimiter As String,
            ByVal enclosure As String,
            ByVal includeHeader As Boolean,
            ByVal newline As String,
            ByVal enc As Encoding,
            ByVal processDate As DateTime,
            ByVal fileBaseName As String,
            ByVal connectionStringName As String,
            ByVal parameters As IEnumerable(Of OdbcParameter),
            ByVal endResponse As Boolean,
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        )
            ' --- 検証（SELECTのみ許可）
            If String.IsNullOrWhiteSpace(sql) OrElse Not sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) Then
                Throw New ArgumentException("SELECT文のみ許可されます。", NameOf(sql))
            End If

            ' --- HTTPコンテキスト事前確認
            Dim ctx = HttpContext.Current
            If ctx Is Nothing Then Throw New InvalidOperationException("HTTPコンテキストが取得できません。")
            Dim resp = ctx.Response

            Dim fileName As String = $"{fileBaseName}_{processDate:yyyyMMdd_HHmmss}.csv"

            ' --- DB 接続・実行（先に例外を出させる）
            Dim connStr As String = ConfigurationManager.ConnectionStrings(connectionStringName).ConnectionString
            Using conn As New OdbcConnection(connStr)
                conn.Open()
                Using cmd As New OdbcCommand(sql, conn)
                    If parameters IsNot Nothing Then
                        For Each p In parameters
                            cmd.Parameters.Add(p)
                        Next
                    End If

                    Using rdr As OdbcDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)

                        ' --- HTTP レスポンス開始
                        resp.Clear()
                        resp.Buffer = False
                        resp.ContentType = "text/csv"

                        ' Content-Disposition（日本語ファイル名配慮）
                        Dim asciiFallback As String = ToAsciiFallback(fileName)
                        Dim dispo As New StringBuilder()
                        dispo.Append($"attachment; filename=""{asciiFallback}""")
                        Try
                            dispo.Append($"; filename*=UTF-8''{Uri.EscapeDataString(fileName)}")
                        Catch
                            ' 環境により例外になる場合があるため無視
                        End Try
                        resp.AddHeader("Content-Disposition", dispo.ToString())

                        ' エンコーディング
                        resp.ContentEncoding = enc
                        resp.Charset = enc.WebName

                        ' UTF-8（BOM付き）の場合はPreambleを書き込み
                        Dim preamble As Byte() = enc.GetPreamble()
                        If preamble IsNot Nothing AndAlso preamble.Length > 0 Then
                            resp.OutputStream.Write(preamble, 0, preamble.Length)
                        End If

                        ' --- ヘッダー行
                        If includeHeader Then
                            Dim headerCols As New List(Of String)(rdr.FieldCount)
                            For i As Integer = 0 To rdr.FieldCount - 1
                                headerCols.Add(CsvFormat(rdr.GetName(i), delimiter, enclosure))
                            Next
                            resp.Write(String.Join(delimiter, headerCols))
                            resp.Write(newline)
                        End If

                        ' --- データ本体（逐次）
                        Dim rowCount As Integer = 0
                        While rdr.Read()
                            Dim cols As New List(Of String)(rdr.FieldCount)
                            For i As Integer = 0 To rdr.FieldCount - 1
                                If rdr.IsDBNull(i) Then
                                    cols.Add("")
                                Else
                                    Dim t = rdr.GetFieldType(i)
                                    Dim v As Object = rdr.GetValue(i)
                                    Dim s As String
                                    '' フォーマット指定追加
                                    'Dim n = rdr.GetName(i)
                                    'Dim f As String = Nothing
                                    '' フィールド名 同じもののフォーマット情報を取得する
                                    'If (formatEx IsNot Nothing) Then
                                    '    f = formatEx.Find(Function(x) x.name.ToUpper() = n.ToUpper()).format
                                    'End If

                                    If t Is GetType(DateTime) Then
                                        s = CType(v, DateTime).ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture)
                                    ElseIf GetType(IFormattable).IsAssignableFrom(t) Then
                                        s = DirectCast(v, IFormattable).ToString("0.##########", CultureInfo.InvariantCulture)
                                    Else
                                        s = Convert.ToString(v, CultureInfo.InvariantCulture)
                                    End If
                                    '' フォーマット指定追加
                                    'If t Is GetType(DateTime) Then
                                    '    f = If(f Is Nothing, "yyyy/MM/dd HH:mm:ss", f)
                                    '    s = CType(v, DateTime).ToString(f, CultureInfo.InvariantCulture)
                                    'ElseIf GetType(IFormattable).IsAssignableFrom(t) Then
                                    '    f = If(f Is Nothing, "0.##########", f)
                                    '    s = DirectCast(v, IFormattable).ToString(f, CultureInfo.InvariantCulture)
                                    'Else
                                    '    s = Convert.ToString(v, CultureInfo.InvariantCulture)
                                    'End If

                                    cols.Add(CsvFormat(s, delimiter, enclosure))
                                End If
                            Next

                            resp.Write(String.Join(delimiter, cols))
                            resp.Write(newline)

                            rowCount += 1
                            If (rowCount Mod 200) = 0 Then
                                resp.Flush()
                            End If
                        End While

                        resp.Flush()
                        If endResponse AndAlso ctx.ApplicationInstance IsNot Nothing Then
                            ctx.ApplicationInstance.CompleteRequest()
                        End If
                        resp.End()

                    End Using
                End Using
            End Using
        End Sub

        '============================================
        ' CSV 出力（HTTP レスポンスへストリーミング）
        '============================================
        ''' <summary>
        ''' SELECT結果をCSVとしてHTTPレスポンスにストリーミング出力（定数コード版）。
        ''' 呼び出し側は Constants のコード（例: DELIMITER_COMMA, QUOTE_D, NEWLINE_CRLF, ENCODING_UTF8_BOM）を渡す。
        ''' </summary>
        Public Sub ExportQueryToCsvResponse2(
            ByVal sql As String,
            ByVal delimiterCode As String,
            ByVal enclosureCode As String,
            ByVal headerYN As String,
            ByVal lineEndingCode As String,
            ByVal charsetCode As String,
            ByVal processDate As DateTime,
            Optional ByVal fileBaseName As String = "EXPORT",
            Optional ByVal connectionStringName As String = "OMSConnection",
            Optional ByVal parameters As IEnumerable(Of OracleParameter) = Nothing,
            Optional ByVal endResponse As Boolean = True,
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        )
            ' コード → 実値へ解決
            Dim delimiter As String = MapDelimiter(delimiterCode)
            Dim enclosure As String = MapQuote(enclosureCode)
            Dim newline As String = MapNewline(lineEndingCode)
            Dim enc As Encoding = MapEncoding(charsetCode)

            ' コアへ委譲
            ExportQueryToCsvResponseCore2(
                sql:=sql,
                delimiter:=delimiter,
                enclosure:=enclosure,
                includeHeader:=String.Equals(headerYN, "Y", StringComparison.OrdinalIgnoreCase),
                newline:=newline,
                enc:=enc,
                processDate:=processDate,
                fileBaseName:=fileBaseName,
                connectionStringName:=connectionStringName,
                parameters:=parameters,
                endResponse:=endResponse,
                formatEx:=formatEx
            )
        End Sub

        '--- 共通コア：実値で処理
        Private Sub ExportQueryToCsvResponseCore2(
            ByVal sql As String,
            ByVal delimiter As String,
            ByVal enclosure As String,
            ByVal includeHeader As Boolean,
            ByVal newline As String,
            ByVal enc As Encoding,
            ByVal processDate As DateTime,
            ByVal fileBaseName As String,
            ByVal connectionStringName As String,
            ByVal parameters As IEnumerable(Of OracleParameter),
            ByVal endResponse As Boolean,
            Optional ByVal formatEx As List(Of (name As String, format As String)) = Nothing
        )
            ' --- 検証（SELECTのみ許可）
            If String.IsNullOrWhiteSpace(sql) OrElse Not sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) Then
                Throw New ArgumentException("SELECT文のみ許可されます。", NameOf(sql))
            End If

            ' --- HTTPコンテキスト事前確認
            Dim ctx = HttpContext.Current
            If ctx Is Nothing Then Throw New InvalidOperationException("HTTPコンテキストが取得できません。")
            Dim resp = ctx.Response
            resp.Clear()
            resp.Buffer = False

            Dim fileName As String = $"{fileBaseName}_{processDate:yyyyMMdd_HHmmss}.csv"

            Try
                ' --- DB 接続・実行（先に例外を出させる）
                Dim connStr As String = ConfigurationManager.ConnectionStrings(connectionStringName).ConnectionString
                Using conn As New OracleConnection(connStr)
                    conn.Open()
                    Using cmd As New OracleCommand(sql, conn)
                        If parameters IsNot Nothing Then
                            For Each p In parameters
                                cmd.Parameters.Add(p)
                            Next
                        End If

                        Using rdr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)

                            ' --- HTTP レスポンス開始
                            resp.Clear()
                            resp.Buffer = False
                            resp.ContentType = "text/csv"

                            ' Content-Disposition（日本語ファイル名配慮）
                            Dim asciiFallback As String = ToAsciiFallback(fileName)
                            Dim dispo As New StringBuilder()
                            dispo.Append($"attachment; filename=""{asciiFallback}""")
                            Try
                                dispo.Append($"; filename*=UTF-8''{Uri.EscapeDataString(fileName)}")
                            Catch
                                ' 環境により例外になる場合があるため無視
                            End Try
                            resp.AddHeader("Content-Disposition", dispo.ToString())

                            ' エンコーディング
                            resp.ContentEncoding = enc
                            resp.Charset = enc.WebName

                            ' UTF-8（BOM付き）の場合はPreambleを書き込み
                            Dim preamble As Byte() = enc.GetPreamble()
                            If preamble IsNot Nothing AndAlso preamble.Length > 0 Then
                                resp.OutputStream.Write(preamble, 0, preamble.Length)
                            End If

                            ' --- ヘッダー行
                            If includeHeader Then
                                Dim headerCols As New List(Of String)(rdr.FieldCount)
                                For i As Integer = 0 To rdr.FieldCount - 1
                                    headerCols.Add(CsvFormat(rdr.GetName(i), delimiter, enclosure))
                                Next
                                resp.Write(String.Join(delimiter, headerCols))
                                resp.Write(newline)
                            End If

                            ' --- データ本体（逐次）
                            Dim rowCount As Integer = 0
                            While rdr.Read()
                                Dim cols As New List(Of String)(rdr.FieldCount)
                                For i As Integer = 0 To rdr.FieldCount - 1
                                    If rdr.IsDBNull(i) Then
                                        cols.Add("")
                                    Else
                                        Dim n = rdr.GetName(i)
                                        Dim t = rdr.GetFieldType(i)
                                        Dim v As Object = rdr.GetValue(i)
                                        Dim s As String
                                        Dim f As String = Nothing
                                        ' フィールド名 同じもののフォーマット情報を取得する
                                        If (formatEx IsNot Nothing) Then
                                            f = formatEx.Find(Function(x) x.name.ToUpper() = n.ToUpper()).format
                                        End If
                                        If t Is GetType(DateTime) Then
                                            f = If(f Is Nothing, "yyyy/MM/dd HH:mm:ss", f)
                                            s = CType(v, DateTime).ToString(f, CultureInfo.InvariantCulture)
                                        ElseIf GetType(IFormattable).IsAssignableFrom(t) Then
                                            f = If(f Is Nothing, "0.##########", f)
                                            s = DirectCast(v, IFormattable).ToString(f, CultureInfo.InvariantCulture)
                                        Else
                                            s = Convert.ToString(v, CultureInfo.InvariantCulture)
                                        End If
                                        cols.Add(CsvFormat(s, delimiter, enclosure))
                                    End If
                                Next

                                resp.Write(String.Join(delimiter, cols))
                                resp.Write(newline)

                                rowCount += 1
                                If (rowCount Mod 200) = 0 Then
                                    resp.Flush()
                                End If
                            End While

                            resp.Flush()
                            'If endResponse AndAlso ctx.ApplicationInstance IsNot Nothing Then
                            'ctx.ApplicationInstance.CompleteRequest()
                            'End If
                            HttpContext.Current.ApplicationInstance.CompleteRequest()
                            'resp.End()
                        End Using
                    End Using
                End Using
            Catch ex As ThreadAbortException
                ' ThreadAbortException は 無視する

                Dim m = ex.Message

            End Try
        End Sub

        '============================================
        ' ファイル転送
        '============================================
        'Public Sub FileTransfer(Response As HttpResponse, Server As HttpServerUtility, filename As String)
        '    Try
        '        'クライアント側に応答
        '        Response.Clear()
        '        Response.Buffer = True
        '        ' 2. コンテンツタイプの設定
        '        Response.ContentType = "application/octet-stream"
        '        ' ファイル名のエンコード（RFC 2231形式がより確実です）
        '        Dim fileNameOnly As String = IO.Path.GetFileName(filename)
        '        Dim encodedFileName As String = HttpUtility.UrlPathEncode(fileNameOnly)
        '        ' 3.HTTPヘッダー
        '        'Response.AddHeader("Content-Disposition", "attachment; filename*=UTF-8''" & encodedFileName)
        '        Response.AddHeader("Content-Disposition", "attachment; filename=" & encodedFileName)

        '        ' 4. ファイルサイズの設定 (オプションだが推奨)
        '        Dim fileInfo = New FileInfo(filename)
        '        Response.AddHeader("Content-Length", fileInfo.Length.ToString())

        '        Response.TransmitFile(filename)
        '        Response.Flush() ' 送信を確定させる
        '        Response.SuppressContent = True
        '        HttpContext.Current.ApplicationInstance.CompleteRequest() ' Response.Endの代わり（推奨）
        '        'Response.End()

        '    Catch ex As Exception
        '        ' ThreadAbortException は 無視する

        '        Dim m = ex.Message
        '    End Try
        'End Sub

        ' Yamaha robotex 内示受注ファイル読み込み 変換
        ''' <summary>
        ''' 横並びのCSVデータテーブルを、4列の縦並びデータテーブルに展開・変換します。
        ''' </summary>
        ''' <param name="filename">Yamaha robotex 内示受注データファイル</param>
        ''' <returns>部品番号、部品名称、日付、需要数 の4列で構成された新しいDataTable</returns>
        Public Function CreateNewDataSetFromCsv(filename As String) As DataTable

            ' CSVをDataTableに変換
            Dim dt = ConvertCsvToDataTable(filename)
            ' 条件に合う行だけを抽出
            Dim sourceDt = RemoveRowsWithoutThisTime(dt)


            ' 1. 固定の列名を定義
            Dim fixedColumns As New HashSet(Of String) From {"部品番号", "部品名称", "Column1", "日付"}

            ' 2. 元のDataTableの列名から、年月ヘッダー（可変）だけを自動抽出
            Dim dateHeaders As New List(Of String)()
            For Each col As DataColumn In sourceDt.Columns
                If Not fixedColumns.Contains(col.ColumnName) Then
                    dateHeaders.Add(col.ColumnName)
                End If
            Next

            ' 3. 新しいDataTableの構造（列）を作成
            Dim newDt As New DataTable()
            newDt.Columns.Add("部品番号", GetType(String))
            newDt.Columns.Add("部品名称", GetType(String))
            newDt.Columns.Add("日付", GetType(String))      ' "20250701" 等の形式
            newDt.Columns.Add("需要数", GetType(Integer))

            ' 元のデータが空の場合は、空の構造だけを返す
            If sourceDt.Rows.Count = 0 Then
                Return newDt
            End If

            ' 4. LINQの SelectMany を使って動的な年月列の数だけ行を展開
            Dim newRows = sourceDt.AsEnumerable().SelectMany(
        Function(row)
            Return dateHeaders.Select(
                Function(header)
                    Dim newRow As DataRow = newDt.NewRow()
                    newRow("部品番号") = row("部品番号")
                    newRow("部品名称") = row("部品名称")

                    ' 取得したヘッダー名（可変）に "01" を付加
                    newRow("日付") = header & "01"

                    ' 需要数の安全な数値化
                    Dim qty As Integer = 0
                    If row(header) IsNot DBNull.Value Then
                        Integer.TryParse(row(header).ToString(), qty)
                    End If
                    newRow("需要数") = qty

                    Return newRow
                End Function)
        End Function)

            ' 5. 展開した行データを新しいDataTableに追加
            For Each row As DataRow In newRows
                newDt.Rows.Add(row)
            Next

            ' 6. 変換後のDataTableを返す
            Return newDt
        End Function
        ''' <summary>
        '''  Yamaha Robotex 内示受注データ有効な"今回"データ以外を削除する
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function RemoveRowsWithoutThisTime(dt As DataTable) As DataTable

            ' 1. LINQで条件に合う行を特定
            Dim query = dt.AsEnumerable().Where(Function(row) row.Field(Of String)("Column1") = "今回")

            Dim filteredDt As DataTable

            ' 2. 該当する行があるかチェック
            If query.Any() Then
                ' 行がある場合はDataTableに変換
                filteredDt = query.CopyToDataTable()
            Else
                ' 1件もない場合は、構造（列）だけをコピーした空のDataTableを作成
                filteredDt = dt.Clone()
            End If
            Return filteredDt

        End Function
        ''' <summary>
        ''' csvファイルをDataTableに変換する
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <returns></returns>
        Public Function ConvertCsvToDataTable(filePath As String) As DataTable
            Dim dt As New DataTable()

            ' TextFieldParserの初期化
            Using parser As New TextFieldParser(filePath, System.Text.Encoding.GetEncoding("Shift_JIS"))
                ' カンマ区切りの設定
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(",")

                ' ダブルクォーテーションの囲みを考慮する設定
                parser.HasFieldsEnclosedInQuotes = True

                ' 1行目（ヘッダー）の読み込みと列作成
                If Not parser.EndOfData Then
                    Dim headers As String() = parser.ReadFields()
                    For Each header As String In headers
                        dt.Columns.Add(header)
                    Next
                End If

                ' 2行目以降（データ）の読み込み
                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()
                    dt.Rows.Add(fields)
                End While
            End Using

            Return dt
        End Function

        ''' <summary>
        ''' DataTableの内容をCSVファイルとして保存します（Shift_JIS、カンマ区切り、ダブルクォーテーション囲み対応）
        ''' </summary>
        ''' <param name="dt">保存対象のDataTable</param>
        ''' <param name="filePath">保存先のフルパス（例: "C:\Output\result.csv"）</param>
        Public Sub SaveDataTableToCsv(dt As DataTable, filePath As String)
            ' Shift_JISのエンコーディングを取得（.NET Core/.NET 5以降の場合は事前に Encoding.RegisterProvider(CodePagesEncodingProvider.Instance) が必要）
            Dim encoding As Encoding = Encoding.GetEncoding("Shift_JIS")

            Using sw As New StreamWriter(filePath, False, encoding)
                ' 1. ヘッダー（列名）の書き込み
                Dim headerLine As New List(Of String)()
                For Each col As DataColumn In dt.Columns
                    headerLine.Add(EscapeCsvField(col.ColumnName))
                Next
                sw.WriteLine(String.Join(",", headerLine))

                ' 2. データ行の書き込み
                For Each row As DataRow In dt.Rows
                    Dim fields As New List(Of String)()
                    For Each col As DataColumn In dt.Columns
                        fields.Add(EscapeCsvField(row(col).ToString()))
                    Next
                    sw.WriteLine(String.Join(",", fields))
                Next
            End Using
        End Sub

        ''' <summary>
        ''' CSVのフィールドとして安全な文字列にエスケープ・囲み処理を行います。
        ''' </summary>
        Private Function EscapeCsvField(field As String) As String
            If String.IsNullOrEmpty(field) Then
                Return """""" ' 空白はダブルクォーテーション2つ
            End If

            ' 値の中にダブルクォーテーションがある場合は「""」に置換
            If field.Contains("""") Then
                field = field.Replace("""", """""")
            End If

            ' カンマ、ダブルクォーテーション、改行が含まれる場合は全体を「"」で囲む
            If field.Contains(",") OrElse field.Contains("""") OrElse field.Contains(vbCr) OrElse field.Contains(vbLf) Then
                Return $"""{field}"""
            End If

            ' それ以外はそのまま「"」で囲んで返す（すべての値を一律囲む仕様にするとより安全です）
            Return $"""{field}"""
        End Function

        '============================================
        ' ファイル転送
        '============================================
        Public Sub FileTransfer2(Response As HttpResponse, Server As HttpServerUtility, filename As String)
            Try
                'クライアント側に応答
                Response.Clear()
                Response.Buffer = True
                ' 2. コンテンツタイプの設定
                Response.ContentType = "application/octet-stream"
                ' ファイル名のエンコード（RFC 2231形式がより確実です）
                Dim fileNameOnly As String = IO.Path.GetFileName(filename)
                Dim encodedFileName As String = HttpUtility.UrlPathEncode(fileNameOnly)
                ' 3.HTTPヘッダー
                'Response.AddHeader("Content-Disposition", "attachment; filename*=UTF-8''" & encodedFileName)
                Response.AddHeader("Content-Disposition", "attachment; filename=" & encodedFileName)

                ' 4. ファイルサイズの設定 (オプションだが推奨)
                Dim fileInfo = New FileInfo(filename)
                Response.AddHeader("Content-Length", fileInfo.Length.ToString())

                Response.TransmitFile(filename)
                Response.Flush() ' 送信を確定させる
                Response.SuppressContent = True
                HttpContext.Current.ApplicationInstance.CompleteRequest() ' Response.Endの代わり（推奨）
                'Response.End()


            Catch ex As Exception
                ' ThreadAbortException は 無視する

                Dim m = ex.Message
            End Try
        End Sub

        '============================================
        ' 複数ファイルをZip圧縮ファイルで転送
        '============================================

        ' DotNetZip ライブラリ (Ionic.Zip)
        Public Sub FilesTransfer(Response As HttpResponse, Server As HttpServerUtility, fileList As List(Of String), filename As String)
            Dim rt = ""
            ' 1ファイルの時
            If (fileList.Count = 1) Then
                FileTransfer2(Response, Server, fileList(0))
                Return
            End If

            Try
                ' 1. レスポンスバッファとヘッダーの初期化
                Response.Clear()
                Response.ClearHeaders()
                Response.ClearContent()

                ' ブラウザやOSとの互換性を高めるため、octet-stream を推奨
                Response.ContentType = "application/octet-stream"
                ' ファイル名にスペースが含まれる場合を考慮し、ダブルクォーテーションで囲む
                Response.AddHeader("Content-Disposition", $"attachment; filename=""{filename}""")

                Using zip As New Ionic.Zip.ZipFile(Encoding.GetEncoding("shift_jis"))
                    zip.AlternateEncoding = Encoding.GetEncoding("shift_jis")
                    zip.AlternateEncodingUsage = Ionic.Zip.ZipOption.AsNecessary

                    '' 必要に応じて自動でUnicode（UTF-8）を切り替える設定(推奨)
                    'zip.AlternateEncoding = System.Text.Encoding.UTF-8
                    'zip.AlternateEncodingUsage = Ionic.Zip.ZipOption.AsNecessary

                    ' 文字コード指定
                    'zip.AlternateEncoding = Encoding.GetEncoding("shift_jis")
                    'zip.AlternateEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                    'zip.AlternateEncodingUsage = Ionic.Zip.ZipOption.Always/AsNecessary/Never
                    ' 4G未満、4G 以上
                    'zip.UseZip64WhenSaving =  = Ionic.Zip.Zip64Option.AsNecessary/Always/Never

                    For Each file As String In fileList
                        ' ファイルを追加
                        zip.AddFile(file, "")
                    Next

                    ' 3. MemoryStreamを経由してシーク可能な状態でZIPを構築（重要）
                    ' ### for DEBUG
                    'Using ms As New FileStream("C:\Temp\TempFile.bin", FileMode.Create)
                    Using ms As New MemoryStream()
                        zip.Save(ms)    ' メモリ上に正しい構造でZIPを保存
                        ms.Position = 0 ' ストリームの位置を先頭に戻す

                        ' 4. クリーンなデータをResponseに書き出す
                        ms.CopyTo(Response.OutputStream)
                    End Using

                    ' バッファを出力
                    Response.Flush()
                End Using

                ' 5. ASP.NETの後続処理を安全に終了する
                HttpContext.Current.ApplicationInstance.CompleteRequest()

            Catch ex As Exception
                rt = ex.Message
            End Try
            'Return rt
        End Sub

        ''' <summary>
        ''' ファイル転送用 ファイル一覧をファイル化する
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="fileList"></param>
        ''' <returns></returns>
        Public Function SaveFileList(filename As String, fileList As List(Of String)) As String
            Dim erros = ""
            Try
                ' リストをファイルに保存 (上書き)
                File.WriteAllLines(filename, fileList)
            Catch ex As Exception
                erros = ex.Message
            End Try
            Return erros
        End Function
        ''' <summary>
        ''' ファイル転送用 ファイル一覧をファイルから読み込む
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <returns></returns>
        Public Function LoadFileList(filename As String) As List(Of String)
            Dim fileList As List(Of String) = New List(Of String)()
            Try
                If File.Exists(filename) Then
                    ' 全ての行を読み込み、List(Of String) に変換
                    fileList = File.ReadAllLines(filename).ToList()
                End If
            Catch ex As Exception
                Dim erros As String = ex.Message
            End Try
            Return fileList
        End Function
        ''' <summary>
        ''' Temporaly filename
        ''' </summary>
        ''' <param name="baseName"></param>
        ''' <returns></returns>
        Public Function GetTempFileName(baseName As String) As String

            Return baseName & DateTime.Now.ToString("yyyyMMddHHmmss") & ".tmp"

        End Function

        '"FileList.txt"
        '============================================
        ' リスト内の ファイル圧縮
        '============================================

        ' DotNetZip ライブラリ(Ionic.Zip)
        'Public Function CompressFiles(fileList As List(Of String), cprfilename As String) As String
        '    Dim rt = ""

        '    Try
        '        Using zip As New Ionic.Zip.ZipFile()
        '            For Each file As String In fileList
        '                ' ファイルを追加
        '                zip.AddFile(file, "")
        '            Next
        '            ' 保存
        '            zip.Save(cprfilename)
        '        End Using

        '    Catch ex As Exception
        '        rt = ex.Message
        '    End Try
        '    Return rt

        'End Function

        Public Function IsLocked(file As String) As Boolean
            If file Is Nothing Then Return False

            Dim fileinfo = New FileInfo(file)

            ' .Not() は拡張メソッドと推測されるため、標準的な Not 演算子で書き換えています
            If Not fileinfo.Exists Then Return False

            Try
                Using stream = fileinfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    ' オープンできればロックされていない
                End Using
            Catch
                ' 例外が発生すればロックされていると判断
                Return True
            End Try

            Return False
        End Function

        ''' <summary>
        ''' 指定されたパターン（例: "3-5-2", "5-5"）に従って客先品目Noを分割し、ハイフンで結合します。
        ''' </summary>
        ''' <param name="cstitemno">元の文字列</param>
        ''' <param name="pattern">ハイフン区切りのパターン（例: "3-5-2"）</param>
        Function FormatByCostmerItemNo(cstitemno As String, pattern As String) As String
            ' 安全チェック：入力が空の場合はそのまま返す
            If String.IsNullOrEmpty(cstitemno) OrElse String.IsNullOrEmpty(pattern) Then
                Return cstitemno
            End If

            ' パターン文字列を「-」で分解して数値の配列にする（例: "3-5-2" -> {3, 5, 2}）
            Dim lengthStrings As String() = pattern.Split("-"c)
            Dim parts As New List(Of String)()
            Dim currentIndex As Integer = 0

            ' 各ブロックの文字数ごとに切り出し処理
            For Each lenStr As String In lengthStrings
                Dim targetLength As Integer = 0

                ' 数字として正しく変換できるかチェック
                If Integer.TryParse(lenStr, targetLength) Then
                    ' 切り出す文字がまだ残っているかチェック
                    If currentIndex < cstitemno.Length Then
                        ' 残り文字数が必要数より少なければ、ある分だけ切り出す
                        Dim remainingLength As Integer = Math.Min(targetLength, cstitemno.Length - currentIndex)
                        parts.Add(cstitemno.Substring(currentIndex, remainingLength))
                        currentIndex += remainingLength
                    Else
                        ' 元の文字列をすべて使い切った場合はループを抜ける
                        Exit For
                    End If
                End If
            Next

            ' 切り出したパーツをハイフン「-」で結合して返す
            Return String.Join("-", parts)
        End Function

        '''' <summary>
        '''' 指定されたパターン（例: "3-5-2", "5-5"）に従って客先品目Noを分割し、ハイフンで結合します。
        '''' </summary>
        '''' <param name="cstitemno">元の文字列</param>
        '''' <param name="pattern">ハイフン区切りのパターン（例: "3-5-2"）</param>
        'Function FormatByCostmerItemNo(cstitemno As String, pattern As String) As String
        '    ' 安全チェック：入力が空の場合はそのまま返す
        '    If String.IsNullOrEmpty(cstitemno) OrElse String.IsNullOrEmpty(pattern) Then
        '        Return cstitemno
        '    End If

        '    ' パターン文字列を「-」で分解して数値の配列にする（例: "3-5-2" -> {3, 5, 2}）
        '    Dim lengthStrings As String() = pattern.Split("-"c)
        '    Dim parts As New List(Of String)()
        '    Dim currentIndex As Integer = 0

        '    ' 各ブロックの文字数ごとに切り出し処理
        '    For Each lenStr As String In lengthStrings
        '        Dim targetLength As Integer = 0

        '        ' 数字として正しく変換できるかチェック
        '        If Integer.TryParse(lenStr, targetLength) Then
        '            ' 切り出す文字がまだ残っているかチェック
        '            If currentIndex < cstitemno.Length Then
        '                ' 残り文字数が必要数より少なければ、ある分だけ切り出す
        '                Dim remainingLength As Integer = Math.Min(targetLength, cstitemno.Length - currentIndex)
        '                parts.Add(cstitemno.Substring(currentIndex, remainingLength))
        '                currentIndex += remainingLength
        '            Else
        '                ' 元の文字列をすべて使い切った場合はループを抜ける
        '                Exit For
        '            End If
        '        End If
        '    Next

        '    ' 切り出したパーツをハイフン「-」で結合して返す
        '    Return String.Join("-", parts)
        'End Function

#End Region

#Region "ドメイン変換（フォルダ区分など）"

        ''' <summary>フォルダ区分コード（Integer）から日本語表示へ変換。未知コードは "不明(コード)"。</summary>
        Public Function ToFolderTypeName(code As Integer) As String
            Dim name As String = Nothing
            If FolderTypeMap.TryGetValue(code, name) Then
                Return name
            End If
            Return $"不明({code})"
        End Function

        ''' <summary>フォルダ区分（文字列）から日本語表示へ安全に変換（Null/空/数値化不可は "未指定"）。</summary>
        Public Function ToFolderTypeNameSafe(value As Object) As String
            If value Is Nothing Then Return ToFolderTypeName(0)
            Dim s = value.ToString().Trim()
            If s.Length = 0 Then Return ToFolderTypeName(0)

            Dim code As Integer
            If Integer.TryParse(s, code) Then
                Return ToFolderTypeName(code)
            End If
            Return ToFolderTypeName(0)
        End Function

        ''' <summary>受注区分コード（Integer）から日本語表示へ変換。未知コードは "不明(コード)"。</summary>
        Public Function ToOrderTypeName(code As Integer) As String
            Dim name As String = Nothing
            If OrderTypeMap.TryGetValue(code, name) Then
                Return name
            End If
            Return $"不明({code})"
        End Function

        ''' <summary>受注区分（文字列）から日本語表示へ安全に変換（Null/空/数値化不可は "未指定"）。</summary>
        Public Function ToOrderTypeNameSafe(value As Object) As String
            If value Is Nothing Then Return ToOrderTypeName(0)
            Dim s = value.ToString().Trim()
            If s.Length = 0 Then Return ToOrderTypeName(0)

            Dim code As Integer
            If Integer.TryParse(s, code) Then
                Return ToOrderTypeName(code)
            End If
            Return ToOrderTypeName(0)
        End Function

        ''' <summary>情報区分コード（文字）から日本語表示へ変換。未知コードは "不明(コード)"。</summary>
        Public Function ToInfoTypeName(code As String) As String
            If String.IsNullOrWhiteSpace(code) Then
                Return "未指定"
            End If
            Dim name As String = Nothing
            If InfoTypeMap.TryGetValue(code.Trim(), name) Then
                Return name
            End If
            Return $"不明({code})"
        End Function

        ''' <summary>情報区分（Object）から日本語表示へ安全に変換（Null/空は "未指定"）。</summary>
        Public Function ToInfoTypeNameSafe(value As Object) As String
            If value Is Nothing Then Return "未指定"
            Dim s As String = value.ToString().Trim()
            If s.Length = 0 Then Return "未指定"
            Return ToInfoTypeName(s) ' 文字として解釈（I/U/D/N）
        End Function

        ''' <summary>権限レベル（Integer）から日本語表示へ変換。未知コードは "不明(コード)"。</summary>
        Public Function ToAuthorityLevelName(code As Integer) As String
            Dim name As String = Nothing
            If AuthorityLevelMap.TryGetValue(code, name) Then
                Return name
            End If
            Return $"不明({code})"
        End Function

        ''' <summary>権限レベル（文字列）から日本語表示へ安全に変換（Null/空/数値化不可は "未指定"）。</summary>
        Public Function ToAuthorityLevelNameSafe(value As Object) As String
            If value Is Nothing Then Return ToAuthorityLevelName(0)
            Dim s = value.ToString().Trim()
            If s.Length = 0 Then Return ToAuthorityLevelName(0)

            Dim code As Integer
            If Integer.TryParse(s, code) Then
                Return ToAuthorityLevelName(code)
            End If
            Return ToAuthorityLevelName(0)
        End Function

#End Region

#Region "管理者表示名"

        Public Function GetDisplayNameAdmin(loginUserId As String, loginPassword As String) As String
            Dim isUserId As Boolean = String.Equals(loginUserId, AdminUserID, StringComparison.OrdinalIgnoreCase)
            Dim isUserPass As Boolean = String.Equals(loginPassword, AdminUserPass, StringComparison.OrdinalIgnoreCase)
            If isUserId AndAlso isUserPass Then
                Return "admin"
            Else
                Return ""
            End If
        End Function

#End Region

#Region "HTML <option> 生成（DB・Web 非依存）"

        ''' <summary>
        ''' HTML の option をまとめて生成。
        ''' 値は HTML 属性用にエンコードして XSS を予防。
        ''' </summary>
        Public Function BuildOptions(values As IEnumerable(Of String)) As String
            If values Is Nothing Then Return String.Empty
            Dim sb As New StringBuilder()
            For Each v In values
                ' 値を安全に（属性値としてエンコード）
                Dim encoded = HttpUtility.HtmlAttributeEncode(v)
                sb.AppendFormat("<option value=""{0}"" />", encoded)
            Next
            Return sb.ToString()
        End Function

#End Region

    End Module

    '===============================
    ' PageHelpers
    '===============================
    Public Module PageHelpers

        ''' <summary>
        ''' セッションの UserName を指定ラベルに表示。
        ''' 無ければログイン画面へリダイレクト。
        ''' </summary>
        Public Sub SetUserName(ByVal page As Page, ByVal label As Label)
            Dim name As String = TryCast(page.Session("UserName"), String)

            If String.IsNullOrEmpty(name) Then
                WebTestSafe.SafeRedirect(page.Response, "~/Pages/Login/Login.aspx", False)
                WebTestSafe.SafeCompleteRequest()
                Exit Sub
            End If

            label.Text = page.Server.HtmlEncode(name)
        End Sub

        ''' <summary>
        ''' セッションの UserId を返す。無ければログインへリダイレクトし、空文字を返す。
        ''' </summary>
        Public Function GetUserId(ByVal page As Page) As String
            Dim userId As String = TryCast(page.Session("UserId"), String)

            If String.IsNullOrEmpty(userId) Then
                WebTestSafe.SafeRedirect(page.Response, "~/Pages/Login/Login.aspx", False)
                WebTestSafe.SafeCompleteRequest()
                Return String.Empty
            End If

            Return userId
        End Function

    End Module

    '===============================
    ' Web テスト用セーフラッパー
    '===============================
    Public Module WebTestSafe

        ''' <summary>
        ''' Response.Redirect の安全呼び出し。
        ''' - 本番：通常どおり Redirect / RedirectPermanent を使用
        ''' - テスト：NRE/HttpException 時は RedirectLocation/StatusCode を直接設定してフォールバック
        ''' </summary>
        Public Sub SafeRedirect(resp As HttpResponse, url As String, Optional endResponse As Boolean = False, Optional permanent As Boolean = False)
            If resp Is Nothing Then Throw New ArgumentNullException(NameOf(resp))
            If String.IsNullOrEmpty(url) Then url = "/"

            Try
                If permanent Then
                    ' .NET 4+ なら RedirectPermanent(url, endResponse) がある（古いFWは無い）
                    Dim mi = GetType(HttpResponse).GetMethod("RedirectPermanent", New Type() {GetType(String), GetType(Boolean)})
                    If mi IsNot Nothing Then
                        mi.Invoke(resp, New Object() {url, endResponse})
                    Else
                        ' 古い FW 向けフォールバック
                        resp.StatusCode = 301
                        resp.RedirectLocation = url
                    End If
                Else
                    ' 2引数版のみ使用（3引数は HttpResponseBase 側）
                    resp.Redirect(url, endResponse)
                End If

            Catch ex As Exception When TypeOf ex Is NullReferenceException OrElse TypeOf ex Is HttpException
                ' テスト環境で発生しがちな例外を吸収：Location を直接設定
                Try
                    resp.RedirectLocation = url
                    resp.StatusCode = If(permanent, 301, 302)
                Catch
                    ' 念のため握りつぶす（テスト時のみ）
                End Try
            End Try
        End Sub

        ''' <summary>
        ''' ApplicationInstance.CompleteRequest をテストでも安全に呼ぶためのラッパー。
        ''' 本番では通常どおり実行、テストで ApplicationInstance が Nothing の場合はスキップ。
        ''' </summary>
        Public Sub SafeCompleteRequest()
            Try
                Dim ctx = HttpContext.Current
                If ctx IsNot Nothing Then
                    Dim app = ctx.ApplicationInstance
                    If app IsNot Nothing Then
                        app.CompleteRequest()
                    End If
                End If
            Catch
                ' テスト環境での例外は握りつぶす
            End Try
        End Sub

    End Module


End Namespace