Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms
Imports System.Xml.Linq
Imports OMS.Common
Imports OMS.Data
Imports Oracle.ManagedDataAccess
Imports Oracle.ManagedDataAccess.Client

Module StraYamahaRobotexMonthlyProcessing


    'Yamaha Robotex の 毎月20日 にタスクマネージャ、スケジュール実行で実行するバッチ処理。
    '受注テーブルのレコード登録日が当月20日(非稼働日の場合前倒し日)以降のものを削除する処理。

    'アセンブリ名:STRA_YAMAHA_ROBOTEX_MONTHLY_PROCESSING
    'ルート名前空間:STRA_YAMAHA_ROBOTEX_MONTHLY_PROCESSING
    'フレームワーク: .NET Framework 4.8
    '種類:コンソールアプリケーション
    '環境
    '参照 アセンブリ
    'System
    'System.Core
    'System.Data
    'System.Data.DataSetExtensions
    'System.Deployment
    'System.Net.Http
    'System.Numeric
    'System.Windows.Forms
    'System.Xml
    'System.Xml.Linq
    '
    '参照 プロジェクト
    'OMS.Common
    'OMS.Data
    '
    ' スケジューラー設定

    Private Const LogFilename As String = "StraYamahaRobotexMonthlyProcessinglog.txt"

    Function Main() As Integer


        Dim rt As Integer = 0
        Dim errors As List(Of String) = New List(Of String)()
        ' 設定
        Dim xmlDoc As XDocument = XDocument.Load("Config.xml")
        Dim userId = ""
        Dim logPath = ""
        Dim _logger As Logger = Nothing
        Try
            For Each StraOrderBacklogUpdate In xmlDoc.Descendants("StraYamahaRobotexMonthlyProcessinglog")
                userId = StraOrderBacklogUpdate.Element("UserId")?.Value
                logPath = StraOrderBacklogUpdate.Element("LogPath")?.Value
            Next
            If (userId = "" Or logPath = "") Then
                Throw New Exception("Configration  Error")
            End If

            Dim repo = New OrderRepository(userId)
            Dim recl = New CalenderRepository(userId)
            Dim customerCode = "5799"
            Dim icalType = "00001"
            Dim nowOption As DateTime = DateTime.Now
            Dim firstDate = recl.AddWorkingDays2(icalType, New DateTime(nowOption.Year, nowOption.Month, 20), 0)
            Dim lastDate = recl.AddWorkingDays2(icalType, New DateTime(nowOption.Year, nowOption.Month, DateTime.DaysInMonth(nowOption.Year, nowOption.Month)), 0)

            ' 当月 20日(非稼働日の場合前倒し日)以降のものを削除する処理
            repo.YamahaRobotexMonthlyProcess(customerCode, firstDate, lastDate)

        Catch ex As Exception
            errors.Add(ex.Message)
            Dim ms As String = String.Join(vbCrLf, errors.Where(Function(x) x <> ""))
            _logger?.Write($"StraYamahaRobotexMonthlyProcessinglog.exe Error: {ms}")
            MessageBox.Show(ms, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rt = -1
        End Try

        Return rt


    End Function

End Module



'monthlyProcess
'SQL			
'DELETE FROM Orders			
'WHERE CREATED_AT >= :p_startDate
'  And CREATED_AT < :p_endDate +1           
'  And Customer_Code = '5977'			
'  And Customer_Order_No Not Like 'R%'			

'SQL
'引数でDateTime型(YYYYMM01) を渡し、その月の最終稼働日を取得する SQL	
'Select Case MAX(fDate) As last_working_day	
'FROM Calem	
'WHERE fHolidayFlag = 'W'	
'  And fDate >= TRUNC(TO_DATE(: input_date, 'YYYYMMDD'), 'MM')	
'  And fDate < ADD_MONTHS(TRUNC(TO_DATE(: input_date, 'YYYYMMDD'), 'MM'), 1);	


