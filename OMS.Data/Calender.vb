
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Bibliography
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Drawing.Diagrams
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OMS.Common
Imports Oracle.ManagedDataAccess.Client

Namespace OMS.Data

    Public Class CalenderRepository
        Private ReadOnly _connectionString As String

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        ''' <summary>
        ''' 指定の年月日 から 指定日 前/後 の稼働日を取得する(OMSDB.Function.ADD_WORKING_DAYS2使用)
        ''' iCaletype は "00001" を指定します
        ''' iDays (オフセット日数)で 0 を指定し、iDate が非稼働日の場合 前倒し して稼働日を返します。
        ''' 2026/6/29
        ''' </summary>
        ''' <param name="iCaleTyp"></param>
        ''' <param name="iDate"></param>
        ''' <param name="iDays"></param>
        ''' <returns>Date</returns>
        Public Function AddWorkingDays(iCaleTyp As String, iDate As Date, iDays As Integer) As Date
            Dim tdate As Date
            Dim piCaleTyp As String = Utils.BuildLikePattern(iCaleTyp, LikeMode.Contains)

            Using conn As New OracleConnection(_connectionString)
                conn.Open()
                Using tran As OracleTransaction = conn.BeginTransaction()
                    Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                        cmd.CommandText = "SELECT
                                   ADD_WORKING_DAYS2(:p_iCaleTyp, :p_iDate, :p_iDays) 
                                   FROM DUAL "
                        cmd.Parameters.Add(":p_iCaleTyp", OracleDbType.Char, 20).Value = piCaleTyp
                        cmd.Parameters.Add(":p_iDate", OracleDbType.Date).Value = iDate
                        cmd.Parameters.Add(":p_iDays", OracleDbType.Int16).Value = iDays
                        tdate = cmd.ExecuteScalar()
                    End Using
                    tran.Commit()
                End Using
                'conn.Close()
            End Using
            Return tdate
        End Function
        ''' <summary>
        ''' 指定の年月日 から 指定日 前/後 の稼働日を取得する(OMSDB.Function.ADD_WORKING_DAYS使用)
        ''' 2026/6/25
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="iCaleTyp"></param>
        ''' <param name="iDate"></param>
        ''' <param name="iDays"></param>
        ''' <returns>Date</returns>
        Public Function AddWorkingDays(conn As OracleConnection, tran As OracleTransaction, iCaleTyp As String, iDate As Date, iDays As Integer) As Date
            Dim tdate As Date
            Dim piCaleTyp As String = Utils.BuildLikePattern(iCaleTyp, LikeMode.Contains)
            Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                cmd.CommandText = "SELECT
                                   ADD_WORKING_DAYS(:p_iCaleTyp, :p_iDate, :p_iDays) 
                                   FROM DUAL "
                cmd.Parameters.Add(":p_iCaleTyp", OracleDbType.Char, 20).Value = piCaleTyp
                cmd.Parameters.Add(":p_iDate", OracleDbType.Date).Value = iDate
                cmd.Parameters.Add(":p_iDays", OracleDbType.Int16).Value = iDays
                tdate = cmd.ExecuteScalar()

            End Using
            Return tdate

        End Function
        ''' <summary>
        ''' 指定の年月日 が稼働日かどうか(OMSDB.Function.IS_WORKING_DAY)
        ''' 2026/6/25
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="iCaleTyp"></param>
        ''' <param name="iDate"></param>
        ''' <returns></returns>
        Public Function IsWorkingDays(conn As OracleConnection, tran As OracleTransaction, iCaleTyp As String, iDate As Date) As Boolean
            Dim judge As Boolean = False
            Dim piCaleTyp As String = Utils.BuildLikePattern(iCaleTyp, LikeMode.Contains)
            Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                cmd.CommandText = "SELECT
                                   IS_WORKING_DAY(:p_iCaleTyp, :p_iDate) 
                                   FROM DUAL "
                cmd.Parameters.Add(":p_iCaleTyp", OracleDbType.Char, 20).Value = piCaleTyp
                cmd.Parameters.Add(":p_iDate", OracleDbType.Date).Value = iDate
                Dim jg = cmd.ExecuteScalar()
                judge = jg = "Y"
            End Using
            Return judge
        End Function
        ''' <summary>
        ''' 指定の年月日 が非稼働日の場合 前方検索して最初の稼働日を返す(OMSDB.Function.IS_WORKING_DAY/OMSDB.Function.ADD_WORKING_DAYS使用)
        ''' 2026/6/25
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="iCaleTyp"></param>
        ''' <param name="iDate"></param>
        ''' <returns></returns>
        Public Function SearchForwardWorkingDays(conn As OracleConnection, tran As OracleTransaction, iCaleTyp As String, iDate As Date, iDays As Integer) As Date

            Return SearchForwardWorkingDays(conn, tran, iCaleTyp, iDate.AddDays(iDays))

        End Function
        ''' <summary>
        ''' 指定の年月日 が非稼働日の場合 前方検索して最初の稼働日を返す(OMSDB.Function.IS_WORKING_DAY/OMSDB.Function.ADD_WORKING_DAYS使用)
        ''' 2026/6/25
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="iCaleTyp"></param>
        ''' <param name="iDate"></param>
        ''' <returns></returns>
        Public Function SearchForwardWorkingDays(conn As OracleConnection, tran As OracleTransaction, iCaleTyp As String, iDate As Date) As Date
            ' 指定日が稼働日かどうかを判定する
            If (IsWorkingDays(conn, tran, iCaleTyp, iDate)) Then
                Return iDate
            End If
            ' 稼働日を前方検索する
            Return AddWorkingDays(conn, tran, iCaleTyp, iDate, -1)
        End Function

        ''' <summary>
        ''' 期間指定で 休日/営業日 日数を取得する
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="dtStart"></param>
        ''' <param name="dtEnd"></param>
        ''' <param name="holiday"></param>
        ''' <returns></returns>
        Public Function NumberOfHolidayDurPeriod(conn As OracleConnection, tran As OracleTransaction, dtStart As Date, dtEnd As Date, holiday As Boolean) As Integer
            Dim cnt As Integer
            Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}
                cmd.CommandText = "
                        SELECT Count (*)
                             Fdate As ""Fdate"",  
                             Fholidayflg As ""Holidayflg""  
                        FROM Calem
                        WHERE Fdate >= to_date(:p_startDate, 'yyyy-mm-dd') 
                        AND Fdate <= to_date(:p_endDate, 'yyyy-mm-dd') 
                        AND Fholidayflg = :p_holidayflg  "
                Dim startDate = dtStart.ToString("yyyy-MM-dd")
                Dim endDate = dtEnd.ToString("yyyy-MM-dd")
                Dim holidayFlag = If(holiday, "H", "W")
                cmd.Parameters.Add(":p_startDate", OracleDbType.Char, 20).Value = startDate
                cmd.Parameters.Add(":p_endDate", OracleDbType.Char, 20).Value = endDate
                cmd.Parameters.Add(":p_holidayflg", OracleDbType.Char, 1).Value = holidayFlag

                cnt = cmd.ExecuteScalar()

            End Using
            Return cnt

        End Function
        ''' <summary>
        ''' 指定日から前方の 稼働日を探す
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="tran"></param>
        ''' <param name="dtTarget"></param>
        ''' <returns></returns>
        Public Function GetWorkingDayDescendingOrder(conn As OracleConnection, tran As OracleTransaction, dtTarget As DateTime) As DateTime
            Dim cnt As DateTime

            Using cmd As New OracleCommand() With {.Connection = conn, .BindByName = True}

                cmd.CommandText = "SELECT * FROM ( 
                                SELECT * 
                                FROM Calender 
                                WHERE TARGET_DATE <= TO_DATE(:p_targetDate, 'YY-MM-DD') 
                                  AND FHOLIDAYFLG = 'W' 
                                ORDER BY TARGET_DATE DESC  
                             )
                             WHERE ROWNUM = 1 "
                Dim targetDate = dtTarget.ToString("yyyy-MM-dd")
                cmd.Parameters.Add(":p_targetDate", OracleDbType.Char, 20).Value = targetDate
                cnt = cmd.ExecuteScalar()

            End Using
            Return cnt

        End Function

        Public Enum CaptureDef
            All = 0
            Holiday = 1
            BusinessDay = 2
        End Enum

        ' 生産計画条件マスタ一覧取得
        Public Function GetCalenderList(conn As OracleConnection, tran As OracleTransaction, dtStart As Date, dtEnd As Date) As List(Of CalenderRow)

            Dim cl = GetCalenderList(conn, tran, dtStart, dtEnd, CaptureDef.All)
            Dim calenderRowList = ToClass(cl)
            Return calenderRowList
        End Function

        ' 生産計画条件マスタ一覧取得
        Public Function GetCalenderList(conn As OracleConnection, tran As OracleTransaction, dtStart As Date, dtEnd As Date, capture As CaptureDef) As DataTable

            Dim dt As New DataTable()
            Dim sb As New StringBuilder()
            sb.AppendLine("SELECT ")
            sb.AppendLine(" fcaletyp          AS ""CalenderType"", ")
            sb.AppendLine(" fdate             AS ""DefDate"", ")
            sb.AppendLine(" fdteind           AS ""Dteind"", ")
            sb.AppendLine(" fholidayflg       AS ""HolidayFlag"", ")
            sb.AppendLine(" fbuckno           AS ""BucketNo"", ")
            sb.AppendLine(" fholidaytyp       AS ""HolidayType"", ")
            sb.AppendLine(" fentdte           AS ""EntryDate"", ")
            sb.AppendLine(" fentusr           AS ""EntryUser"", ")
            sb.AppendLine(" fupddte           AS ""UpdateDate"", ")
            sb.AppendLine(" fupdusr           AS ""UpdateUser"", ")
            sb.AppendLine(" fupdprg           AS ""UpdateProgramID"", ")
            sb.AppendLine(" fcacheindex       AS ""ChashIndex"", ")
            sb.AppendLine(" fdateseq          AS ""DateSequence"", ")
            sb.AppendLine(" fbefdayholiday    AS ""BeforeHoliday"", ")
            sb.AppendLine(" ffactor           AS ""Factor"", ")
            sb.AppendLine(" ftmtyp            AS ""TmType""")
            sb.AppendLine("FROM calem ")
            sb.AppendLine("WHERE 1=1 ")
            Dim prm As New List(Of OracleParameter)()

            Dim startDate = dtStart.ToString("yyyy-MM-dd")
            Dim endDate = dtEnd.ToString("yyyy-MM-dd")
            Dim holidayFlag As String = ""
            If (capture = CaptureDef.All) Then
            ElseIf (capture = CaptureDef.Holiday) Then
                holidayFlag = "H"
            ElseIf (capture = CaptureDef.BusinessDay) Then
                holidayFlag = "W"
            End If
            If capture <> CaptureDef.All Then
                sb.AppendLine("AND fholidayflg = :p_holidayflg ")
                prm.Add(New OracleParameter(":p_holidayflg", OracleDbType.Char, 1) With {.Value = holidayFlag})
            End If
            sb.AppendLine("AND fdate >= to_date( :p_startDate, 'yyyy-mm-dd' ) ")
            prm.Add(New OracleParameter(":p_startDate", OracleDbType.Varchar2) With {.Value = startDate})
            sb.AppendLine("AND fdate <= to_date( :p_endDate, 'yyyy-mm-dd' ) ")
            prm.Add(New OracleParameter(":p_endDate", OracleDbType.Varchar2) With {.Value = endDate})
            sb.AppendLine("ORDER BY fdate ")
            Try
                Using cmd As New OracleCommand(sb.ToString(), conn)
                    cmd.BindByName = True
                    cmd.CommandType = CommandType.Text
                    If prm.Count > 0 Then cmd.Parameters.AddRange(prm.ToArray())
                    'conn.Open()
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        'Dim reader As OracleDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            Catch ex As Exception
                Dim m = ex.Message
            End Try

            Return dt

        End Function
        ''' <summary>
        ''' CalenderRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataRow) As CalenderRow
            If dt Is Nothing Then
                Return Nothing
            End If

            Dim cr = New CalenderRow
            cr.CalenderType = dt.Field(Of String)("CalenderType")
            cr.DefDate = dt.Field(Of Date?)("DefDate")
            cr.Dteind = dt.Field(Of Int16?)("Dteind")
            cr.HolidayFlag = dt.Field(Of String)("HolidayFlag")
            cr.BucketNo = dt.Field(Of String)("BucketNo")
            cr.HolidayType = dt.Field(Of String)("HolidayType")
            cr.EntryDate = dt.Field(Of Date?)("EntryDate")
            cr.UpdateDate = dt.Field(Of Date?)("UpdateDate")
            cr.UpdateUser = dt.Field(Of String)("UpdateUser")
            cr.UpdateProgramID = dt.Field(Of String)("UpdateProgramID")
            cr.ChashIndex = dt.Field(Of String)("ChashIndex")
            cr.DateSequence = dt.Field(Of Integer?)("DateSequence")
            cr.BeforeHoliday = dt.Field(Of Long?)("BeforeHoliday")
            cr.Factor = dt.Field(Of Double?)("Factor")
            cr.TmType = dt.Field(Of String)("TmType")
            Return cr

        End Function
        ''' <summary>
        ''' CalenderRow to class
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        Public Function ToClass(dt As DataTable) As List(Of CalenderRow)
            Dim osrs = New List(Of CalenderRow)()
            For Each dtRow In dt.Rows
                osrs.Add(ToClass(dtRow))
            Next
            Return osrs
        End Function

    End Class

    Public Class CalenderRow
        Public Property CalenderType As String      ' FCALETYP          カレンダー               CHAR(9)
        Public Property DefDate As Date?            ' FDATE             日                       DATE
        Public Property Dteind As Int16?            ' FDTEIND           カレンダー旬             NUMBER (0)
        Public Property HolidayFlag As String       ' FHOLIDAYFLG       休日フラグ               CHAR(1)
        Public Property BucketNo As String          ' FBUCKNO           バケットNo               CHAR(8)
        Public Property HolidayType As String       ' FHOLIDAYTYP       休日タイプ               CHAR(1)
        Public Property EntryDate As Date?          ' FENTDTE           登録日時                 DATE
        Public Property EntryUser As String         ' FENTUSR           登録ユーザーID           VCHAR2(9)
        Public Property UpdateDate As Date?         ' FUPDDTE           更新日時                 DATE
        Public Property UpdateUser As String        ' FUPDUSR           更新ユーザーID           VCHAR2(9)
        Public Property UpdateProgramID As String   ' FUPDPRG           更新プログラムID         VCHAR2(150)
        Public Property ChashIndex As String        ' FCACHEINDEX       キャッシュインデックス   CHAR(36)
        Public Property DateSequence As Long?       ' FDATESEQ          データ順序               NUMBER(6,0)
        Public Property BeforeHoliday As Long?      ' FBEFDAYHOLIDAY    前日マデノ連休日数       NUMBER(10,0)
        Public Property Factor As Double?           ' FFACTOR           出荷係数                 NUMBER(12,4)
        Public Property TmType As String            ' FTMTYP            日別ＴＭタイプ           CHAR(9)
    End Class
End Namespace
