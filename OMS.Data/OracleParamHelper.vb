Imports Oracle.ManagedDataAccess.Client
Imports System.Data

Namespace OMS.Data
    Public Module OracleParamHelper

        ' パラメータ名を :prefix + 小文字 に正規化
        Private Function Normalize(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("param name is empty.")
            Dim n = name.Trim()
            If Not n.StartsWith(":"c) Then n = ":" & n
            Return n.ToLowerInvariant()
        End Function

        Public Sub AddVarchar(cmd As OracleCommand, name As String, value As String)
            cmd.Parameters.Add(Normalize(name), OracleDbType.Varchar2).Value = value
        End Sub

        Public Sub AddVarcharOrNull(cmd As OracleCommand, name As String, value As String)
            cmd.Parameters.Add(Normalize(name), OracleDbType.Varchar2).Value =
                If(String.IsNullOrWhiteSpace(value), CType(DBNull.Value, Object), value)
        End Sub

        Public Sub AddChar(cmd As OracleCommand, name As String, value As String)
            cmd.Parameters.Add(Normalize(name), OracleDbType.Char).Value =
                If(String.IsNullOrEmpty(value), CType(DBNull.Value, Object), value)
        End Sub

        Public Sub AddIntOrNull(cmd As OracleCommand, name As String, value As String)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Int32)
            If String.IsNullOrWhiteSpace(value) Then
                p.Value = DBNull.Value
            Else
                Dim i As Integer
                p.Value = If(Integer.TryParse(value.Trim(), i), i, CType(DBNull.Value, Object))
            End If
        End Sub

        Public Sub AddDecimal(cmd As OracleCommand, name As String, value As String)
            Dim d As Decimal
            cmd.Parameters.Add(Normalize(name), OracleDbType.Decimal).Value =
                If(Decimal.TryParse(value, d), d, CType(DBNull.Value, Object))
        End Sub

        Public Sub AddOutDecimal(cmd As OracleCommand, name As String)
            cmd.Parameters.Add(Normalize(name), OracleDbType.Decimal, ParameterDirection.Output)
        End Sub

        Public Sub AddInt32(cmd As OracleCommand, name As String, value As Integer)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Int32)
            p.Value = value
        End Sub

        Public Sub AddInt32OrNull(cmd As OracleCommand, name As String, value As String)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Int32)
            If String.IsNullOrWhiteSpace(value) Then
                p.Value = DBNull.Value
            Else
                Dim i As Integer
                p.Value = If(Integer.TryParse(value.Trim(), Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, i),
                             i, CType(DBNull.Value, Object))
            End If
        End Sub

        ' 64bit 整数（IDの上限が32bitを超える可能性に備える）
        Public Sub AddInt64(cmd As OracleCommand, name As String, value As Long)
            ' 注意：OracleDbType.Int64 は ODP.NET では Decimal にマップされることがあります。
            ' ID が NUMBER(19,0) 以内なら Int64 でOK。それ以上の桁数は Decimal を推奨。
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Int64)
            p.Value = value
        End Sub

        ' NUMBER(38,0) 相当で安全に行く場合（ID や大きい整数）
        Public Sub AddNumberInteger(cmd As OracleCommand, name As String, value As Decimal)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Decimal)
            p.Precision = 38
            p.Scale = 0
            p.Value = value
        End Sub

        Public Sub AddNumberIntegerOrNull(cmd As OracleCommand, name As String, value As String)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Decimal)
            p.Precision = 38
            p.Scale = 0
            If String.IsNullOrWhiteSpace(value) Then
                p.Value = DBNull.Value
            Else
                Dim d As Decimal
                ' 整数のみに限定（小数は受け付けない）。InvariantCultureで安全に。
                If Decimal.TryParse(value.Trim(), Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, d) Then
                    p.Value = d
                Else
                    p.Value = DBNull.Value
                End If
            End If
        End Sub

        Public Sub AddTimeStamp(cmd As OracleCommand, name As String, value As DateTime)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.TimeStamp)
            p.Value = value
        End Sub

        Public Sub AddTimeStampOrNull(cmd As OracleCommand, name As String, value As DateTime?)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.TimeStamp)
            p.Value = If(value.HasValue, CType(value.Value, Object), DBNull.Value)
        End Sub

        Public Sub AddClob(cmd As OracleCommand, name As String, value As String)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Clob)
            p.Value = If(value Is Nothing, CType(DBNull.Value, Object), value)
        End Sub

        Public Sub AddBlob(cmd As OracleCommand, name As String, value As Byte())
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Blob)
            p.Value = If(value Is Nothing, CType(DBNull.Value, Object), value)
        End Sub

        ' 文字列をトリムしてNULL化したいユースケースが多いなら共通化
        Public Sub AddVarcharTrimOrNull(cmd As OracleCommand, name As String, value As String)
            Dim p = cmd.Parameters.Add(Normalize(name), OracleDbType.Varchar2)
            Dim v = If(value, String.Empty).Trim()
            p.Value = If(String.IsNullOrEmpty(v), CType(DBNull.Value, Object), v)
        End Sub

        Public Sub AddDate(cmd As OracleCommand, name As String, value As Date)
            'Dim d As Decimal
            cmd.Parameters.Add(Normalize(name), OracleDbType.Date).Value = value
        End Sub

    End Module
End Namespace