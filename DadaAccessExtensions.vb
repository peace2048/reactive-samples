Imports System
Imports System.Data
Imports System.Runtime.CompilerServices

Public Module DataAccessExtensions

  ''' <summary>
  ''' IDbCommand にパラメータを追加するメソッドを拡張します。
  ''' </summary>
  ''' <param name="command">拡張するコマンド</param>
  ''' <param name="parameterName">パラメータ名</param>
  ''' <param name="dbType">データタイプ</param>
  ''' <param name="size">データサイズ</param>
  ''' <param name="precision">有効桁数</param>
  ''' <param name="scale">少数部桁数</param>
  ''' <param name="direction">方向</param>
  ''' <returns>IDbDataParameter</returns>
  ''' <remarks></remarks>
  <Extension>
  Public Function AddParameter(ByVal command As IDbCommand,
                               ByVal parameterName As String,
                               Optional ByVal dbType As DbType? = Nothing,
                               Optional ByVal size As Integer? = Nothing,
                               Optional ByVal precision As Byte? = Nothing,
                               Optional ByVal scale As Byte? = Nothing,
                               Optional ByVal direction As ParameterDirection? = Nothing) As IDbDataParameter
  
    Dim p = command.CreateParameter()
    p.ParameterName = parameterName
    
    If dbType.HasValue Then p.DbType = dbType.Value
    If size.HasValue Then p.Size = size.Value
    If precision.HasValue Then p.Precision = precision.Value
    If scale.HasValue Then p.Scale = scale.Value
    If direction.HasValue Then p.Direction = direction.Value
    
    command.Parameters.Add(p)
    Return p
  End Function

End Module
