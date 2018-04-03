Option Explicit On 
Option Strict On

Imports System.Xml.Serialization

Namespace DragD

Namespace QuickWhereComponent

Public Class WhereItem
	Implements IWhereEntity

	Protected strFieldName As String
	Protected strTableName As String
	Protected strValue As String
	Protected strValue2 As String
	Protected vTypeOfValues As TypeOfValues
	Protected wWhereOperator As WhereConditions
	Protected strValueOperation As String
	Protected strValue2Operation As String

	Protected bUsed As Boolean
	Protected bSkipFieldNameBrackets As Boolean
	Protected bSkipTableNameBrackets As Boolean

	Public Sub New()

		 strFieldName = ""
		 strTableName = ""
		 strValue = ""
		 strValue2 = ""
		 vTypeOfValues = TypeOfValues.StringType
		 wWhereOperator = WhereConditions.EqualTo
		 strValueOperation = ""
		 strValue2Operation = ""

		 bUsed = True
		 bSkipFieldNameBrackets = False
		 bSkipTableNameBrackets = False

		End Sub

	Public Property FieldName() As String
		Get
			Return strFieldName
		End Get
		Set(ByVal Value As String)
			strFieldName = Value
		End Set
	End Property

	Public Property TableName() As String
		Set(ByVal Value As String)
			strTableName = Value
		End Set
		Get
			Return strTableName
		End Get
	End Property

	Public Property Value() As String
		Get
			Return strValue
		End Get
		Set(ByVal ControlValue As String)
			strValue = ControlValue
		End Set
	End Property

	Public Property ValueOperation() As String
		Set(ByVal Value As String)
			strValueOperation = Value
		End Set
		Get
			Return strValueOperation
		End Get
	End Property

	Public Property Value2() As String
		Get
			Return strValue2
		End Get
		Set(ByVal ControlValue2 As String)
			strValue2 = ControlValue2
		End Set
	End Property

	Public Property Value2Operation() As String
		Set(ByVal Value As String)
			strValue2Operation = Value
		End Set
		Get
			Return strValue2Operation
		End Get
	End Property

	Public Property ValueType() As TypeOfValues
		Set(ByVal Value As TypeOfValues)
			vTypeOfValues = Value
		End Set
		Get
			Return vTypeOfValues
		End Get
	End Property

	Public Property WhereCondition() As WhereConditions
		Get
			Return wWhereOperator
		End Get
		Set(ByVal Value As WhereConditions)
			wWhereOperator = Value
		End Set
	End Property

	Public Function Copy() As WhereItem
		Dim TheCopy As New WhereItem()
		With TheCopy
			.SkipFieldNameBrackets = bSkipFieldNameBrackets
			.SkipTableNameBrackets = bSkipTableNameBrackets
			.FieldName = strFieldName
			.TableName = strTableName
			.ValueType = vTypeOfValues
			.Value = strValue
			.Value2 = strValue2
			.WhereCondition = wWhereOperator
			.ValueOperation = strValueOperation
			.Value2Operation = strValue2Operation
			.Used = bUsed
		End With
		Return TheCopy

	End Function

	Public Property SkipFieldNameBrackets() As Boolean
		Get
			Return bSkipFieldNameBrackets
		End Get
		Set(ByVal Value As Boolean)
			bSkipFieldNameBrackets = Value
		End Set
	End Property

	Public Property SkipTableNameBrackets() As Boolean
		Get
			Return bSkipTableNameBrackets
		End Get
		Set(ByVal Value As Boolean)
			bSkipTableNameBrackets = Value
		End Set
	End Property

#Region " IWhereEntity implementation "

	Public Property Used() As Boolean Implements IWhereEntity.Used
		Get
			Return bUsed
		End Get
		Set(ByVal Value As Boolean)
			bUsed = Value
		End Set
	End Property

	Public Function Idle(Optional ByVal WithValues As Boolean = False) As Boolean Implements IWhereEntity.Idle

			If strFieldName = "" Or IsNothing(strFieldName) Then
				Return True
			End If

			If wWhereOperator = WhereConditions.IsNull Or _
				 wWhereOperator = WhereConditions.NotIsNull Then
				 Return False
			End If

			If WithValues Then

					Dim v1IsIdle As Boolean
					Dim v2IsIdle As Boolean

						If (IsNothing(strValue)) Or (strValue = "") Then
							v1IsIdle = True
						Else
							v1IsIdle = False
						End If

						If (IsNothing(strValue2)) Or (strValue2 = "") Then
							v2IsIdle = True
						Else
							v2IsIdle = False
						End If

					Return v1IsIdle And v2IsIdle

			End If

	End Function

	Public Overridable Function GetSql() As String Implements IWhereEntity.GetSql
		Dim sWhere As New System.Text.StringBuilder()
		Dim sOperator As String
		Dim sValue As String
		Dim sValue2 As String

		If Not (bUsed) Or (Me.Idle) Then
			Return ""
		End If

		sValue = strValue
		sValue2 = strValue2

		If (sValue = "" And wWhereOperator <> WhereConditions.IsNull And wWhereOperator <> WhereConditions.NotIsNull) Or _
		 (sValue = "" And sValue2 = "" And (wWhereOperator = WhereConditions.Between Or wWhereOperator = WhereConditions.NotBetween)) Then
			Return ""
		End If

		Select Case wWhereOperator
			Case WhereConditions.EqualTo
				sOperator = " = "
			Case WhereConditions.NotEqualTo
				sOperator = " <> "
			Case WhereConditions.Less
				sOperator = " < "
			Case WhereConditions.Greater
				sOperator = " > "
			Case WhereConditions.LessOrEqual
				sOperator = " <= "
			Case WhereConditions.GreaterOrEqual
				sOperator = " >= "
			Case WhereConditions.IsNull
				sOperator = " IS NULL"
			Case WhereConditions.NotIsNull
				sOperator = " NOT IS NULL"
			Case WhereConditions.LikeAs, _
						WhereConditions.BeginsWith, _
						WhereConditions.EndsWith, _
						WhereConditions.Contains
				sOperator = " LIKE "
			Case WhereConditions.NotLikeAs, _
						WhereConditions.NotBeginsWith, _
						WhereConditions.NotEndsWith, _
						WhereConditions.NotContains
				sOperator = " NOT LIKE "
			Case WhereConditions.Between
				sOperator = " BETWEEN "
			Case WhereConditions.NotBetween
				sOperator = " NOT BETWEEN "
			Case WhereConditions.InCondition
				sOperator = " IN "
			Case WhereConditions.NotInOperator
				sOperator = " NOT IN "
		End Select

		If strTableName <> "" Then
			If bSkipTableNameBrackets Then
				sWhere.Append(strTableName & ".")
			Else
				sWhere.Append("[" & strTableName & "].")
			End If
		End If

		If bSkipFieldNameBrackets Then
			sWhere.Append(strFieldName)
		Else
			sWhere.Append("[" & strFieldName & "]")
		End If

		Select Case wWhereOperator
			Case WhereConditions.LikeAs, _
					 WhereConditions.NotLikeAs
				sWhere.Append(sOperator & GetProperValue(SetProperMarkers(sValue), vTypeOfValues, 1))

			Case WhereConditions.BeginsWith, _
					 WhereConditions.NotBeginsWith, _
					 WhereConditions.EndsWith, _
					 WhereConditions.NotEndsWith, _
					 WhereConditions.Contains, _
					 WhereConditions.NotContains
				sWhere.Append(sOperator & GetProperValue(GetValueForBEC(sValue, wWhereOperator), vTypeOfValues, 1))

			Case Else

				sValue = GetProperValue(sValue, vTypeOfValues, 1)
				If Not sValue2 Is Nothing Then
					sValue2 = GetProperValue(sValue2, vTypeOfValues, 2)
				End If

				If wWhereOperator = WhereConditions.InCondition Or _
					 wWhereOperator = WhereConditions.NotInOperator Then
					sValue = "(" & sValue & ")"
					sWhere.Append(sOperator & sValue)

				ElseIf wWhereOperator = WhereConditions.Between Then

					Select Case ProperBetween(sValue, sValue2)
						Case 0 : sWhere.Append(" = " & sValue)
						Case 1 : sWhere.Append(" >= " & sValue)
						Case 2 : sWhere.Append(" <= " & sValue2)
						Case 3 : sWhere.Append(sOperator & sValue & " AND " & sValue2)
					End Select

				ElseIf wWhereOperator = WhereConditions.NotBetween Then

					Select Case ProperBetween(sValue, sValue2)
						Case 0 : sWhere.Append(" <> " & sValue)
						Case 1 : sWhere.Append(" <= " & sValue)
						Case 2 : sWhere.Append(" >= " & sValue2)
						Case 3 : sWhere.Append(sOperator & sValue & " AND " & sValue2)
					End Select

				ElseIf wWhereOperator = WhereConditions.IsNull Or _
					wWhereOperator = WhereConditions.NotIsNull Then

					sWhere.Append(sOperator)

				Else

					sWhere.Append(sOperator & sValue)

				End If

		End Select

		Return sWhere.ToString

	End Function

#End Region

	Protected Function GetOperation(ByVal strOperation As String, _
																	ByRef strLeft As String, _
																	ByRef strRight As String) As Boolean

		Dim Position As Integer
		Dim ReplLen As Integer

		Position = InStr(1, strOperation, QuickWhere.strReplacement, vbTextCompare)

		If Position = 0 Then
			Return False
		Else
			ReplLen = Len(QuickWhere.strReplacement)
			strLeft = Left(strOperation, Position - 1)
			strRight = Mid(strOperation, Position + ReplLen)
			Return True
		End If

	End Function

	Protected Function GetProperValue(ByVal vValue As String, _
																		ByVal ValType As TypeOfValues, _
																		ByVal ValueNumber As Byte) As String
		Dim NewValue As String

		If Not (IsNothing(vValue)) Then

			Select Case ValType

				Case TypeOfValues.BooleanType
					Dim LeftChar As Char
					LeftChar = CType(UCase(Left(vValue, 1)), Char)

					If LeftChar = "T" Or _
						 LeftChar = "Y" Then
						NewValue = "-1"
					ElseIf LeftChar = "F" Or _
								 LeftChar = "N" Then
						NewValue = "0"
					Else
						NewValue = CType(IIf(CBool(Val(vValue)), "-1", "0"), String)
					End If

				Case TypeOfValues.DateType
					NewValue = ProcessDates(vValue)

				Case TypeOfValues.NumericType
					If vValue.ToString <> "" Then
						If vValue.IndexOf(","c) > 0 Then ' is it like "2,56,809" ?
							 NewValue = vValue
						Else
							 NewValue = CStr(Val(vValue))
						End If
					End If

				Case TypeOfValues.StringType
					NewValue = ProcessStrings(vValue)

				Case Else
					NewValue = vValue

			End Select

		Else
			NewValue = Nothing
		End If

		If Not NewValue Is Nothing Then
			Dim Op1Left As String
			Dim Op1Right As String
			Dim Op2Left As String
			Dim Op2Right As String

			If ValueNumber = 1 And GetOperation(strValueOperation, Op1Left, Op1Right) Then
				NewValue = Op1Left & CType(NewValue, String) & Op1Right
			End If

			If ValueNumber = 2 And GetOperation(strValue2Operation, Op2Left, Op2Right) Then
				NewValue = Op1Left & CType(NewValue, String) & Op1Right
			End If

		End If

		Return NewValue

	End Function

	Protected Function GetValueForBEC(ByVal sValue As String, ByVal BEC As WhereConditions) As String
		'BEC=Begins/Ends/Contains
		Dim sStr As String

		sStr = Replace(sValue, "*", "")
		sStr = Replace(sStr, "%", "")

		Select Case BEC
			Case WhereConditions.BeginsWith
				sStr = sStr & QuickWhere.strWildCardMarker
			Case WhereConditions.EndsWith
				sStr = QuickWhere.strWildCardMarker & sStr
			Case WhereConditions.Contains
				sStr = QuickWhere.strWildCardMarker & sStr & QuickWhere.strWildCardMarker
		End Select

		GetValueForBEC = sStr

	End Function

	Protected Function ProperBetween(ByVal strValueOne As String, _
																	 ByVal strValueTwo As String) As Byte

		If strValueOne.Equals(strValueTwo) Then
			Return 0
		ElseIf strValueOne <> "" And strValueTwo = "" Then
			Return 1
		ElseIf strValueOne = "" And strValueTwo <> "" Then
			Return 2
		ElseIf strValueOne <> "" And strValueTwo <> "" Then
			Return 3
		End If

	End Function

	Protected Function SetProperMarkers(ByVal sSql As String) As String
		Dim sVal As String

		sVal = Replace(sSql, "*", QuickWhere.strWildCardMarker)
		sVal = Replace(sVal, "%", QuickWhere.strWildCardMarker)

		sVal = Replace(sVal, "?", QuickWhere.strCharMarker)
		sVal = Replace(sVal, "_", QuickWhere.strCharMarker)
		Return sVal

	End Function

	Protected Function ProcessStrings(ByVal StringToProcess As String) As String
		Dim I As Integer
		Dim NewString As String
		Dim CurString As String

		If StringToProcess = Nothing Or _
			 StringToProcess = "" Then
			Return ""
		End If

		ProcessStrings = ""
		CurString = ""

		StringToProcess = Replace(StringToProcess, QuickWhere.strStringMarker, QuickWhere.strEscapeChar & QuickWhere.strStringMarker)

		' Converting the string "aa,bb,cc" into
		' "'aa','bb','cc'"  
		'	(where the strStringDelimeter is comma
		' and the strStringMarker is apostrophe)

		For I = 0 To CountToken(StringToProcess, QuickWhere.strStringDelimeter) - 1
			If Left(GetToken(StringToProcess, I), 1) <> QuickWhere.strStringMarker Then
				NewString = QuickWhere.strStringMarker & CStr(GetToken(StringToProcess, I))
			Else
				NewString = CStr(GetToken(StringToProcess, I))
			End If

			If Right(GetToken(StringToProcess, I), 1) <> QuickWhere.strStringMarker Then
				NewString = NewString & QuickWhere.strStringMarker
			End If
			CurString = CurString & QuickWhere.strStringDelimeter & NewString
		Next I
		Return Mid(CurString, 2)

	End Function

	Protected Function ProcessDates(ByVal strDate As String) As String
		Dim I As Integer
		Dim NewDate As String
		Dim CurDate As String
		Dim OrigValue As String

		If strDate Is Nothing Or strDate = "" Then
			Return ""
		End If

		ProcessDates = ""
		CurDate = ""
		OrigValue = strDate
		Try

			For I = 0 To CountToken(strDate) - 1

				If Left(GetToken(strDate, I), 1) = QuickWhere.strDateMarker Then
					NewDate = Mid(strDate, 2)
				Else
					NewDate = strDate
				End If

				If Right(GetToken(strDate, I), 1) = QuickWhere.strDateMarker Then
					NewDate = Mid(NewDate, 1, Len(NewDate) - 1)
				End If
				NewDate = QuickWhere.strDateMarker & CDate(NewDate) & QuickWhere.strDateMarker
				CurDate = CurDate & "," & NewDate
			Next I
			Return Mid(CurDate, 2)

		Catch e As Exception
			Throw e

		End Try

	End Function

#Region " Helpers "

	Private Sub StringToArray(ByVal Delimeted As String, _
															ByRef arrTokens() As String, _
															Optional ByVal Sep As Char = ","c, _
															Optional ByRef Lenght As Integer = 0)


			arrTokens = Delimeted.Split(Sep)
			Lenght = arrTokens.Length

	End Sub

	Private Function CountToken(ByVal Delimeted As String, _
											Optional ByVal Sep As Char = ","c) As Integer
			Dim arrTokens() As String

			Call StringToArray(Delimeted, arrTokens, Sep)
			Return arrTokens.Length

	End Function

	Private Function GetToken(ByVal Delimeted As String, _
															ByVal Pos As Integer, _
															Optional ByVal Sep As Char = ","c) As String

			Dim arrTokens() As String
			Dim lCount As Integer
			Dim I As Integer

			Call StringToArray(Delimeted, arrTokens, Sep, lCount)

			If Pos > (lCount - 1) Or _
					Pos < 0 Then
					ThrowException(System.Reflection.MethodInfo.GetCurrentMethod.Name, "Wrong Token Position Exception.")
					Exit Function
			End If

			GetToken = arrTokens(Pos)

	End Function

	Protected Sub ThrowException(ByVal strSource As String, ByVal strMessage As String)

			Dim e As System.Exception
			e = New Exception(strMessage)
			e.Source = strSource
			Throw e

	End Sub

#End Region

End Class

End Namespace

End Namespace