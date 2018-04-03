Option Explicit On 
Option Strict On

Imports System.Windows.Forms
Imports System.Xml.Serialization
Imports System.IO

Namespace DragD

Namespace QuickWhereComponent

#Region " Public Enums "

	Public Enum WhereConditions
		BeginsWith = 0
		Between = 1
		Contains = 2
		EndsWith = 3
		EqualTo = 4
		Greater = 5
		GreaterOrEqual = 6
		InCondition = 7
		IsNull = 8
		Less = 9
		LessOrEqual = 10
		LikeAs = 11
		NotBeginsWith = 12
		NotBetween = 13
		NotContains = 14
		NotEndsWith = 15
		NotEqualTo = 16
		NotInOperator = 17
		NotIsNull = 18
		NotLikeAs = 19
	End Enum

	Public Enum Operators
			OrOperator = 0
			AndOperator = 1
	End Enum

	Public Enum TypeOfValues
		StringType = 0
		NumericType = 1
		BooleanType = 2
		DateType = 3
	End Enum

#End Region

Public Class QuickWhere

	Inherits System.Collections.CollectionBase
	Implements IWhereEntity

#Region " Protected Members "

	Protected Friend Shared strStringDelimeter As Char
	Protected Friend Shared strEscapeChar As Char
	Protected Friend Shared strDateMarker As Char
	Protected Friend Shared strStringMarker As Char
	Protected Friend Shared strWildCardMarker As Char
	Protected Friend Shared strCharMarker As Char
	Protected Friend Shared strReplacement As String

	Protected Shared frm As Form
	Private Const vbQuote As Char = """"c

	Protected oOperator As Operators
	Protected bUsed As Boolean = True
#End Region


	' This constructor is required for the XML serialization
	' with no parameters 
		Public Sub New()

			MyBase.New()
			Call SetGenerals()
			oOperator = Operators.AndOperator
			bUsed = True

		End Sub

		Public Sub New(ByVal Operator As Operators)

			MyBase.New()
			Call SetGenerals()
			oOperator = Operator
			bUsed = True

		End Sub

		Default Property Item(ByVal Index As Integer) As WhereControl

		Get
				Return CType(List.Item(Index), WhereControl)
		End Get

		Set(ByVal Value As WhereControl)
				List(Index) = Value
		End Set

	End Property

#Region " IWhereEntity implementation "
		Public Property Used() As Boolean Implements IWhereEntity.Used
				Set(ByVal Value As Boolean)
						bUsed = Value
				End Set
				Get
						Return bUsed
				End Get
		End Property

		Public Function Idle(Optional ByVal WithValues As Boolean = False) As Boolean Implements IWhereEntity.Idle
			Return Me.WhereClause.Idle(WithValues)
		End Function

		Public Overloads Function GetSql() As String Implements IWhereEntity.GetSql
		Dim wc As WhereClause

			wc = Me.WhereClause(oOperator)
			If wc.Idle(True) Or (Not bUsed) Then
				Return ""
			Else
				Return wc.GetSql
			End If

			wc = Nothing

		End Function

#End Region

		Public Overloads Function GetSql(ByVal Operator As Operators) As String
			Dim wc As WhereClause

			wc = Me.WhereClause(Operator)
			Return wc.GetSql
			wc = Nothing

		End Function

		Public Function WhereClause(Optional ByVal Operator As Operators = Operators.AndOperator) As WhereClause
			' always get fresh new WhereClause
				Dim aWhereClause As New WhereClause(Operator)
				Dim aWhereControl As WhereControl
				Dim enumControls As IEnumerator

				enumControls = MyBase.GetEnumerator
				While enumControls.MoveNext
					aWhereControl = CType(enumControls.Current, WhereControl)
					' FillTheValues is required otherwise 
					' the aWhereControl will appear idle
					aWhereControl.FillTheValues()
					aWhereClause.Add(aWhereControl)
				End While

				Return aWhereClause

				aWhereClause = Nothing
				aWhereControl = Nothing

		End Function

#Region " Add Overloads "

			Public Overloads Function Add(ByVal aWhereControl As WhereControl) As Integer
				Return List.Add(aWhereControl)
			End Function

			Public Overloads Function Add(ByVal FieldName As String, _
																		ByVal aControl As Control) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl)

					aWhereControl.FieldName = FieldName
					Return List.Add(aWhereControl)

			End Function

	Public Overloads Function Add(ByVal FieldName As String, _
																ByVal ValueType As TypeOfValues, _
																ByVal aControl As Control) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl)

					With aWhereControl
						.FieldName = FieldName
						.ValueType = ValueType
					End With

					Return List.Add(aWhereControl)

			End Function

			Public Overloads Function Add(ByVal FieldName As String, _
																		ByVal WhereCondition As WhereConditions, _
																		ByVal aControl As Control) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl)

					With aWhereControl
						.FieldName = FieldName
						.WhereCondition = WhereCondition
					End With

					Return List.Add(aWhereControl)

			End Function

			Public Overloads Function Add(ByVal FieldName As String, _
																		ByVal WhereCondition As WhereConditions, _
																		ByVal ValueType As TypeOfValues, _
																		ByVal aControl As Control) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl)

					With aWhereControl
						.FieldName = FieldName
						.WhereCondition = WhereCondition
						.ValueType = ValueType
					End With

					Return List.Add(aWhereControl)

			End Function

			Public Overloads Function Add(ByVal FieldName As String, _
																		ByVal TableName As String, _
																		ByVal WhereCondition As WhereConditions, _
																		ByVal aControl As Control) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl)

					With aWhereControl
						.FieldName = FieldName
						.TableName = TableName
						.WhereCondition = WhereCondition
					End With

					Return List.Add(aWhereControl)

			End Function

			Public Overloads Function Add(ByVal FieldName As String, _
																		ByVal TableName As String, _
																		ByVal ValueType As TypeOfValues, _
																		ByVal WhereCondition As WhereConditions, _
																		ByVal aControl1 As Control, _
																		Optional ByVal aControl2 As Control = Nothing) As Integer

				Dim aWhereControl As WhereControl = New WhereControl(aControl1, aControl2)

					With aWhereControl
						.FieldName = FieldName
						.TableName = TableName
						.ValueType = ValueType
						.WhereCondition = WhereCondition
					End With

					Return List.Add(aWhereControl)

			End Function

#End Region

#Region " Generals and Shared "
			Public Property CharMarker() As Char
				Get
					Return strCharMarker
				End Get
				Set(ByVal Value As Char)
					strCharMarker = Value
				End Set
			End Property

			Public Property DateMarker() As Char
				Get
					Return strDateMarker
				End Get
				Set(ByVal Value As Char)
					strDateMarker = Value
				End Set
			End Property

			Public Property WildCardMarker() As Char
				Get
					Return strWildCardMarker
				End Get
				Set(ByVal Value As Char)
					strWildCardMarker = Value
				End Set
			End Property

			Public Property StringMarker() As Char
				Get
					Return strStringMarker
				End Get
				Set(ByVal Value As Char)
					strStringMarker = Value
				End Set
			End Property

			Public Property StringDelimeter() As Char
				Get
					Return strStringDelimeter
				End Get
				Set(ByVal Value As Char)
					strStringDelimeter = Value
				End Set
			End Property

			Public Property EscapeChar() As Char
				Get
					Return strEscapeChar
				End Get
				Set(ByVal Value As Char)
					strEscapeChar = Value
				End Set
			End Property

			Public Property Replacement() As String
				' ====== EXAMPLE ===========
				' The strReplacement in conjunction with
				' ValueOperation property of the WhereItem
				' works like this:

				'Dim wi As New WhereItem()
				'Call Globals.SetGenerals()
				'With wi
				'.FieldName = "F1"
				'.Value = "TestValue"
				'.ValueOperation = "Mid(@@,2,3)"
				'.SkipFieldNameBrackets = True
				'.SkipTableNameBrackets = True
				'.ValueType = WhereItem.TypeOfValues.StringType
				'.WhereCondition = QuickWhere.WhereConditions.EqualTo
				'MsgBox(wi.GetSql) 
				'End With

				' The msgbox gives the following:
				' F1 = Mid('TestValue',2,3)

				Get
					Return strReplacement
				End Get
				Set(ByVal Value As String)
					strReplacement = Value
				End Set
			End Property

		Public Shared Sub SetGenerals(Optional ByVal sStringMarker As Char = vbQuote, _
														Optional ByVal sDateMarker As Char = "#"c, _
														Optional ByVal sCharMarker As Char = "?"c, _
														Optional ByVal sWildCardMarker As Char = "*"c, _
														Optional ByVal sStringDelimeter As Char = ","c, _
														Optional ByVal sEscapeChar As Char = "\"c, _
														Optional ByVal sReplacement As String = "@@")

				strStringMarker = sStringMarker
				strDateMarker = sDateMarker
				strCharMarker = sCharMarker
				strWildCardMarker = sWildCardMarker
				strStringDelimeter = sStringDelimeter
				strEscapeChar = sEscapeChar
				strReplacement = sReplacement

		End Sub
#End Region

#Region " Serialization "

	Public Shared Sub SaveFilter(ByVal FileName As String, ByVal aQuickWhere As QuickWhere)
		Dim Writer As TextWriter = New StreamWriter(FileName)
		Dim X As XmlSerializer

			aQuickWhere.GetSql() ' it is necessary in to collect the values from the controls
													 ' and to fill the QuickWhere

			Try
					X = New XmlSerializer(GetType(QuickWhere))
					X.Serialize(Writer, aQuickWhere)
			Catch e As Exception
					Stop
					'Throw e
			Finally
					Writer.Close()
			End Try

	End Sub

	Public Shared Function LoadFilter(ByVal FileName As String, _
																		ByVal ParentForm As Form, _
																		Optional ByVal SetControlValues As Boolean = True) As QuickWhere

		Dim aFileStream As FileStream
		Dim X As XmlSerializer
		Dim Q As New QuickWhere()

		If Not File.Exists(FileName) Then
			Return Q
		End If

		Try
			aFileStream = New FileStream(FileName, FileMode.Open)
			X = New XmlSerializer(GetType(QuickWhere))

			Q.ParentForm = ParentForm
			Q = CType(X.Deserialize(aFileStream), QuickWhere)

			If SetControlValues Then
				Dim wc As WhereControl
				Dim enumControls As IEnumerator

					enumControls = Q.GetEnumerator
					While enumControls.MoveNext
						CType(enumControls.Current, WhereControl).SetControlValues()
					End While
			End If

			Return Q

		Catch e As Exception
			Throw e

		Finally
				aFileStream.Close()
		End Try

	 End Function

	 <System.Xml.Serialization.XmlIgnoreAttribute()> _
	Public Shared Property ParentForm() As Form
		Get
			Return frm
		End Get
		Set(ByVal Value As Form)
			frm = Value
		End Set
	End Property

#End Region

		Protected Overrides Sub Finalize()
					MyBase.Finalize()
				End Sub
	End Class

End Namespace

End Namespace