Option Explicit On 
Option Strict On

Namespace DragD

Namespace QuickWhereComponent

Public Class WhereClause

	Inherits CollectionBase
	Implements IWhereEntity

		Protected oOperator As Operators
		Protected bUsed As Boolean = True

		Public Function Copy() As WhereClause
			Dim TheCopy As New WhereClause()
			Dim eList As IEnumerator

			With TheCopy
				.Used = bUsed
				.Operator = oOperator
				eList = .GetEnumerator
			End With

			While eList.MoveNext
				TheCopy.List.Add(eList.Current)
			End While

			Return TheCopy

		End Function

		Public Sub New(ByVal Operator As Operators)
			oOperator = Operator
		End Sub

		' This constructor is required for the XML serialization	with no parameters
		Public Sub New()
			oOperator = Operators.AndOperator
		End Sub

		' This property is required for the XML serialization.
		' For the WhereClause class it must return Object
		' because the WhereClause could accept WhereItems and WhereClauses
		' This does not work with IWhereEntity
		Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Object
				Get
				Return List(index)
				End Get
		End Property

		Public Property Operator() As Operators
				Set(ByVal Value As Operators)
						oOperator = Value
				End Set
				Get
						Return oOperator
				End Get
		End Property

		Public Sub Remove(ByVal Index As Integer)
			List.Remove(Index)
		End Sub

#Region " Add Overloads "

		Private Overloads Function Add(ByVal WhereCondition As WhereConditions, _
							ByVal sFieldName As String, _
							Optional ByVal sTableName As String = "", _
							Optional ByVal strValue As String = "", _
							Optional ByVal strValue2 As String = "", _
							Optional ByVal ValueType As TypeOfValues = TypeOfValues.StringType, _
							Optional ByVal bIncludeWhereItem As Boolean = True) As Integer

		Dim WI As New WhereItem()
		With WI
			.WhereCondition = WhereCondition
			.FieldName = sFieldName

			If sTableName <> "" Then .TableName = sTableName
			If (Not strValue = "") Then .Value = strValue
			If (Not strValue2 = "") Then .Value2 = strValue2
			.ValueType = ValueType
		End With

		Return Add(WI)

	End Function


		Public Overloads Function Add(ByVal OneWhereItem As WhereItem) As Integer
					Return list.Add(OneWhereItem)
			End Function

		Public Overloads Function Add(ByVal OneWhereClause As WhereClause) As Integer
			Return list.Add(OneWhereClause)
		End Function

#End Region

#Region " IWhereEntity implementation "

	Public Property Used() As Boolean Implements IWhereEntity.Used
			Set(ByVal Value As Boolean)
					bUsed = Value
			End Set
			Get
					Return bUsed
			End Get
	End Property

	Public Overloads Function GetSql() As String Implements IWhereEntity.GetSql
		 Return GetSql(oOperator)
	End Function

	Public Overloads Function GetSql(ByVal Operator As Operators) As String
		Dim sWhere As String
		Dim sOperator As String
		Dim oOneWhere As IWhereEntity
		Dim oOneWhereSQL As String

		If Me.Idle(True) Or (Not bUsed) Then

			Return ""

		Else

			sWhere = ""

			If Operator = Operators.AndOperator Then
				sOperator = " AND "
			Else
				sOperator = " OR "
			End If

			For Each oOneWhere In List
				oOneWhereSQL = oOneWhere.GetSql
				If oOneWhereSQL <> "" Then
					sWhere = sWhere & sOperator & "(" & oOneWhereSQL & ")"
				End If
			Next oOneWhere

			sWhere = Mid(sWhere, Len(sOperator) + 1)
			Return sWhere

		End If
	End Function

	Public Function Idle(Optional ByVal WithValues As Boolean = False) As Boolean Implements IWhereEntity.Idle
		Dim enumColWhereClauseItems As IEnumerator = List.GetEnumerator
		Dim bIdle As Boolean = True

		While enumColWhereClauseItems.MoveNext
			 If Not CType(enumColWhereClauseItems.Current, IWhereEntity).Idle(WithValues) Then
					bIdle = False
					Exit While
			 End If
		End While

		Return bIdle

	End Function

#End Region

		Protected Overrides Sub Finalize()
				MyBase.Finalize()
			End Sub

End Class

End Namespace

End Namespace