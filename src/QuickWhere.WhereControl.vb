Option Explicit On 
Option Strict On

Imports System.Windows.Forms
Namespace DragD

Namespace QuickWhereComponent

Public Class WhereControl
		Inherits WhereItem

		Protected obControl1 As Control
		Protected obControl2 As Control
		Protected bSelectedValueIsUsed As Boolean
		Protected bSelectionEndIsUsed As Boolean

		Public Sub New()

				MyBase.New()
				obControl1 = Nothing
				obControl2 = Nothing
				bSelectedValueIsUsed = True
				bSelectionEndIsUsed = False

		End Sub

		Public Sub New(ByVal aControl1 As Control, _
									 Optional ByVal aControl2 As Control = Nothing, _
									 Optional ByVal aSelectedValueIsUsed As Boolean = False, _
									 Optional ByVal aSelectionEndIsUsed As Boolean = False)

				MyBase.New()
				obControl1 = aControl1
				obControl2 = aControl2

				MyBase.Value = GetControlText(obControl1)
				MyBase.Value2 = GetControlText(obControl2)

				bSelectedValueIsUsed = aSelectedValueIsUsed
				bSelectionEndIsUsed = aSelectionEndIsUsed

		End Sub

#Region " Public Properties "

		Public Property Control1() As String
			Get
				If Not (obControl1 Is Nothing) Then
					Return obControl1.Name
				Else
					Return ""
				End If
			End Get
			Set(ByVal Value As String)
				obControl1 = GetControlByName(Value)
			End Set
		End Property

		Public Property Control2() As String
			Get
				If Not (obControl2 Is Nothing) Then
					Return obControl2.Name
				Else
					Return ""
				End If
			End Get
			Set(ByVal Value As String)
				obControl2 = GetControlByName(Value)
			End Set
		End Property

		Public Property SelectedValueIsUsed() As Boolean
			Get
				Return bSelectedValueIsUsed
			End Get
			Set(ByVal aBoolean As Boolean)
				bSelectedValueIsUsed = aBoolean
			End Set
		End Property


#End Region

#Region " Serialization Helpers "
		Protected Function GetControlText(ByRef obControl As Control, _
																			Optional ByVal aSelectionEndIsUsed As Boolean = False) As String

			Try

					If obControl Is Nothing Then
						Dim str As String = MyBase.Value

							If str <> "" Then
								Return str
							Else
								Return ""
							End If

							Exit Function

					End If

					If TypeOf obControl Is TextBox Then
							Return CType(obControl, TextBox).Text

					ElseIf TypeOf obControl Is ComboBox Then
							If bSelectedValueIsUsed Then
								Return CType(CType(obControl, ComboBox).SelectedValue, String)
							Else
								Return CType(obControl, ComboBox).Text
							End If

					ElseIf TypeOf obControl Is ListBox Then
							If bSelectedValueIsUsed Then
								Return CType(CType(obControl, ListBox).SelectedValue, String)
							Else
								Return CType(obControl, ListBox).Text
							End If

					ElseIf TypeOf obControl Is CheckedListBox Then
							If bSelectedValueIsUsed Then
								Return CType(CType(obControl, CheckedListBox).SelectedValue, String)
							Else
								Return CType(obControl, CheckedListBox).Text
							End If

					ElseIf TypeOf obControl Is CheckBox Then
							Return CType(obControl, CheckBox).Checked.ToString

					ElseIf TypeOf obControl Is RadioButton Then
							Return CType(obControl, RadioButton).Checked.ToString

					ElseIf TypeOf obControl Is DateTimePicker Then
							Return CType(obControl, DateTimePicker).Value.ToString

					ElseIf TypeOf obControl Is MonthCalendar Then
							If aSelectionEndIsUsed Then
								Return CType(obControl, MonthCalendar).SelectionEnd.ToString
							Else
								Return CType(obControl, MonthCalendar).SelectionStart.ToString
							End If

					ElseIf TypeOf obControl Is TrackBar Then
							Return CType(obControl, TrackBar).Value.ToString

					ElseIf TypeOf obControl Is DomainUpDown Then
							Return CType(obControl, DomainUpDown).Text

					ElseIf TypeOf obControl Is NumericUpDown Then
							Return CType(obControl, NumericUpDown).Value.ToString

					Else
							Return ""	 ' ignore other types of controls
					End If
			Catch e As Exception
				Throw e
			End Try

		End Function

		Protected Sub SetControlText(ByRef obControl As Control, _
																 ByVal aString As String, _
																 Optional ByVal aSelectionEndIsUsed As Boolean = False)

					If obControl Is Nothing Then Exit Sub

					If aString = "" Then aString = Nothing
			Try
					If TypeOf obControl Is TextBox Then
							CType(obControl, TextBox).Text = aString

					ElseIf TypeOf obControl Is ComboBox Then
							If bSelectedValueIsUsed Then
								CType(obControl, ComboBox).SelectedValue = aString
							Else
								CType(obControl, ComboBox).Text = aString
							End If

					ElseIf TypeOf obControl Is ListBox Then
							If bSelectedValueIsUsed Then
								CType(obControl, ListBox).SelectedValue = aString
							Else
								CType(obControl, ListBox).Text = aString
							End If

					ElseIf TypeOf obControl Is CheckedListBox Then
							If bSelectedValueIsUsed Then
								CType(obControl, CheckedListBox).SelectedValue = aString
							Else
								CType(obControl, CheckedListBox).Text = aString
							End If

					ElseIf TypeOf obControl Is CheckBox Then
							CType(obControl, CheckBox).Checked = CBool(aString)

					ElseIf TypeOf obControl Is RadioButton Then
							CType(obControl, RadioButton).Checked = CBool(aString)

					ElseIf TypeOf obControl Is DateTimePicker Then
							CType(obControl, DateTimePicker).Value = Convert.ToDateTime(aString)

					ElseIf TypeOf obControl Is MonthCalendar Then
							If aSelectionEndIsUsed Then
								CType(obControl, MonthCalendar).SelectionEnd = Convert.ToDateTime(aString)
							Else
								CType(obControl, MonthCalendar).SelectionStart = Convert.ToDateTime(aString)
							End If

					ElseIf TypeOf obControl Is TrackBar Then
							CType(obControl, TrackBar).Value = Convert.ToInt16(aString)

					ElseIf TypeOf obControl Is NumericUpDown Then
							CType(obControl, NumericUpDown).Value = Convert.ToInt16(aString)

					ElseIf TypeOf obControl Is DomainUpDown Then
							CType(obControl, DomainUpDown).Text = aString

					Else
						' do nothing and ignore other types of controls
					End If

			Catch e As Exception
				Throw e
			End Try

End Sub

		Protected Friend Sub SetControlValues()
		Dim bSelectionEndIsUsed As Boolean

			' Catch the special case for MonthCalendar control
			bSelectionEndIsUsed = (MyBase.WhereCondition = WhereConditions.Between Or _
														 MyBase.WhereCondition = WhereConditions.NotBetween) And _
														 MyBase.ValueType = TypeOfValues.DateType And _
														 obControl1 Is obControl2

					SetControlText(obControl1, MyBase.Value)
					SetControlText(obControl2, MyBase.Value2, bSelectionEndIsUsed)

		End Sub

	Protected Function GetControlByName(ByVal strControl As String) As Control
		Dim eControl As IEnumerator
		eControl = QuickWhere.ParentForm.Controls.GetEnumerator

		While eControl.MoveNext
			If CType(eControl.Current, Control).Name = strControl Then
				Return CType(eControl.Current, Control)
			End If
		End While

		Return Nothing

	End Function

#End Region

#Region " IWhereEntity implementation "

	Public Overrides Function GetSql() As String

		Call FillTheValues()
		Return MyBase.GetSql

	End Function

		Friend Sub FillTheValues()
			Dim bSelectionEndIsUsed As Boolean

						' Catch the special case for MonthCalendar control
						bSelectionEndIsUsed = _
													(MyBase.WhereCondition = WhereConditions.Between Or _
													MyBase.WhereCondition = WhereConditions.NotBetween) And _
													MyBase.ValueType = TypeOfValues.DateType And _
													obControl1 Is obControl2

							MyBase.Value = GetControlText(obControl1)
							MyBase.Value2 = GetControlText(obControl2, bSelectionEndIsUsed)

		End Sub

#End Region

	End Class

End Namespace

End Namespace