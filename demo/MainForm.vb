Option Explicit On 
Option Strict On

Public Class MainForm
		Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

		Public Sub New()
				MyBase.New()

				'This call is required by the Windows Form Designer.
				InitializeComponent()

				'Add any initialization after the InitializeComponent() call

		End Sub

		'Form overrides dispose to clean up the component list.
		Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
				If disposing Then
						If Not (components Is Nothing) Then
								components.Dispose()
						End If
				End If
				MyBase.Dispose(disposing)
		End Sub

		'Required by the Windows Form Designer
		Private components As System.ComponentModel.IContainer

		'NOTE: The following procedure is required by the Windows Form Designer
		'It can be modified using the Windows Form Designer.  
		'Do not modify it using the code editor.

	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
	Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
	Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
	Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
	Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
		<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container()
Me.Label1 = New System.Windows.Forms.Label()
Me.Label2 = New System.Windows.Forms.Label()
Me.LinkLabel4 = New System.Windows.Forms.LinkLabel()
Me.LinkLabel3 = New System.Windows.Forms.LinkLabel()
Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
Me.SuspendLayout()
'
'Label1
'
Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label1.Location = New System.Drawing.Point(16, 16)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(388, 48)
Me.Label1.TabIndex = 0
Me.Label1.Text = "QuickWhere is component that helps you in the creation of complex forms for searc" & _
"hing databases."
Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'Label2
'
Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Label2.Location = New System.Drawing.Point(16, 84)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(388, 20)
Me.Label2.TabIndex = 0
Me.Label2.Text = "Use the links below to view the examples of how to use QuickWhere."
Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'LinkLabel4
'
Me.LinkLabel4.Location = New System.Drawing.Point(331, 128)
Me.LinkLabel4.Name = "LinkLabel4"
Me.LinkLabel4.Size = New System.Drawing.Size(72, 16)
Me.LinkLabel4.TabIndex = 2
Me.LinkLabel4.TabStop = True
Me.LinkLabel4.Text = "Example 4"
Me.LinkLabel4.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.ToolTip1.SetToolTip(Me.LinkLabel4, "This example demonstrates how the general setting of QuickWhere reflect the gener" & _
"ated WHERE clause." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the controls show the field name, the value " & _
"type, the operation and more information.")
'
'LinkLabel3
'
Me.LinkLabel3.Location = New System.Drawing.Point(226, 128)
Me.LinkLabel3.Name = "LinkLabel3"
Me.LinkLabel3.Size = New System.Drawing.Size(72, 16)
Me.LinkLabel3.TabIndex = 2
Me.LinkLabel3.TabStop = True
Me.LinkLabel3.Text = "Example 3"
Me.LinkLabel3.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.ToolTip1.SetToolTip(Me.LinkLabel3, "This example demonstrates:" & Microsoft.VisualBasic.ChrW(10) & " - the use of IN(...) clause generated from ListBox or" & _
" CheckedListBox" & Microsoft.VisualBasic.ChrW(10) & " - how to use mixed ANDs and ORs in two alternative ways" & Microsoft.VisualBasic.ChrW(10) & " - how " & _
"to use operation over values" & Microsoft.VisualBasic.ChrW(10) & " - how to use IS (NOT) NULL clause" & Microsoft.VisualBasic.ChrW(10) & " - the applicati" & _
"on of Used property" & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the controls show the field name, the value" & _
" type, the operation and more information.")
'
'LinkLabel2
'
Me.LinkLabel2.Location = New System.Drawing.Point(121, 128)
Me.LinkLabel2.Name = "LinkLabel2"
Me.LinkLabel2.Size = New System.Drawing.Size(72, 16)
Me.LinkLabel2.TabIndex = 2
Me.LinkLabel2.TabStop = True
Me.LinkLabel2.Text = "Example 2"
Me.LinkLabel2.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.ToolTip1.SetToolTip(Me.LinkLabel2, "This example demonstrates how to:" & Microsoft.VisualBasic.ChrW(10) & "- attach various types of Win controls to Quick" & _
"Where that represent different types of values." & Microsoft.VisualBasic.ChrW(10) & "- to save the user input to a fi" & _
"le as filter for further use" & Microsoft.VisualBasic.ChrW(10) & "- to load a file as filter" & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "The tooltips of the att" & _
"ached controls show the field name, the value type, the operation and more infor" & _
"mation.")
'
'LinkLabel1
'
Me.LinkLabel1.Location = New System.Drawing.Point(16, 128)
Me.LinkLabel1.Name = "LinkLabel1"
Me.LinkLabel1.Size = New System.Drawing.Size(72, 16)
Me.LinkLabel1.TabIndex = 2
Me.LinkLabel1.TabStop = True
Me.LinkLabel1.Text = "Example 1"
Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.ToolTip1.SetToolTip(Me.LinkLabel1, "This is very simple example. It demonstrates the basic use of QuickWhere.")
'
'MainForm
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(424, 161)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.LinkLabel4, Me.Label1, Me.Label2, Me.LinkLabel3, Me.LinkLabel2, Me.LinkLabel1})
Me.Name = "MainForm"
Me.Text = "QuickWhere Demo"
Me.ResumeLayout(False)

		End Sub

#End Region

	Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
		Dim frmForm1 As New Form1()
		frmForm1.ShowDialog()
	End Sub

	Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
		Dim frmForm2 As New Form2()
		frmForm2.ShowDialog()
	End Sub

	Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
		Dim frmForm3 As New Form3()
		frmForm3.ShowDialog()
	End Sub

	Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
		Dim frmForm4 As New Form4()
		frmForm4.ShowDialog()
	End Sub
End Class
