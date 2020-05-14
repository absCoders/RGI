Public Class WBCCHKB1
        ' Implements the CreationFilter interface
        Implements IUIElementCreationFilter

        Public ColumnNames As List(Of String) = New List(Of String)

        ' This event will fire when the CheckBox is clicked. 
        Public Event HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As HeaderCheckBoxEventArgs)

        Public Sub AfterCreateChildElements(ByVal parent As Infragistics.Win.UIElement) Implements Infragistics.Win.IUIElementCreationFilter.AfterCreateChildElements
            ' Check for the HeaderUIElement
            If TypeOf parent Is UltraWinGrid.HeaderUIElement Then
                ' Get the actual ColumnHeader that the HeaderUIElement is attached to
                Dim aColHeader As Infragistics.Win.UltraWinGrid.ColumnHeader
                aColHeader = CType(parent, UltraWinGrid.HeaderUIElement).Header

                ' Only put the Checkbox in the header of the ComboBox, AllowEdit
                If aColHeader.Column.Style = UltraWinGrid.ColumnStyle.CheckBox _
                        AndAlso (ColumnNames.Count = 0 OrElse ColumnNames.Contains(aColHeader.Column.Key)) _
                        AndAlso aColHeader.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                    Dim aTextUIElement As TextUIElement
                    Dim aCheckBoxUIElement As CheckBoxUIElement

                    ' Since the grid sometimes re-uses UIElements, we need to check to make sure 
                    ' the header does not already have a CheckBoxUIElement attached to it.
                    ' If it does, we just get a reference to the existing CheckBoxUIElement,
                    ' and reset its properties.
                    aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                    If aCheckBoxUIElement Is Nothing Then
                        ' Create a New CheckBoxUIElement
                        aCheckBoxUIElement = New CheckBoxUIElement(parent)
                    End If

                    ' Get the TextUIElement - this is where the text for the 
                    ' Header is displayed. We need this so we can push it to the right
                    ' in order to make room for the CheckBox
                    aTextUIElement = CType(parent.GetDescendant(GetType(TextUIElement)), TextUIElement)

                    ' Sanity check
                    If aTextUIElement Is Nothing Then Exit Sub

                    ' Get the Header and see if the Tag has been set. I the Tag is 
                    ' set, we will assume it's the stored CheckState. This has to be
                    ' done in order to maintain the CheckState when the grid repaints and
                    ' UIElement are destroyed and recreated. 
                    Dim aHeader As Infragistics.Win.UltraWinGrid.ColumnHeader = CType(aCheckBoxUIElement.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)).GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader)), Infragistics.Win.UltraWinGrid.ColumnHeader)

                    If aHeader.Tag Is Nothing Then
                        ' If the tag was nothing, this is probably the first time this 
                        ' HeaderRow is being displayed, so default to Unchecked
                        aHeader.Tag = CheckState.Unchecked
                    Else
                        aCheckBoxUIElement.CheckState = CType(aHeader.Tag, CheckState)
                    End If

                    ' Hook the ElementClick of the CheckBoxUIElement
                    AddHandler aCheckBoxUIElement.ElementClick, AddressOf aCheckBoxUIElement_ElementClick

                    ' Add the CheckBoxUIElement to the HeaderUIElement
                    parent.ChildElements.Add(aCheckBoxUIElement)

                    ' Position the CheckBoxUIElement. The number 3 here is used for 3
                    ' pixels of padding between the CheckBox and the side of the header
                    ' The CheckBox is shifted down slightly so it is centered in the header
                    aCheckBoxUIElement.Rect = New Drawing.Rectangle(parent.Rect.X + 3, parent.Rect.Y + ((parent.Rect.Height - aCheckBoxUIElement.CheckSize.Height) / 2), aCheckBoxUIElement.CheckSize.Width, aCheckBoxUIElement.CheckSize.Height)

                    ' Push the TextUIElement to the right a little to make 
                    ' room for the CheckBox. 3 pixels of padding are used again. 
                    aTextUIElement.Rect = New Drawing.Rectangle(aCheckBoxUIElement.Rect.Right + 3, aTextUIElement.Rect.Y, parent.Rect.Width - (aCheckBoxUIElement.Rect.Right - parent.Rect.X), aTextUIElement.Rect.Height)
                Else
                    ' If the column is not a boolean column, we do not want to have a checkbox in it
                    ' Since UIElements can be reused by the grid, there is a chance that one of the
                    ' HeaderUIElements that we added a checkbox to for a boolean column header
                    ' will be reused in a column that is not boolean.  In this case, we must remove
                    ' the checkbox so that it will not appear in an inappropriate column header.
                    Dim aCheckBoxUIElement As CheckBoxUIElement
                    aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                    If Not aCheckBoxUIElement Is Nothing Then
                        parent.ChildElements.Remove(aCheckBoxUIElement)
                        aCheckBoxUIElement.Dispose()
                    End If
                End If
            End If
        End Sub

        Public Function BeforeCreateChildElements(ByVal parent As Infragistics.Win.UIElement) As Boolean Implements Infragistics.Win.IUIElementCreationFilter.BeforeCreateChildElements
            ' Don't need to do anything here.
            Return False
        End Function

        Private Sub aCheckBoxUIElement_ElementClick(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)
            ' Get the CheckBoxUIElement that was clicked
            Dim aCheckBoxUIElement As CheckBoxUIElement = CType(e.Element, CheckBoxUIElement)

            ' Get the Header associated with this particular element
            Dim aHeaderUIElement As UltraWinGrid.HeaderUIElement = CType(aCheckBoxUIElement.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)), UltraWinGrid.HeaderUIElement)
            Dim aHeader As Infragistics.Win.UltraWinGrid.ColumnHeader = CType(aHeaderUIElement.GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader)), Infragistics.Win.UltraWinGrid.ColumnHeader)

            ' Set the Tag on the Header to the new CheckState
            aHeader.Tag = aCheckBoxUIElement.CheckState

            ' So that we can apply various changes only to the relevant Rows collection that the header belongs to
            Dim hRows As UltraWinGrid.RowsCollection = CType(aHeaderUIElement.GetContext(GetType(UltraWinGrid.RowsCollection)), UltraWinGrid.RowsCollection)

            ' Raise an event so the programmer can do something when the CheckState changes
            RaiseEvent HeaderCheckBoxClicked(Me, New HeaderCheckBoxEventArgs(aHeader, aCheckBoxUIElement.CheckState, hRows))
        End Sub

        ' EventArgs used for the HeaderCheckBoxClicked event. This event has to pass in the CheckState and the Header
        Public Class HeaderCheckBoxEventArgs
            Inherits EventArgs

            Public Sub New(ByVal Header As Infragistics.Win.UltraWinGrid.ColumnHeader, ByVal CheckState As CheckState, ByRef Rows As UltraWinGrid.RowsCollection)
                mvarHeader = Header
                mvarCheckState = CheckState
                mvarRowsCollection = Rows
            End Sub

            Private mvarRowsCollection As UltraWinGrid.RowsCollection
            Private mvarHeader As Infragistics.Win.UltraWinGrid.ColumnHeader
            Private mvarCheckState As CheckState

            ' Expose the rows collection for the specific row island that the header belongs to
            Public ReadOnly Property Rows() As UltraWinGrid.RowsCollection
                Get
                    Return mvarRowsCollection
                End Get
            End Property

            Public ReadOnly Property Header() As Infragistics.Win.UltraWinGrid.ColumnHeader
                Get
                    Return mvarHeader
                End Get
            End Property

            Public Property CheckState() As CheckState
                Get
                    Return mvarCheckState
                End Get
                Set(ByVal Value As CheckState)
                    mvarCheckState = Value
                End Set
            End Property
        End Class
    End Class
