Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO

Module Utilities
    Public Function insertproduct(ByVal ProductName As String, ByVal categories As String, ByVal Quantity As String, ByVal Price As String, ByVal RetailedPrice As String, ByVal MarkUp As String, ByVal barcode As Byte()) As Boolean
        Try
            Dim query = "insert into Inventory ([categories], [ProductName],[Quantity],[Price],[RetailedPrice],[MarkUp]) values (@categories, @ProductName,@Quantity,@Price,@RetailedPrice,@MarkUp)"
            Connection.AddParam("categories", categories)
            Connection.AddParam("ProductName", ProductName)
            Connection.AddParam("Quantity", Quantity)
            Connection.AddParam("Price", Price)
            Connection.AddParam("RetailedPrice", RetailedPrice)
            Connection.AddParam("MarkUp", MarkUp)
            Connection.Parameters.Add(New SqlParameter("@br", barcode))

            If Connection.Insert(query) Then
                MsgBox("Added Succesfully!", MsgBoxStyle.Information, "Success!")
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error adding product: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function UpdateProduct(ByVal id As String, ByVal categories As String, ByVal ProductName As String, ByVal Quantity As String, ByVal Price As String, ByVal RetailedPrice As String, ByVal MarkUp As String, ByVal barcode As Byte()) As Boolean
        Try
            Dim query = "UPDATE Inventory SET ProductName=@ProductName, Quantity=@Quantity, Price=@Price, RetailedPrice=@RetailedPrice, MarkUp=@MarkUp WHERE id=@id"
            Connection.AddParam("@id", id)
            Connection.AddParam("@categories", categories)
            Connection.AddParam("@ProductName", ProductName)
            Connection.AddParam("@Quantity", Quantity)
            Connection.AddParam("@Price", Price)
            Connection.AddParam("RetailedPrice", RetailedPrice)
            Connection.AddParam("MarkUp", MarkUp)
            Connection.Parameters.Add(New SqlParameter("@br", barcode))
            If Connection.Update(query) Then
                MsgBox("Update Succesfully!", MsgBoxStyle.Information, "Success!")
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error updating student: " & ex.Message)
        End Try
        Return False
    End Function

    Public Function DeleteProduct(ByVal id As String) As Boolean
        Try
            Dim query = "DELETE FROM Inventory WHERE id=@id"
            Connection.AddParam("@id", id)

            If Connection.Delete(query) Then
                MsgBox("Delete Succesfully!", MsgBoxStyle.Information, "Success!")
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error deleting product: " & ex.Message)
        End Try
        Return False
    End Function

    Public Sub SaveBarcode(ByVal barcodeImg As Image)
        If barcodeImg Is Nothing Then Return

        Dim saveDialog As New SaveFileDialog()
        saveDialog.Filter = "bmp (*.bmp)|*.bmp|jpeg (*.jpeg)| *.jpeg|png (*.png)|*.png|tiff (*.tiff)|*.tiff"
        If saveDialog.ShowDialog() = DialogResult.OK Then
            barcodeImg.Save(saveDialog.FileName)
        End If
    End Sub
End Module
