Imports System.Threading.Tasks
Imports System.Text
Imports System.Linq
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System

Namespace SampleVBProject
    Class Program
        Shared Sub Main()
            TryCreateTable()
            Console.WriteLine("Simple DB Manipulation Program (VB)" & vbLf.ToUpper())
            While True
                DisplayMenu()
            End While
        End Sub

        ''' <summary>
        ''' This method attempts to create the SampleProducts SQL Table
        ''' Prints an error if table already exists.
        ''' </summary>
        Shared Sub TryCreateTable()
            Using conn As New SqlConnection(My.Settings.SampleProducts1ConnectionString)
                conn.Open()
                Try
                    Using command As New SqlCommand("CREATE TABLE SampleProducts1 (ItemNum INT, ItemName TEXT, ItemDesc TEXT)", conn)
                        command.ExecuteNonQuery()
                    End Using
                Catch
                    Console.WriteLine("Table found, new table not created.".ToUpper())
                End Try
            End Using
        End Sub

        ''' <summary>
        ''' Menu organization
        ''' </summary>
        Shared Sub DisplayMenu()
            Console.WriteLine("Enter desired operation: DISPLAY, DELETE, UPDATE, INSERT, EXIT")
            Console.Write(vbLf & " =>")
            Dim input As String() = Console.ReadLine().Split(","c)
            Dim switch_on As String = input(0).ToUpper()
            Select Case switch_on
                Case "DISPLAY"
                    ' Displays products currently in DB
                    Console.WriteLine("Current Product Entries:")
                    DisplayProducts()
                    Exit Select

                Case "DELETE"
                    ' Dialog to delete a record from DB
                    Console.WriteLine("Enter the item number to be deleted.")
                    Console.Write(vbLf & " =>")
                    Dim toDelete As String() = Console.ReadLine().Split(","c)
                    Dim deleteVal As Integer = 0
                    While Not Int32.TryParse(toDelete(0), deleteVal)
                        Console.WriteLine("Enter the item number to be deleted.")
                        Console.Write(vbLf & " =>")
                        toDelete = Console.ReadLine().Split(","c)
                    End While
                    Try
                        RemoveProduct(Convert.ToInt32(toDelete(0)))
                    Catch
                        Console.WriteLine("Could not Delete record from DB. Please verify Item Number and try again." & vbLf)
                    End Try
                    Exit Select

                Case "UPDATE"
                    ' Dialog to Update a current record in DB
                    Console.WriteLine("Enter the item number to be updated.")
                    Console.Write(vbLf & " =>")
                    Dim toUpdate As String() = Console.ReadLine().Split(","c)
                    Dim updateVal As Integer = 0
                    While Not Int32.TryParse(toUpdate(0), updateVal)
                        Console.WriteLine("Enter the item number to be updated.")
                        Console.Write(vbLf & " =>")
                        toUpdate = Console.ReadLine().Split(","c)
                    End While
                    Console.WriteLine("Enter the new item name and item description to be updated to." & vbLf + "Enter in the form: new name, new description")
                    Console.Write(vbLf & " =>")
                    Dim newData As String() = Console.ReadLine().Split(","c)
                    UpdateProduct(Integer.Parse(toUpdate(0)), newData(0), newData(1))
                    Exit Select

                Case "INSERT"
                    Console.WriteLine("Enter Data in Form:" & vbLf & "item num, item name, item description")
                    Console.Write(vbLf & " =>")
                    Dim insertData As String() = Console.ReadLine().Split(","c)
                    Try
                        Dim itemNum As Integer = Integer.Parse(insertData(0))
                        ' The item num.
                        Dim itemName As String = insertData(1)
                        ' The name string.
                        Dim itemDesc As String = insertData(2)
                        ' The item description string.
                        ' Add the data to the SQL database.
                        AddProduct(itemNum, itemName, itemDesc)
                    Catch
                        Console.WriteLine("Could not insert data into DB." & vbLf + "Please verify data is in correct format and try again." & vbLf)
                    End Try
                    Exit Select

                Case "EXIT"
                    Environment.[Exit](0)
                    Exit Select

                Case "FILL"
                    FillDB()
                    Exit Select
                Case Else
                    Console.WriteLine("Input not recognized as valid command. Please enter a different command." & vbLf)
                    Exit Select
            End Select
        End Sub

        ''' <summary>
        ''' Insert Product into database
        ''' </summary>
        ''' <param name="ItemNum">The Item number for the product.</param>
        ''' <param name="ItemName">The Product Name</param>
        ''' <param name="ItemDesc">The Item Description</param>
        Shared Sub AddProduct(ItemNum As Integer, ItemName As String, ItemDesc As String)
            Using conn As New SqlConnection(My.Settings.SampleProducts1ConnectionString)
                conn.Open()
                Try
                    Using command As New SqlCommand("INSERT INTO SampleProducts1 VALUES(@ItemNum, @ItemName, @ItemDesc)", conn)
                        command.Parameters.Add(New SqlParameter("ItemNum", ItemNum))
                        command.Parameters.Add(New SqlParameter("ItemName", ItemName))
                        command.Parameters.Add(New SqlParameter("ItemDesc", ItemDesc))
                        command.ExecuteNonQuery()
                    End Using
                Catch
                    Console.WriteLine("Could not insert product into DB.")
                End Try
            End Using
        End Sub

        ''' <summary>
        ''' Update Product in database
        ''' </summary>
        ''' <param name="ItemNum">The Item number for the product.</param>
        ''' <param name="newName">The Product Name</param>
        ''' <param name="newDesc">The Item Description</param>
        Shared Sub UpdateProduct(ItemNum As Integer, newName As String, newDesc As String)
            Using conn As New SqlConnection(My.Settings.SampleProducts1ConnectionString)
                conn.Open()
                Try
                    Using command As New SqlCommand("UPDATE SampleProducts1 SET ItemName=@ItemName, ItemDesc=@ItemDesc WHERE ItemNum=@ItemNum", conn)
                        command.Parameters.Add(New SqlParameter("ItemNum", ItemNum))
                        command.Parameters.Add(New SqlParameter("ItemName", newName))
                        command.Parameters.Add(New SqlParameter("ItemDesc", newDesc))
                        command.ExecuteNonQuery()
                    End Using
                Catch
                    Console.WriteLine("Could not update product into DB.")
                End Try
            End Using
        End Sub

        ''' <summary>
        ''' Remove Product from database
        ''' </summary>
        ''' <param name="Num">The Item number for the product.</param>
        Shared Sub RemoveProduct(Num As Integer)
            Using conn As New SqlConnection(My.Settings.SampleProducts1ConnectionString)
                conn.Open()
                Try
                    Using command As New SqlCommand("DELETE SampleProducts1 WHERE ItemNum=@ItemNum", conn)
                        command.Parameters.Add(New SqlParameter("@ItemNum", Num))
                        command.ExecuteNonQuery()
                    End Using
                Catch
                    Console.WriteLine("Could not Delete product from DB.")
                End Try
            End Using
        End Sub

        '''<summary>
        '''Build up database of products
        '''</summary>
        Shared Sub FillDB()
            Dim products As New List(Of Product)()
            products.Add(New Product(10, "Filet Mignon", "Super tender steak"))
            products.Add(New Product(11, "PorterHouse", "2 in 1 steak"))
            products.Add(New Product(12, "Top Sirloin", "Good quality Steak"))
            products.Add(New Product(13, "Twice Baked Potatoes", "Potatoes covered with cheese and baked twice"))
            products.Add(New Product(14, "Lobster Tails", "North Atlantic Lobster Tails"))

            For Each iter As Product In products
                AddProduct(iter.m_ItemNum, iter.m_ItemName, iter.m_ItemDesc)
            Next
        End Sub

        ''' <summary>
        ''' Read in all rows from SampleProducts1 table and store in a list.
        ''' </summary>
        Shared Sub DisplayProducts()
            Dim products As New List(Of Product)()
            Using conn As New SqlConnection(My.Settings.SampleProducts1ConnectionString)
                conn.Open()

                Using command As New SqlCommand("SELECT * FROM SampleProducts1", conn)
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim ItemNum As Integer = reader.GetInt32(0)
                        ' ItemNum int
                        Dim ItemName As String = reader.GetString(1)
                        ' ItemName string
                        Dim ItemDesc As String = reader.GetString(2)
                        ' ItemDesc string
                        products.Add(New Product(ItemNum, ItemName, ItemDesc))
                    End While
                End Using
            End Using
            products.Sort(Function(x, y) x.m_ItemNum.CompareTo(y.m_ItemNum))
            For Each iter As Product In products
                Console.WriteLine(iter)
            Next
            Console.WriteLine(vbLf)
        End Sub
    End Class

    Class Product
        Public Sub New(num As Integer, name As String, desc As String)
            m_ItemNum = num
            m_ItemName = name
            m_ItemDesc = desc
        End Sub
        Public Property m_ItemNum() As Integer
        Public Property m_ItemName() As String
        Public Property m_ItemDesc() As String
        Public Overrides Function ToString() As String
            Return String.Format("Item Number: {0}, Item Name: {1}, Item Description: {2}", m_ItemNum, m_ItemName, m_ItemDesc)
        End Function
    End Class
End Namespace