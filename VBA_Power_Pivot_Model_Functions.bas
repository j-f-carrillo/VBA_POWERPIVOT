Attribute VB_Name = "VBA_Power_Pivot_Model_Functions"
Sub add_model_measures()
    Dim myTable As ListObject
    Dim myArray As Variant
    Dim x   As Long
    Dim mdltbl As ModelTable
'    Set mdltbl = get_Model_Table(ActiveSheet.Range("A2").Value)
    Dim existing_measures As Variant
    existing_measures = model_measures_arr()
    
    
    'Set path for Table variable
    Set myTable = ActiveSheet.ListObjects(1)
    
    
    'Create Array List from Table
    myArray = myTable.DataBodyRange
    On Error Resume Next
    Dim maxRows As Integer
    maxRows = UBound(myArray)
    
    Dim form As String, mname As String
    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
    For x = LBound(myArray) To UBound(myArray)
         Set mdltbl = get_Model_Table(myArray(x, 1))
        
        form = myArray(x, 3)
        mname = myArray(x, 2)
        
        If IsInArray(myArray(x, 2), existing_measures) <> True Then
            
            
            With ActiveWorkbook.Model
                .ModelMeasures.Add mname, mdltbl, form, FormatInformation:=ActiveWorkbook.Model.ModelFormatDecimalNumber
            End With
            Debug.Print x & " of " & maxRows & " - " & mname & vbTab & " Added To Model"
           
            
            
        Else: Debug.Print x & " of " & maxRows & " - " & mname & vbTab & " Already Exist Moving to Next"
       
        End If
        
    Next x
    Debug.Print "Finished"
End Sub
'Basic function to add 1 Meausre
Function Add_Single_ModelMeasure(mTable As ModelTable, mname As String, mformula As String, Optional ByVal mformat As String)


    With ActiveWorkbook.Model
        
        .ModelMeasures.Add mname, mTable, mformula, .ModelFormatDecimalNumber
        
    End With
    
End Function
Function get_Model_Table(str As Variant) As ModelTable
'Find Model Table and Selects
    Dim mdl As Model
    Set mdl = ActiveWorkbook.Model
    For Each x In mdl.ModelTables
        
        
        If str = x.Name Then
            Set get_Model_Table = x
            'Debug.Print "Selected " & x.Name
            
            Exit Function
        End If
        
    Next x
    
    
End Function

Function model_measures_arr() As Variant
'Create and Array of Currently Present Meaures
    Dim mdl As Model
    Set mdl = ActiveWorkbook.Model
    Dim arr As Variant
    Dim i   As Integer
    i = 0
    ReDim arr(mdl.ModelMeasures.Count - 1)
    
    For Each x In mdl.ModelMeasures
        arr(i) = x.Name
        
        i = i + 1
    Next x
    model_measures_arr = arr
    
End Function
Sub is_Table_in_Model()
'Search Model for Table

    Dim myTable As ListObject
    Dim myArray As Variant
    Dim x   As Long
    Dim tablearr As Variant
    tablearr = model_tables_name_arr()
    
    
    'Set path for Table variable
    Set myTable = ActiveSheet.ListObjects(1)
    
    'Create Array List from Table
    myArray = myTable.DataBodyRange
    
    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
    For x = LBound(myArray) To UBound(myArray)
        Debug.Print myArray(x, 2) & " " & myArray(x, 4)
        
        If IsInArray(myArray(x, 2), tablearr) = True Then
            If IsInArray(myArray(x, 4), tablearr) = True Then
                
                variant_add_Power_Pivot_Relationship myArray(x, 2), myArray(x, 3), myArray(x, 4), myArray(x, 5)
                
            Else: Debug.Print "failed" & myArray(x, 4)
            End If
        Else: Debug.Print "failed" & myArray(x, 2)
        End If
        
        
    Next x
End Sub
Sub AddNewRelationship()

    'Declare Variables

    Dim myModel As Model
    Dim ModelRelt As ModelRelationship
    Dim ModelRelts As ModelRelationships

    'Create a reference to the data model in our workbook
    Set myModel = ActiveWorkbook.Model
    'myModel.AddConnection ConnectionToDataSource:=WrkBookConn
    
    'Create Variables to house Model Tables.
    Dim ModelTbl1 As ModelTable
    Dim ModelTbl2 As ModelTable
    
    'Get the necessary model tables.
    Set ModelTbl1 = myModel.ModelTables.Item("raw_data")
    Set ModelTbl2 = myModel.ModelTables.Item("aa_weekend")
    
    'Create variables to get model table columns.
    Dim PrimCol As ModelTableColumn
    Dim ForgCol As ModelTableColumn
    
    'Get the necessary model table columns.
    Set ForgCol = ModelTbl1.ModelTableColumns.Item("week_end")
    Set PrimCol = ModelTbl2.ModelTableColumns.Item("Weekend Date")
    
    'Add a new relationship
    myModel.ModelRelationships.Add ForeignKeyColumn:=ForgCol, PrimaryKeyColumn:=PrimCol
    
    'Create a reference to the relationships in your model.
    Set ModelRelts = myModel.ModelRelationships
    
    'Print if they're active or not. TRUE means active.
    For Each ModelRelt In ModelRelts
        Debug.Print ModelRelt.Active
    Next
    
End Sub
Sub ModelTableObject()

    'Declare our variables
    Dim myModel As Model
    Dim ModelTbls As ModelTables
    Dim ModelTbl As ModelTable
    Dim ModelCols As ModelTableColumns
    Dim ModelCol As ModelTableColumn
    Dim ModelConn As ModelConnection

    'Create a reference to our power pivot model
    Set myModel = ActiveWorkbook.Model
    
    'Reference the model tables collection
    Set ModelTbls = myModel.ModelTables
    
    'Count the number of tables
    Debug.Print ModelTbls.Count
    
    'Get the parent object name
    Debug.Print ModelTbls.Parent.Name
    
    'Loop through each table
    For Each ModelTbl In ModelTbls
        Debug.Print ModelTbl.Name
        Debug.Print ModelTbl.RecordCount
        Debug.Print ModelTbl.SourceName
        Debug.Print ModelTbl.SourceWorkbookConnection
    Next
    
    Set ModelTbl = ModelTbls.Item("Price_Data")
    Set ModelCols = ModelTbl.ModelTableColumns
    
    'Looping through each col
    For Each ModelCol In ModelCols
        Debug.Print ModelCol.Name
        Debug.Print ModelCol.DataType
        Debug.Print ModelCol.Application
    Next
    
    'Lets work with the model connection object now.
    Set ModelConn = myModel.DataModelConnection.ModelConnection
    Debug.Print ModelConn.CommandType
    Debug.Print ModelConn.CommandText
    Debug.Print ModelConn.ADOConnection
    
End Sub
Function add_Power_Pivot_Relationship(Foreign_Key_Table As String, _
         Foreign_Key_Column As String, _
         Primary_Key_Table As String, _
         Primary_Key_column As String)

    Dim myModel As Model
    Dim ModelRelt As ModelRelationship
    Dim ModelRelts As ModelRelationships

    'Create a reference to the data model in our workbook
    Set myModel = ActiveWorkbook.Model
    'myModel.AddConnection ConnectionToDataSource:=WrkBookConn
    
    'Create Variables to house Model Tables.
    Dim ModelTbl1 As ModelTable
    Dim ModelTbl2 As ModelTable
    
    'Get the necessary model tables.
    Set ModelTbl1 = myModel.ModelTables.Item(Foreign_Key_Table)
    Set ModelTbl2 = myModel.ModelTables.Item(Primary_Key_Table)
    
    'Create variables to get model table columns.
    Dim PrimCol As ModelTableColumn
    Dim ForgCol As ModelTableColumn
    
    'Get the necessary model table columns.
    Set ForgCol = ModelTbl1.ModelTableColumns.Item(Foreign_Key_Column)
    Set PrimCol = ModelTbl2.ModelTableColumns.Item(Primary_Key_column)
    
    'Add a new relationship
    myModel.ModelRelationships.Add ForeignKeyColumn:=ForgCol, PrimaryKeyColumn:=PrimCol
    
    'Create a reference to the relationships in your model.
    Set ModelRelts = myModel.ModelRelationships
    
    'Print if they're active or not. TRUE means active.
    For Each ModelRelt In ModelRelts
        Debug.Print ModelRelt.Active
    Next
End Function


Function model_tables_name_arr() As Variant
    Dim mdl As Model
    Set mdl = ActiveWorkbook.Model
    Dim arr As Variant, i As Integer
    i = 0
    ReDim arr(mdl.ModelTables.Count - 1)
    For Each x In mdl.ModelTables
        
        arr(i) = x.Name
        i = i + 1
        
        
    Next x
    
    model_tables_name_arr = arr
    
End Function
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean

    Dim element As Variant
    On Error GoTo IsInArrayError:                 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
