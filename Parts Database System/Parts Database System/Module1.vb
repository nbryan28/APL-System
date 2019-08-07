Imports MySql.Data.MySqlClient

Module Module1

    Public secret_word As String
    Public version_c As String  'store current version of program
    Public ok_press As Boolean  'Login form use this to verified if OK button has been pressed
    Public Log_out As Boolean  'boolean varianle that indicates if someone has logged in/out
    Public current_user As String 'store the log in user
    Public enable_mess As Boolean 'enable APL and email messaging

    'Roles boolean
    Public Engineer As Boolean
    Public Engineer_management As Boolean
    Public Procurement As Boolean
    Public Procurement_management As Boolean
    Public Inventory_m As Boolean
    Public Manufacturing As Boolean
    Public General_management As Boolean


    'Version
    Public myVersion As String
    Public mydatagrid As Integer ' this control the search form

    'Public connection As New MySqlConnection("datasource=localhost;port=3306;username=root;password=atronixatl;database=jobs")


End Module


