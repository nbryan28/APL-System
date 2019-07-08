Imports MySql.Data.MySqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class myQuote

    '---- initialization variables


    Public donotcompute As Boolean
    Public totales As Double
    Public labor As Double
    Public materials As Double
    Public expenses As Double
    Public subcontract As Double
    Public my_Set_info As New Dictionary(Of String, String)
    Public type_list_panel
    Public type_list_plc
    Public type_list_field
    Public type_list_control
    Public start_recal As Boolean
    Public datatable As DataTable  'use for solutions
    Public my_alloc_table As DataTable ' contains part, vendor pretty much a BOM or PR
    Public scapegoat As DataTable  'contains parts that will be added to mysql dummy table
    Public current_job As String
    Public RIO_lis 'table that stores IO Modules
    Public RIO_press As Boolean

    Dim my_assemblies
    Dim my_kits

    Public counter As Integer

    ' Cables
    Public M12_cables As New Dictionary(Of String, Double)
    Public M12_ES_cables As New Dictionary(Of String, Double)
    Public avoid_change_m12 As Boolean

    'dimensions
    Public Total_dim As Double
    Public Total_amps As Double
    Public panel_qty As Double
    Public panel_qty_30 As Double
    Public dropbo As Integer '1 = 24' panel, 2 = 30' panel, 3 = custom
    Public switch_b As Boolean
    Public my_qtys_panels As New Dictionary(Of String, Double)  'store modified qty for 24 per set
    Public my_qtys_panels_30 As New Dictionary(Of String, Double) 'store modified qty for 30 per set
    Public real_panels As New Dictionary(Of String, Double)  'store qty needed for 24 per set
    Public real_panels_30 As New Dictionary(Of String, Double) 'store qty needed for 30 per set
    Public pick_panel As New Dictionary(Of String, String)
    Public dimen_table As DataTable 'store info from panel_count form

    '-- IO
    Public inputs_io As Double
    Public outputs_io As Double
    Public motion_io As Double

    Public IO_custom As Boolean
    Public do_not_IO_recal As Boolean 'this trigger stop the recal of rio2grid 

    'enable general calculation
    Public go_on As Boolean
    Public change_sol As Boolean

    'Bulk and Labor cost variabel
    Public labor_t As Double
    Public bulk_t As Double

    '- calculate installation 
    Public cal_ins As Boolean

    'FLA table store data
    Public fla_data As DataTable


    '-- datatable feature_codes
    Public table_Feature_code As DataTable

    '--- databale feature parts
    Public table_feature_parts As DataTable

    '----- datatable call_feature
    Public table_call_feature As DataTable

    '-------- datatable f_dimensions
    Public table_f_dimensions As DataTable

    '--------datatable remote_io
    Public table_remote_io As DataTable

    '----------- datatable io_points
    Public table_io_points As DataTable

    '---------- datatable parts
    Public table_parts As DataTable
    '------------ datatable vendors

    Public table_vendors As DataTable

    '------------------- datatable assemblies
    Public table_assem As DataTable

    '-------------- datatable adv
    Public table_adv As DataTable

    Private Sub myQuote_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TabControl1.TabPages.Remove(TabPage8)  'remove cables estop
        RIO_press = False 'RIO warning
        cal_ins = True

        '---- initialize IO --

        inputs_io = 0
        outputs_io = 0
        motion_io = 0
        IO_custom = False
        do_not_IO_recal = False

        '------------------------------------------------------


        '///////////// LOAD TOTAL SPREADSHEET QUOTE ////////////////////////////////////////////////////

        For i = 5 To 6
            Quote_grid.Columns(i).HeaderCell.Style.BackColor = Color.DarkSalmon
        Next
        For i = 8 To 9
            Quote_grid.Columns(i).HeaderCell.Style.BackColor = Color.PowderBlue
        Next

        Quote_grid.Rows.Add(New String() {"", "ADA Panels and options"})
        Quote_grid.Rows(0).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(0).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(0).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(0).ReadOnly = True


        Quote_grid.Rows.Add(New String() {"", "Starter Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Control Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "PLC Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Bulk - Material Cost", "", "", "1", "INFO HERE", "", "20", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "Labor Cost", "", "", "1", "INFO HERE", "", "20", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "", "", "", "", "", "", "", ""})

        For i = 1 To 6
            Quote_grid.Rows(i).Cells(1).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {"", "Project Specific Items below"})
        Quote_grid.Rows(7).DefaultCellStyle.BackColor = Color.SlateGray
        Quote_grid.Rows(7).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(7).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})

        For i = 8 To 15
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next


        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(16).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(16).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(16).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(17).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(17).ReadOnly = True


        Quote_grid.Rows.Add(New String() {"", "Field Items"})
        Quote_grid.Rows(18).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(18).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(18).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(18).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Field Parts – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "1", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "M12 Cables – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "M12 Estop Cables – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})

        'this just to be camera and m12 labor
        Quote_grid.Rows(23).ReadOnly = True
        Quote_grid.Rows(24).ReadOnly = True
        Quote_grid.Rows(25).ReadOnly = True
        Quote_grid.Rows(26).ReadOnly = True

        For i = 19 To 26
            Quote_grid.Rows(i).Cells(1).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {"", "Project Specific Items below"})
        Quote_grid.Rows(27).DefaultCellStyle.BackColor = Color.SlateGray
        Quote_grid.Rows(27).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(27).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "ZOE HMI PC (Linux)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "ZOE HMI Enclosure (Global)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanners (Top)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanners Side (Single)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Scanner Frame (Top)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Scanner Frame (Side)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanner Tech (#Days)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Subcontract"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})

        For i = 28 To 42
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(43).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(43).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(43).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(44).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(44).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "In-Office Labor"})
        Quote_grid.Rows(45).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(45).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(45).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(45).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Engineering", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Layouts", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Panel Sets", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "PLC Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "HMI Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Services Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Reports Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "FAT", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Documentation", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Project Management", "", "", "0", "$62", "", "", "$95", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Project Labor"})

        For i = 46 To 56
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 46 To 56
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(57).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(57).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(57).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(58).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(58).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Onsite Labor"})
        Quote_grid.Rows(59).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(59).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(59).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(59).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "I/O Checkout", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "PLC Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "HMI Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "WCS Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Standby", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "SAT", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Training", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Startup Labor"})

        For i = 60 To 68
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 60 To 68
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(69).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(69).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(69).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(70).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(70).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Start Up Expenses And Travel"})
        Quote_grid.Rows(71).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(71).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(71).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(71).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Travel Time (TT)", "0", "", "0", "$62", "", "", "$75", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Mileage", "0", "", "0", "$0.57", "", "", "$0.57", "", "", "Startup Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Start-Up Expenses (SE)", "0", "", "0", "$250", "", "", "$250", "", "", "Startup Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Airfare (SE)", "0", "", "0", "$750", "", "", "$750", "", "", "Startup Expenses"})

        For i = 72 To 75
            Quote_grid.Rows(i).Cells(2).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 72 To 75
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(76).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(76).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(76).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(77).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(77).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "1", "include?"})
        Quote_grid.Rows(78).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(78).Cells(2).Style.BackColor = Color.White
        Quote_grid.Rows(78).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(78).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(78).ReadOnly = True
        Quote_grid.Rows(78).Cells(2).ReadOnly = False

        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Labor", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Labor"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Materials", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Materials"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Expenses", "1", "", "1", "INFO HERE", "", "0%", "", "", "", "Install Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Subcontract", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Subcontract"})


        For i = 79 To 82
            Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.White
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(83).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(83).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(83).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(83).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(84).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(84).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Est. Shipping", "1", "include?"})
        Quote_grid.Rows(85).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(85).Cells(2).Style.BackColor = Color.White
        Quote_grid.Rows(85).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(85).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(85).ReadOnly = True
        Quote_grid.Rows(85).Cells(2).ReadOnly = False

        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING I/O", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING Motor Big", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING – Scanners", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})

        For i = 86 To 88
            Quote_grid.Rows(i).Cells(2).ReadOnly = False
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.White
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(89).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(89).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(89).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(89).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(90).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(90).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(91).DefaultCellStyle.BackColor = Color.DarkSlateGray
        Quote_grid.Rows(91).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(92).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(92).ReadOnly = True

        '-------------- TOTAL SECTION --------------

        Quote_grid.Rows.Add(New String() {"", "TOTALS – By Component"})
        Quote_grid.Rows(93).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(93).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(93).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(93).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "ADA Panels", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Field Items", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "In-Office Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Onsite Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Start Up Expenses And Travel", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Shipping", "", "", "", "", "", "", "", "", ""})

        For i = 94 To 100
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(101).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(101).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(101).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(101).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(101).ReadOnly = True


        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(102).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(102).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – Sum all but Installation"})
        Quote_grid.Rows(103).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(103).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(103).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(103).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Hardware and Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Shipping", "", "", "", "", "", "", "", "", ""})

        For i = 104 To 106
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(107).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(107).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(107).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(107).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(107).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(108).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(108).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – By Type"})
        Quote_grid.Rows(109).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(109).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(109).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(109).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Project Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Project Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Project Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Manufacturing Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Manufacturing Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Subcontract", "", "", "", "", "", "", "", "", ""})

        For i = 110 To 122
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(123).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(123).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(123).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(123).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(123).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(124).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(124).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – Summary"})
        Quote_grid.Rows(125).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(125).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(125).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(125).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Subcontract", "", "", "", "", "", "", "", "", ""})

        For i = 126 To 129
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(130).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(130).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(130).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(130).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(130).ReadOnly = True


        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(131).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(131).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""}) : Quote_grid.Rows.Add(New String() {""}) : Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "EXPENSE CALCULATIONS"}) : Quote_grid.Rows(135).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {"", "Config"}) : Quote_grid.Rows(136).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "hours/ day onsite", "10"}) : Quote_grid.Rows(138).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(138).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "days onsite / trip", "12"}) : Quote_grid.Rows(139).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(139).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Travel time per trip", "12"}) : Quote_grid.Rows(140).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(140).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Miles per trip", "70"}) : Quote_grid.Rows(141).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(141).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Extra days of expenses per trip (travel days)", "0"}) : Quote_grid.Rows(142).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(142).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {""})


        Quote_grid.Rows.Add(New String() {"", "Output"}) : Quote_grid.Rows(144).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "DAYS ONSITE", ""}) : Quote_grid.Rows(146).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(146).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(146).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "# TRIPS", ""}) : Quote_grid.Rows(147).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(147).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(147).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Travel Time (TT)", ""}) : Quote_grid.Rows(148).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(148).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(148).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Mileage", ""}) : Quote_grid.Rows(149).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(149).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(149).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Start-Up Expenses (SE)", ""}) : Quote_grid.Rows(150).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(150).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(150).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Airfare (SE)", ""}) : Quote_grid.Rows(151).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(151).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(151).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})

        M12_cables.Add("V15-G-2M-PUR-V15-G", 0)
        M12_cables.Add("V15-G-3M-PUR-V15-G", 0)
        M12_cables.Add("V15-G-5M-PUR-V15-G", 0)
        M12_cables.Add("V15-G-10M-PUR-V15-G", 0)
        M12_cables.Add("V15-G-15M-PUR-V15-G", 0)
        M12_cables.Add("V15-G-20M-PUR-V15-G", 0)

        M12_ES_cables.Add("V15-G-S-YE2M-PUR-A-V15-G", 0)
        M12_ES_cables.Add("7000-40041-0150300", 0)
        M12_ES_cables.Add("V15-G-S-YE5M-PUR-A-V15-G", 0)
        M12_ES_cables.Add("V15-G-S-YE10M-PUR-A-V15-G", 0)
        M12_ES_cables.Add("V15-G-S-YE15M-PUR-A-V15-G", 0)
        M12_ES_cables.Add("V15-G-S-YE20M-PUR-A-V15-G", 0)


        '---------------- EStop cables table ---
        cables_grid.Rows.Add(New String() {})
        cables_grid.Rows.Add(New String() {})
        cables_grid.Rows.Add(New String() {})
        cables_grid.Rows.Add(New String() {})
        cables_grid.Rows.Add(New String() {})

        cables_grid.Rows(0).Cells(0).Value = 20
        cables_grid.Rows(0).Cells(1).Value = 0.35
 

        cables_grid.Rows(1).Cells(0).Value = 15
        cables_grid.Rows(1).Cells(1).Value = 0.2
        cables_grid.Rows(1).Cells(2).Value = 4
        cables_grid.Rows(1).Cells(3).Value = 60

        cables_grid.Rows(2).Cells(0).Value = 10
        cables_grid.Rows(2).Cells(1).Value = 0.15
        cables_grid.Rows(2).Cells(2).Value = 5
        cables_grid.Rows(2).Cells(3).Value = 50

        cables_grid.Rows(3).Cells(0).Value = 5
        cables_grid.Rows(3).Cells(1).Value = 0.2
        cables_grid.Rows(3).Cells(2).Value = 12
        cables_grid.Rows(3).Cells(3).Value = 60

        cables_grid.Rows(4).Cells(0).Value = 2
        cables_grid.Rows(4).Cells(1).Value = 0.1
        cables_grid.Rows(4).Cells(2).Value = 15
        cables_grid.Rows(4).Cells(3).Value = 30


        donotcompute = True

        Call compute()

        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        '----------------/////////////////////  LOAD INSTALLATION SPREADSHEET  //////////////////////--------

        Install_grid.Rows.Add(New String() {"INSTALLATION"})
        Install_grid.Rows(0).DefaultCellStyle.BackColor = Color.Gray
        Install_grid.Rows(0).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Install_grid.Rows(0).ReadOnly = True

        For i = 4 To 12
            Install_grid.Columns(i).HeaderCell.Style.BackColor = Color.DarkSalmon
        Next

        Install_grid.Rows.Add(New String() {"PE"}) : Install_grid.Rows.Add(New String() {"Pull Cord Stations"}) : Install_grid.Rows.Add(New String() {"Pull Cord Cable/eyelit (per foot)"})
        Install_grid.Rows.Add(New String() {"Pull Cord Pulley Kits"}) : Install_grid.Rows.Add(New String() {"OS Station"}) : Install_grid.Rows.Add(New String() {"Limit/Prox Switch"})
        Install_grid.Rows.Add(New String() {"Pressure Switch"}) : Install_grid.Rows.Add(New String() {"Beacons"}) : Install_grid.Rows.Add(New String() {"Horns"})
        Install_grid.Rows.Add(New String() {"Solenoids"}) : Install_grid.Rows.Add(New String() {"Interface form-Robot Controller"}) : Install_grid.Rows.Add(New String() {"Interface Box"})
        Install_grid.Rows.Add(New String() {"Clutch Brake J-Box"}) : Install_grid.Rows.Add(New String() {"I/O Receptacle"}) : Install_grid.Rows.Add(New String() {"Encoder"})
        Install_grid.Rows.Add(New String() {"Scanners w/ Frame ( No commissioning)"}) : Install_grid.Rows.Add(New String() {"Scale Controller Install with frame"}) : Install_grid.Rows.Add(New String() {"Print & Apply (Self Supported)"})
        Install_grid.Rows.Add(New String() {"Print & Apply (With 80/20 Frame)"}) : Install_grid.Rows.Add(New String() {"ADA Box Mounting"}) : Install_grid.Rows.Add(New String() {"24VDC Motor Drops"})
        Install_grid.Rows.Add(New String() {"Open wire basket cable tray (Per Foot)"}) : Install_grid.Rows.Add(New String() {"24VDC Interfaced Zones"}) : Install_grid.Rows.Add(New String() {"2 EMT Chase - Personnel Gate"})
        Install_grid.Rows.Add(New String() {"2' EMT Chase"}) : Install_grid.Rows.Add(New String() {"Sorter Terminations ( Control Only / Includes Cables )"}) : Install_grid.Rows.Add(New String() {"Sorter Terminations ( VFD & Controls )"})
        Install_grid.Rows.Add(New String() {"Sorter Terminations ( Device Net )"}) : Install_grid.Rows.Add(New String() {"Ethernet cabling (Surface Mounted)"}) : Install_grid.Rows.Add(New String() {"Ethernet Terminations"})
        Install_grid.Rows.Add(New String() {"Man Lift per week"}) : Install_grid.Rows.Add(New String() {"Atronix Job Box Shipping (Round Trip)"})

        '-----------------------------------------------------------------------------
        Install_grid.Rows.Add(New String() {"SUPERVISION"})
        Install_grid.Rows(33).DefaultCellStyle.BackColor = Color.Gray
        Install_grid.Rows(33).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Install_grid.Rows(33).ReadOnly = True

        Install_grid.Rows.Add(New String() {"Stand-by & Supervision"})
        Install_grid.Rows.Add(New String() {"Travel Time (per  round trip)"})
        '-------------------------------------------------------------------------------

        Install_grid.Rows.Add(New String() {"EXPENSES"})
        Install_grid.Rows(36).DefaultCellStyle.BackColor = Color.Gray
        Install_grid.Rows(36).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Install_grid.Rows(36).ReadOnly = True

        Install_grid.Rows.Add(New String() {"Travel Expenses (per day)"}) : Install_grid.Rows.Add(New String() {"Travel Expenses (per day) Low Voltage"})
        Install_grid.Rows.Add(New String() {"Mileage"}) : Install_grid.Rows.Add(New String() {"Airfare"}) : Install_grid.Rows.Add(New String() {"Car Rental"})


        '--------------------------------------------------------------------------
        Install_grid.Rows.Add(New String() {"POWER SUBCONTRACTOR"})
        Install_grid.Rows(42).DefaultCellStyle.BackColor = Color.Gray
        Install_grid.Rows(42).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Install_grid.Rows(42).ReadOnly = True


        Install_grid.Rows.Add(New String() {"Contractor Permitting fees"}) : Install_grid.Rows.Add(New String() {"400amp 480v 3PH PDP Install"}) : Install_grid.Rows.Add(New String() {"200amp 480v 3PH PDP Install"})
        Install_grid.Rows.Add(New String() {"125amp 120/240 PDP Install"}) : Install_grid.Rows.Add(New String() {"50KVA 480v-240v XFMR Install"}) : Install_grid.Rows.Add(New String() {"ADA 30a 480v Power Drops ( Price Per)"})
        Install_grid.Rows.Add(New String() {"Air Compressor Wiring W/ Disc. 15HP Max."}) : Install_grid.Rows.Add(New String() {"480v 30a 3-phase 3 wire Ckt. (L.F.)"}) : Install_grid.Rows.Add(New String() {"Motors (Price Per)"})
        Install_grid.Rows.Add(New String() {"Motor Disconnects (Price Per.)"}) : Install_grid.Rows.Add(New String() {"120v 20a Quad Receptacles"}) : Install_grid.Rows.Add(New String() {"120v 20a Ckt (L.F.)"})
        Install_grid.Rows.Add(New String() {"120v 30a Twist Lock Receptacles"}) : Install_grid.Rows.Add(New String() {"120v 30a 1-PH Ckt. (L.F.)"}) : Install_grid.Rows.Add(New String() {"Power Supply Feeds"})
        Install_grid.Rows.Add(New String() {"Robot Controller Feeds"}) : Install_grid.Rows.Add(New String() {"Travel Costs"}) : Install_grid.Rows.Add(New String() {"Demo of existing system"})
        Install_grid.Rows.Add(New String() {"Retro fit exisisting system to new"})

        For j = 0 To 61
            Install_grid.Rows(j).Cells(0).ReadOnly = True
        Next

        Install_grid.Rows(34).Cells(6).ReadOnly = False
        Install_grid.Rows(35).Cells(6).ReadOnly = False

        '///////////////////////////// Start //////////////////////////

        my_assemblies = New List(Of String)()
        my_kits = New List(Of String)()


        Try
            '--------------  add to kit
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT distinct Legacy_ADA_Number from kits"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    my_kits.Add(reader(0))
                End While
            End If

            reader.Close()

            '--------------  add to device
            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assemblies.Add(reader2(0))
                End While
            End If

            reader2.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        'Load my features codes from features code table


        'This list will store the specific type of feature code of each type
        type_list_panel = New List(Of String)()
        type_list_plc = New List(Of String)()
        type_list_field = New List(Of String)()
        type_list_control = New List(Of String)()

        'set default solution and populate apanel combobox
        Try

            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT solution_name from quote_table.feature_solutions order by solution_name desc"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    apanel_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

            ' apanel_box.Text = "SWDMS/EIPRIO"
            apanel_box.Text = "EIP-MS/EIP-RIO"
            sol_label.Text = apanel_box.Text

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        Try


            '-------- Specific type -------------
            Dim cmd_panel As New MySqlCommand
            cmd_panel.Parameters.AddWithValue("@sol", sol_label.Text)
            cmd_panel.CommandText = "SELECT distinct specific_type, type from quote_table.feature_codes where show_menu = 'Y' and Solution = @sol"
            cmd_panel.Connection = Login.Connection
            Dim reader_panel As MySqlDataReader
            reader_panel = cmd_panel.ExecuteReader

            If reader_panel.HasRows Then
                While reader_panel.Read
                    If String.Equals(reader_panel(1).ToString, "Panel") = True Then
                        type_list_panel.add(reader_panel(0).ToString)
                    ElseIf String.Equals(reader_panel(1).ToString, "PLC") = True Then
                        type_list_plc.add(reader_panel(0).ToString)
                    ElseIf String.Equals(reader_panel(1).ToString, "Field") = True Then
                        type_list_field.add(reader_panel(0).ToString)
                    ElseIf String.Equals(reader_panel(1).ToString, "Control Panel") = True Then
                        type_list_control.add(reader_panel(0).ToString)
                    End If
                End While
            End If

            reader_panel.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        Dim cmd_dt As New MySqlCommand
        Dim reader_panel2 As MySqlDataReader
        Dim reader_plc As MySqlDataReader
        Dim reader_field As MySqlDataReader
        Dim reader_control As MySqlDataReader

        Dim c_1 As Integer : c_1 = 0
        Dim c_2 As Integer : c_2 = 0
        Dim c_3 As Integer : c_3 = 0
        Dim c_4 As Integer : c_4 = 0

        Try
            '--------- Panel ------
            For j = 0 To type_list_panel.count - 1


                Panel_grid.Rows.Add(New String() {type_list_panel(j)})  'add type line
                Panel_grid.Rows(c_1).DefaultCellStyle.BackColor = Color.Goldenrod
                Panel_grid.Rows(c_1).ReadOnly = True

                cmd_dt.Parameters.Clear()
                cmd_dt.Parameters.AddWithValue("@s_type", type_list_panel(j))
                cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'AB' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'panel' and Solution = @sol"
                cmd_dt.Connection = Login.Connection
                reader_panel2 = cmd_dt.ExecuteReader

                If reader_panel2.HasRows Then
                    While reader_panel2.Read
                        Panel_grid.Rows.Add(New String() {reader_panel2(0).ToString})
                        c_1 = c_1 + 1
                    End While
                End If
                reader_panel2.Close()

                c_1 = c_1 + 1
            Next

            '-------------- PLC ------------
            For j = 0 To type_list_plc.count - 1

                PLC_grid.Rows.Add(New String() {type_list_plc(j)})  'add type line
                PLC_grid.Rows(c_2).DefaultCellStyle.BackColor = Color.Goldenrod
                PLC_grid.Rows(c_2).ReadOnly = True

                cmd_dt.Parameters.Clear()
                cmd_dt.Parameters.AddWithValue("@s_type", type_list_plc(j))
                cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'AB' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'PLC' and Solution = @sol"
                cmd_dt.Connection = Login.Connection
                reader_plc = cmd_dt.ExecuteReader

                If reader_plc.HasRows Then
                    While reader_plc.Read
                        PLC_grid.Rows.Add(New String() {reader_plc(0).ToString})
                        c_2 = c_2 + 1
                    End While
                End If
                reader_plc.Close()

                c_2 = c_2 + 1
            Next

            '-------------- FIELD -----------------
            For j = 0 To type_list_field.count - 1

                Field_grid.Rows.Add(New String() {type_list_field(j)})  'add type line
                Field_grid.Rows(c_3).DefaultCellStyle.BackColor = Color.Goldenrod
                Field_grid.Rows(c_3).ReadOnly = True

                cmd_dt.Parameters.Clear()
                cmd_dt.Parameters.AddWithValue("@s_type", type_list_field(j))
                'default solution EIP-MS/EIP-RIO, Please change it manually for now
                cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'AB' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'Field' and Solution = 'EIP-MS/EIP-RIO'"
                cmd_dt.Connection = Login.Connection
                reader_field = cmd_dt.ExecuteReader

                If reader_field.HasRows Then
                    While reader_field.Read
                        Field_grid.Rows.Add(New String() {reader_field(0).ToString})
                        c_3 = c_3 + 1
                    End While
                End If
                reader_field.Close()

                c_3 = c_3 + 1
            Next

            '--------- Control Panel ----------------
            For j = 0 To type_list_control.count - 1

                Control_grid.Rows.Add(New String() {type_list_control(j)})  'add type line
                Control_grid.Rows(c_4).DefaultCellStyle.BackColor = Color.Goldenrod
                Control_grid.Rows(c_4).ReadOnly = True

                cmd_dt.Parameters.Clear()
                cmd_dt.Parameters.AddWithValue("@s_type", type_list_control(j))
                cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'AB' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'Control Panel' and Solution = @sol"
                cmd_dt.Connection = Login.Connection
                reader_control = cmd_dt.ExecuteReader

                If reader_control.HasRows Then
                    While reader_control.Read
                        Control_grid.Rows.Add(New String() {reader_control(0).ToString})
                        c_4 = c_4 + 1
                    End While
                End If
                reader_control.Close()

                c_4 = c_4 + 1
            Next


            '---------------- solution settings --------------
            Dim cmd_sol As New MySqlCommand
            cmd_sol.CommandText = "SELECT feature_code, description, solution, solution_description from quote_table.feature_codes where solution_default = 'Y' and type = 'Field'"
            cmd_sol.Connection = Login.Connection
            Dim reader_sol As MySqlDataReader
            reader_sol = cmd_sol.ExecuteReader

            datatable = New DataTable 'this table will contain solution settings
            datatable.Columns.Add("Feature codes", GetType(String))
            datatable.Columns.Add("Description", GetType(String))
            datatable.Columns.Add("solution", GetType(String))
            datatable.Columns.Add("Solution description", GetType(String))


            If reader_sol.HasRows Then
                While reader_sol.Read
                    datatable.Rows.Add(reader_sol(0).ToString, reader_sol(1).ToString, reader_sol(2).ToString, reader_sol(3).ToString)
                End While
            End If

            reader_sol.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '================== create dummy datatable =================
        scapegoat = New DataTable
        scapegoat.Columns.Add("parts", GetType(String))
        scapegoat.Columns.Add("qty", GetType(String))
        scapegoat.Columns.Add("set_name", GetType(String))
        scapegoat.Columns.Add("type", GetType(String))

        '=================== Panel datatable ====================
        dimen_table = New DataTable
        dimen_table.Columns.Add("ADA_Set", GetType(String))
        dimen_table.Columns.Add("Mode", GetType(String))
        dimen_table.Columns.Add("24_nd", GetType(String))
        dimen_table.Columns.Add("24_qty", GetType(String))
        dimen_table.Columns.Add("30_nd", GetType(String))
        dimen_table.Columns.Add("30_qty", GetType(String))

        '================== create allocation datatable =================
        my_alloc_table = New DataTable
        my_alloc_table.Columns.Add("parts", GetType(String))
        my_alloc_table.Columns.Add("part_description", GetType(String))
        my_alloc_table.Columns.Add("Manufacturer", GetType(String))
        my_alloc_table.Columns.Add("vendor", GetType(String))
        my_alloc_table.Columns.Add("cost", GetType(String))
        my_alloc_table.Columns.Add("qty", GetType(String))
        my_alloc_table.Columns.Add("subtotal", GetType(String))
        my_alloc_table.Columns.Add("type", GetType(String)) 'mfg type
        my_alloc_table.Columns.Add("g_type", GetType(String)) 'quote_type (starter panel, plc, field, a-panel)
        my_alloc_table.Columns.Add("ADA", GetType(String)) 'ada number

        '================ FLA datatable ==================
        fla_data = New DataTable
        fla_data.Columns.Add("qty", GetType(Double))
        fla_data.Columns.Add("fla", GetType(Double))
        fla_data.Columns.Add("set", GetType(String))

        '///////////// ==================  datatable to store (HERE MY TABLES)  =================== ////////////////////

        table_Feature_code = New DataTable
        table_Feature_code.Columns.Add("Feature_code", GetType(String))
        table_Feature_code.Columns.Add("description", GetType(String))
        table_Feature_code.Columns.Add("Solution", GetType(String))
        table_Feature_code.Columns.Add("type", GetType(String))
        table_Feature_code.Columns.Add("labor_cost", GetType(Double))
        table_Feature_code.Columns.Add("bulk_cost", GetType(Double))


        table_feature_parts = New DataTable
        table_feature_parts.Columns.Add("Feature_code", GetType(String))
        table_feature_parts.Columns.Add("solution", GetType(String))
        table_feature_parts.Columns.Add("type", GetType(String))
        table_feature_parts.Columns.Add("part_name", GetType(String))
        table_feature_parts.Columns.Add("qty", GetType(Double))

        table_f_dimensions = New DataTable
        table_f_dimensions.Columns.Add("feature_code", GetType(String))
        table_f_dimensions.Columns.Add("Solution", GetType(String))
        table_f_dimensions.Columns.Add("Full_DR", GetType(Double))
        table_f_dimensions.Columns.Add("Half_DR", GetType(Double))
        table_f_dimensions.Columns.Add("FLA", GetType(Double))

        table_call_feature = New DataTable
        table_call_feature.Columns.Add("feature_code", GetType(String))
        table_call_feature.Columns.Add("description", GetType(String))
        table_call_feature.Columns.Add("solution", GetType(String))
        table_call_feature.Columns.Add("formula", GetType(String))


        table_remote_io = New DataTable
        table_remote_io.Columns.Add("feature_code", GetType(String))
        table_remote_io.Columns.Add("solutiom", GetType(String))
        table_remote_io.Columns.Add("inputs", GetType(Integer))
        table_remote_io.Columns.Add("output", GetType(Integer))
        table_remote_io.Columns.Add("motion", GetType(Integer))
        table_remote_io.Columns.Add("description", GetType(String))

        table_io_points = New DataTable
        table_io_points.Columns.Add("feature_code", GetType(String))
        table_io_points.Columns.Add("description", GetType(String))
        table_io_points.Columns.Add("solution", GetType(String))
        table_io_points.Columns.Add("input", GetType(Integer))
        table_io_points.Columns.Add("output", GetType(Integer))
        table_io_points.Columns.Add("motion", GetType(Integer))

        '--- parts
        table_parts = New DataTable
        table_parts.Columns.Add("Part_Name", GetType(String))
        table_parts.Columns.Add("Part_Description", GetType(String))
        table_parts.Columns.Add("Manufacturer", GetType(String))
        table_parts.Columns.Add("Primary_Vendor", GetType(String))
        table_parts.Columns.Add("MFG_type", GetType(String))
        table_parts.Columns.Add("Legacy_ADA_Number", GetType(String))


        '---vendors
        table_vendors = New DataTable
        table_vendors.Columns.Add("Part_Name", GetType(String))
        table_vendors.Columns.Add("Vendor_Name", GetType(String))
        table_vendors.Columns.Add("Cost", GetType(Decimal))
        table_vendors.Columns.Add("Purchase_Date", GetType(Date))


        '---- assem
        table_assem = New DataTable
        table_assem.Columns.Add("Legacy_ADA_Number", GetType(String))
        table_assem.Columns.Add("Labor_Cost", GetType(Decimal))
        table_assem.Columns.Add("Bulk_Cost", GetType(Decimal))


        '----- adv
        table_adv = New DataTable
        table_adv.Columns.Add("Part_Name", GetType(String))
        table_adv.Columns.Add("ADV_Number", GetType(String))
        table_adv.Columns.Add("Qty", GetType(Double))
        table_adv.Columns.Add("Legacy_ADA", GetType(String))

        Try
            Dim cmd_panel As New MySqlCommand
            cmd_panel.CommandText = "SELECT Feature_code, description, Solution, type, labor_cost, bulk_cost from quote_table.feature_codes"
            cmd_panel.Connection = Login.Connection
            Dim reader_panel As MySqlDataReader
            reader_panel = cmd_panel.ExecuteReader

            If reader_panel.HasRows Then
                While reader_panel.Read
                    table_Feature_code.Rows.Add(reader_panel(0).ToString, reader_panel(1).ToString, reader_panel(2), reader_panel(3).ToString, If(IsNumeric(reader_panel(4)) = False, 0, reader_panel(4)), If(IsNumeric(reader_panel(5)) = False, 0, reader_panel(5)))
                End While
            End If

            reader_panel.Close()

            '--- feature_parts --
            Dim cmd_panel1 As New MySqlCommand
            cmd_panel1.CommandText = "SELECT * from quote_table.feature_parts"
            cmd_panel1.Connection = Login.Connection
            Dim reader_panel1 As MySqlDataReader
            reader_panel1 = cmd_panel1.ExecuteReader

            If reader_panel1.HasRows Then
                While reader_panel1.Read
                    table_feature_parts.Rows.Add(reader_panel1(0).ToString, reader_panel1(1).ToString, reader_panel1(2).ToString, reader_panel1(3).ToString, reader_panel1(4).ToString)
                End While
            End If

            reader_panel1.Close()

            '---- f dimensions
            Dim cmd_panel2 As New MySqlCommand
            cmd_panel2.CommandText = "SELECT * from quote_table.f_dimensions"
            cmd_panel2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd_panel2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    table_f_dimensions.Rows.Add(reader2(0), reader2(1), reader2(2), reader2(3), reader2(4))
                End While
            End If

            reader2.Close()

            '----- call function
            Dim cmd_panel3 As New MySqlCommand
            cmd_panel3.CommandText = "SELECT * from quote_table.call_feature"
            cmd_panel3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd_panel3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    table_call_feature.Rows.Add(reader3(0), reader3(1), reader3(2), reader3(3))
                End While
            End If

            reader3.Close()


            '------------ remote_io
            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT feature_code, solutiom, inputs, output, motion, description  from quote_table.Remote_IO"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    table_remote_io.Rows.Add(reader4(0), reader4(1), reader4(2), reader4(3), reader4(4), reader4(5))
                End While
            End If

            reader4.Close()

            '---------- io_points
            Dim cmd5 As New MySqlCommand
            cmd5.CommandText = "SELECT feature_code, description, solution, input, output, motion  from quote_table.IO_points"
            cmd5.Connection = Login.Connection
            Dim reader5 As MySqlDataReader
            reader5 = cmd5.ExecuteReader

            If reader5.HasRows Then
                While reader5.Read
                    table_io_points.Rows.Add(reader5(0), reader5(1), reader5(2), reader5(3), reader5(4), reader5(5))
                End While
            End If

            reader5.Close()

            '---------- parts
            Dim cmd6 As New MySqlCommand
            cmd6.CommandText = "SELECT Part_Name, Part_Description, Manufacturer, Primary_Vendor, MFG_type, Legacy_ADA_Number from parts_table"
            cmd6.Connection = Login.Connection
            Dim reader6 As MySqlDataReader
            reader6 = cmd6.ExecuteReader

            If reader6.HasRows Then
                While reader6.Read
                    table_parts.Rows.Add(reader6(0), reader6(1), reader6(2), reader6(3), reader6(4), reader6(5))
                End While
            End If

            reader6.Close()

            '----------- vendors
            Dim cmd7 As New MySqlCommand
            cmd7.CommandText = "SELECT Part_Name, Vendor_Name, Cost, Purchase_Date from vendors_table"
            cmd7.Connection = Login.Connection
            Dim reader7 As MySqlDataReader
            reader7 = cmd7.ExecuteReader

            If reader7.HasRows Then
                While reader7.Read
                    table_vendors.Rows.Add(reader7(0), reader7(1), reader7(2), reader7(3))
                End While
            End If

            reader7.Close()

            '----------- assem
            Dim cmd8 As New MySqlCommand
            cmd8.CommandText = "SELECT Legacy_ADA_Number, Labor_Cost, Bulk_Cost  from devices"
            cmd8.Connection = Login.Connection
            Dim reader8 As MySqlDataReader
            reader8 = cmd8.ExecuteReader

            If reader8.HasRows Then
                While reader8.Read
                    table_assem.Rows.Add(reader8(0), reader8(1), reader8(2))
                End While
            End If

            reader8.Close()

            '---------- adv
            Dim cmd9 As New MySqlCommand
            cmd9.CommandText = "SELECT Part_Name, ADV_Number, Qty, Legacy_ADA from adv"
            cmd9.Connection = Login.Connection
            Dim reader9 As MySqlDataReader
            reader9 = cmd9.ExecuteReader

            If reader9.HasRows Then
                While reader9.Read
                    table_adv.Rows.Add(reader9(0), reader9(1), reader9(2), reader9(3))
                End While
            End If

            reader9.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        '--------------------------------------------------------

        counter = 1


        '------- default solution -------

        go_on = False
        start_recal = True
        change_sol = False

        Call Fill_installation() 'load fixed installation values

    End Sub

    Private Sub Install_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Install_grid.CellValueChanged
        Call Cal_install()
    End Sub

    Sub Cal_install()

        'Spreadsheet like calculations installation grid
        If cal_ins = True Then

            For i = 1 To Install_grid.Rows.Count - 1
                'check if one of the cells has a number
                If Install_grid.Rows(i).IsNewRow Then Continue For  'new row skip

                If (IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(2).Value())) Then
                    If i <> 2 Then
                        'if one cell is empty or has a non number turn into zero
                        Install_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = False, 0, Install_grid.Rows(i).Cells(1).Value())
                        Install_grid.Rows(i).Cells(2).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(2).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(2).Value()) = False, 0, Install_grid.Rows(i).Cells(2).Value())

                        Install_grid.Rows(i).Cells(3).Value() = Install_grid.Rows(i).Cells(1).Value() * Install_grid.Rows(i).Cells(2).Value()
                    Else

                        Install_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = False, 0, Install_grid.Rows(i).Cells(1).Value())
                        Install_grid.Rows(i).Cells(2).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(2).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(2).Value()) = False, 0, Install_grid.Rows(i).Cells(2).Value())

                        Install_grid.Rows(i).Cells(3).Value() = Install_grid.Rows(i).Cells(1).Value() * Install_grid.Rows(i).Cells(2).Value() '/ 10
                    End If
                End If

                '-----------------------
                If (IsNumeric(Install_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value())) Then
                    Install_grid.Rows(i).Cells(3).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(3).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value()) = False, 0, Install_grid.Rows(i).Cells(3).Value())
                    Install_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(4).Value()) = False, 0, Install_grid.Rows(i).Cells(4).Value())

                    Install_grid.Rows(i).Cells(5).Value() = (Install_grid.Rows(i).Cells(4).Value() * Install_grid.Rows(i).Cells(3).Value())
                End If

                '----------------------
                If (IsNumeric(Install_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value())) Then

                    Install_grid.Rows(i).Cells(6).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(6).Value()) = False, 0, Install_grid.Rows(i).Cells(6).Value())
                    Install_grid.Rows(i).Cells(3).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(3).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value()) = False, 0, Install_grid.Rows(i).Cells(3).Value())

                    Install_grid.Rows(i).Cells(7).Value() = (Install_grid.Rows(i).Cells(6).Value() * Install_grid.Rows(i).Cells(3).Value())

                End If

                '-----------------------------
                If (IsNumeric(Install_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value())) Then

                    Install_grid.Rows(i).Cells(8).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(8).Value()) = False, 0, Install_grid.Rows(i).Cells(8).Value())
                    Install_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = False, 0, Install_grid.Rows(i).Cells(1).Value())

                    Install_grid.Rows(i).Cells(9).Value() = (Install_grid.Rows(i).Cells(8).Value() * Install_grid.Rows(i).Cells(1).Value())
                End If

                '-------------------------------------
                If (IsNumeric(Install_grid.Rows(i).Cells(10).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value())) Then

                    Install_grid.Rows(i).Cells(10).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(10).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(10).Value()) = False, 0, Install_grid.Rows(i).Cells(10).Value())
                    Install_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = False, 0, Install_grid.Rows(i).Cells(1).Value())

                    Install_grid.Rows(i).Cells(11).Value() = Install_grid.Rows(i).Cells(10).Value() * Install_grid.Rows(i).Cells(1).Value()
                End If

                '-------------------------------------
                If (IsNumeric(Install_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(7).Value()) Or IsNumeric(Install_grid.Rows(i).Cells(9).Value()) Or IsNumeric(Install_grid.Rows(i).Cells(11).Value())) Then

                    Install_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(5).Value()) = False, 0, Install_grid.Rows(i).Cells(5).Value())
                    Install_grid.Rows(i).Cells(7).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(7).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(7).Value()) = False, 0, Install_grid.Rows(i).Cells(7).Value())
                    Install_grid.Rows(i).Cells(9).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(9).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(9).Value()) = False, 0, Install_grid.Rows(i).Cells(9).Value())
                    Install_grid.Rows(i).Cells(11).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(11).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(11).Value()) = False, 0, Install_grid.Rows(i).Cells(11).Value())

                    Install_grid.Rows(i).Cells(12).Value() = "$" & (Install_grid.Rows(i).Cells(5).Value() + Install_grid.Rows(i).Cells(7).Value() + Install_grid.Rows(i).Cells(9).Value() + Install_grid.Rows(i).Cells(11).Value())

                End If

            Next


            labor = 0
            materials = 0
            expenses = 0
            subcontract = 0
            totales = 0

            For i = 1 To Install_grid.Rows.Count - 1
                If IsNumeric(Install_grid.Rows(i).Cells(7).Value()) = True Then
                    labor = labor + Install_grid.Rows(i).Cells(7).Value()
                End If
                If IsNumeric(Install_grid.Rows(i).Cells(9).Value()) = True Then
                    materials = materials + Install_grid.Rows(i).Cells(9).Value()
                End If
                If IsNumeric(Install_grid.Rows(i).Cells(11).Value()) = True Then
                    expenses = expenses + Install_grid.Rows(i).Cells(11).Value()
                End If
                If IsNumeric(Install_grid.Rows(i).Cells(5).Value()) = True Then
                    subcontract = subcontract + Install_grid.Rows(i).Cells(5).Value()
                End If
            Next

            totales = labor + materials + expenses + subcontract

            labor_l.Text = "Labor: $" & labor
            materials_l.Text = "Materials: $" & materials
            expenses_l.Text = "Expenses: $" & expenses
            subcontract_l.Text = "Subcontract: $" & subcontract
            total_l.Text = "Total: $" & totales



            If start_recal = True Then
                Quote_grid.Rows(79).Cells(5).Value() = "$" & labor
                Quote_grid.Rows(80).Cells(5).Value() = "$" & materials
                Quote_grid.Rows(81).Cells(5).Value() = "$" & expenses
                Quote_grid.Rows(82).Cells(5).Value() = "$" & subcontract
            End If

        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        volts_575.Checked = False
        volts_480.Checked = True
        volts_230.Checked = False
    End Sub

    Private Sub volts_575_Click(sender As Object, e As EventArgs) Handles volts_575.Click
        volts_575.Checked = True
        volts_480.Checked = False
        volts_230.Checked = False
    End Sub

    Private Sub volts_230_Click(sender As Object, e As EventArgs) Handles volts_230.Click
        volts_575.Checked = False
        volts_480.Checked = False
        volts_230.Checked = True
    End Sub



    Private Sub PF525ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PF525_c.Click

        Dim result As DialogResult = MessageBox.Show("Current Quote changes will be deleted. Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then

            PF525_c.Checked = True
            DE1_c.Checked = False

            Panel_grid.Rows.Clear()

            Dim list_panel = New List(Of String)()
            Dim c_1 As Integer : c_1 = 0

            Try
                '-------- Panel Specific type -------------
                Dim cmd_panel As New MySqlCommand
                cmd_panel.Parameters.AddWithValue("@sol", sol_label.Text)
                cmd_panel.CommandText = "SELECT distinct specific_type, type from quote_table.feature_codes where show_menu = 'Y' and Solution = @sol"
                cmd_panel.Connection = Login.Connection
                Dim reader_panel As MySqlDataReader
                reader_panel = cmd_panel.ExecuteReader

                If reader_panel.HasRows Then
                    While reader_panel.Read
                        If String.Equals(reader_panel(1).ToString, "Panel") = True Then
                            list_panel.Add(reader_panel(0).ToString)
                        End If
                    End While
                End If

                reader_panel.Close()


                Dim cmd_dt As New MySqlCommand
                Dim reader_panel2 As MySqlDataReader

                For j = 0 To list_panel.Count - 1

                    Panel_grid.Rows.Add(New String() {list_panel(j)})  'add type line
                    Panel_grid.Rows(c_1).DefaultCellStyle.BackColor = Color.Goldenrod
                    Panel_grid.Rows(c_1).ReadOnly = True

                    cmd_dt.Parameters.Clear()
                    cmd_dt.Parameters.AddWithValue("@s_type", list_panel(j))
                    cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                    cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'AB' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'panel' and Solution = @sol"
                    cmd_dt.Connection = Login.Connection
                    reader_panel2 = cmd_dt.ExecuteReader

                    If reader_panel2.HasRows Then
                        While reader_panel2.Read
                            Panel_grid.Rows.Add(New String() {reader_panel2(0).ToString})
                            c_1 = c_1 + 1
                        End While
                    End If
                    reader_panel2.Close()

                    c_1 = c_1 + 1
                Next


                If Allocation_parts.Visible = True Then
                    Allocation_parts.alloc_grid.Rows.Clear()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub DE1_c_Click(sender As Object, e As EventArgs) Handles DE1_c.Click
        Dim result As DialogResult = MessageBox.Show("Current Quote changes will be deleted. Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then
            PF525_c.Checked = False
            DE1_c.Checked = True

            Panel_grid.Rows.Clear()

            Dim list_panel = New List(Of String)()
            Dim c_1 As Integer : c_1 = 0

            Try
                '-------- Panel Specific type -------------
                Dim cmd_panel As New MySqlCommand
                cmd_panel.Parameters.AddWithValue("@sol", sol_label.Text)
                cmd_panel.CommandText = "SELECT distinct specific_type, type from quote_table.feature_codes where show_menu = 'Y' and Solution = @sol"
                cmd_panel.Connection = Login.Connection
                Dim reader_panel As MySqlDataReader
                reader_panel = cmd_panel.ExecuteReader

                If reader_panel.HasRows Then
                    While reader_panel.Read
                        If String.Equals(reader_panel(1).ToString, "Panel") = True Then
                            list_panel.Add(reader_panel(0).ToString)
                        End If
                    End While
                End If

                reader_panel.Close()


                Dim cmd_dt As New MySqlCommand
                Dim reader_panel2 As MySqlDataReader

                For j = 0 To list_panel.Count - 1

                    Panel_grid.Rows.Add(New String() {list_panel(j)})  'add type line
                    Panel_grid.Rows(c_1).DefaultCellStyle.BackColor = Color.Goldenrod
                    Panel_grid.Rows(c_1).ReadOnly = True

                    cmd_dt.Parameters.Clear()
                    cmd_dt.Parameters.AddWithValue("@s_type", list_panel(j))
                    cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                    cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = 'Eaton' or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'panel' and Solution = @sol"
                    cmd_dt.Connection = Login.Connection
                    reader_panel2 = cmd_dt.ExecuteReader

                    If reader_panel2.HasRows Then
                        While reader_panel2.Read
                            Panel_grid.Rows.Add(New String() {reader_panel2(0).ToString})
                            c_1 = c_1 + 1
                        End While
                    End If
                    reader_panel2.Close()

                    c_1 = c_1 + 1
                Next

                If Allocation_parts.Visible = True Then
                    Allocation_parts.alloc_grid.Rows.Clear()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' ADD a new ADA Set
        Set_wizard.ShowDialog()

    End Sub

    Private Sub ClearAllADAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllADAToolStripMenuItem.Click
        'remove all ada columns
        For i = Panel_grid.Columns.Count - 1 To 1 Step -1
            Panel_grid.Columns.RemoveAt(i)
            PLC_grid.Columns.RemoveAt(i)
            Field_grid.Columns.RemoveAt(i)
            Control_grid.Columns.RemoveAt(i)
            fla_set_grid.Columns.RemoveAt(i)
        Next

        counter = 1
        my_Set_info.Clear()

        If Allocation_parts.Visible = True Then
            Allocation_parts.alloc_grid.Rows.Clear()
            M12_grid.Rows.Clear()
            M12_ES_grid.Rows.Clear()
        End If

        Call General_calculation()

    End Sub

    Sub General_calculation()

        If go_on = False Then

            '////////////// This is the main subroutine that it will be called everytime the a cell in all the datagrids (panel, projetc, plc, ...) is changed /////////////////////// 

            ' reset values, clear tables

            do_not_IO_recal = True
            my_alloc_table.Rows.Clear()
            fla_data.Rows.Clear()
            bulk_t = 0
            labor_t = 0


            For Each kvp As KeyValuePair(Of String, Double) In M12_cables.ToArray
                M12_cables(kvp.Key) = 0
            Next

            For Each kvp As KeyValuePair(Of String, Double) In M12_ES_cables.ToArray
                M12_ES_cables(kvp.Key) = 0
            Next


            If Allocation_parts.Visible = True Then
                Allocation_parts.alloc_grid.Rows.Clear() 'clear alloc datagrid
            End If

            ' Clear Dummy_quote Table
            'Try
            '    Dim cmd As New MySqlCommand
            '    cmd.Parameters.AddWithValue("@user", current_user)
            '    cmd.CommandText = "delete from quote_table.Dummy_quote where current_u = @user"
            '    cmd.Connection = Login.Connection
            '    cmd.ExecuteNonQuery()
            'Catch ex As Exception
            '    MessageBox.Show(ex.ToString)
            'End Try

            '---------- Panel section ---------------
            For i = 1 To Panel_grid.Columns.Count - 1
                For j = 0 To Panel_grid.Rows.Count - 1
                    If IsNumeric(Panel_grid.Rows(j).Cells(i).Value()) = True Then
                        Call Feature_code_solve2(Panel_grid.Rows(j).Cells(0).Value(), Math.Ceiling(CType(Panel_grid.Rows(j).Cells(i).Value(), Double)), sol_label.Text, "Panel", Panel_grid.Columns.Item(i).HeaderText)
                    End If
                Next

            Next

            '---------- PLC section ---------------
            For i = 1 To PLC_grid.Columns.Count - 1
                For j = 0 To PLC_grid.Rows.Count - 1
                    If IsNumeric(PLC_grid.Rows(j).Cells(i).Value()) = True Then
                        Call Feature_code_solve2(PLC_grid.Rows(j).Cells(0).Value(), Math.Ceiling(CType(PLC_grid.Rows(j).Cells(i).Value(), Double)), sol_label.Text, "PLC", PLC_grid.Columns.Item(i).HeaderText)
                    End If
                Next
            Next


            '------------- Field section ---------------
            For i = 1 To Field_grid.Columns.Count - 1
                For j = 0 To Field_grid.Rows.Count - 1
                    If IsNumeric(Field_grid.Rows(j).Cells(i).Value()) = True Then
                        Call Feature_code_solve2(Field_grid.Rows(j).Cells(0).Value(), Math.Ceiling(CType(Field_grid.Rows(j).Cells(i).Value(), Double)), find_solution(Field_grid.Rows(j).Cells(0).Value()), "Field", Field_grid.Columns.Item(i).HeaderText)
                    End If
                Next
            Next


            '------------- Control Panel section ---------------
            For i = 1 To Control_grid.Columns.Count - 1
                For j = 0 To Control_grid.Rows.Count - 1
                    If IsNumeric(Control_grid.Rows(j).Cells(i).Value()) = True Then
                        Call Feature_code_solve2(Control_grid.Rows(j).Cells(0).Value(), Math.Ceiling(CType(Control_grid.Rows(j).Cells(i).Value(), Double)), sol_label.Text, "Control Panel", Control_grid.Columns.Item(i).HeaderText)
                    End If
                Next
            Next

            '------- A panel allocation ---------- (old solution do not erase)
            If Include_A.Checked = True Then
                If Panel_grid.Columns.Count > 1 Then
                    Call Feature_code_solve2("A-Panels", Panel_grid.Columns.Count - 1, sol_label.Text, "A-Panel", Panel_grid.Columns.Item(1).HeaderText)
                End If
            End If


            '-------------- extra allocation of parts ----------------
            If Panel_grid.Columns.Count > 1 Then

                'ESB is the sum of all Estop devices
                Dim ESB As Double : ESB = get_field_qty("E-Stop Box", Control_grid) + get_field_qty("ESTOP-PB (RD W/ LT), ADA-ASM-EPB101", Field_grid) + get_field_qty("GUARDED ESTOP-PB (RD W/ LT), ADA-ASM-EPB102", Field_grid) + get_field_qty("ESTOP PULL CORD SINGLE, ADA-ASM-EPC101", Field_grid) + get_field_qty("ESTOP PULL CORD DOUBLE, ADA-ASM-EPC102", Field_grid)
                Dim MS As Double : MS = get_field_qty("MS_Panel", Panel_grid)
                Dim RIOM As Double : RIOM = get_field_qty("Remote IO IO-Link Master 8IOL", Field_grid)

                Call Feature_code_solve2("Ethernet Connectors", (ESB + MS + RIOM), "EIP-MS/EIP-RIO", "Field", Panel_grid.Columns.Item(1).HeaderText)
                Call Feature_code_solve2("M12-L Coded Connectors", (ESB + MS + RIOM), "EIP-MS/EIP-RIO", "Field", Panel_grid.Columns.Item(1).HeaderText)


                Call M12_process2()  'calculate M12 and M12_E cables
                Call Estops_alloc2(ESB) 'calculate EStop cables
                Call RIO_analyze()  'load the RIO analyzer
                Call FLA_load() 'fill fla table
                Call Create_Main_table2()  'create datatable and datagrid (partcost)
                Call alloc_Totals() 'this will put the total cost in the quote form

                switch_b = True 'reset switch
                IO_custom = False

            End If

            RIO_press = False
            do_not_IO_recal = False
            Call BOM_refresh()

        End If

    End Sub


    '---------//////////////// Feature code solver, add parts to dummy table where all the operations are going to take place //////////////////////-------

    Sub Feature_code_solve(feature_desc As String, qty As Double, solution As String, type As String, set_name As String)

        '------------ preparation and clear tables
        Dim got_assemblies = New List(Of String)()
        Dim got_kits = New List(Of String)()
        Dim my_Feature_code As String : my_Feature_code = ""


        '--------------- Start collecting parts according to feature code ------------

        Try

            '------- Get Feature code from feature description -------- (Note: I may add this in the feature_parts to make everything faster)
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@feature_desc", feature_desc)
            cmd.Parameters.AddWithValue("@sol", solution)
            cmd.CommandText = "SELECT Feature_code, labor_cost, bulk_cost from quote_table.feature_codes where description = @feature_desc and Solution = @sol"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    my_Feature_code = reader(0).ToString
                    bulk_t = bulk_t + (If(IsNumeric(reader(1)) = True, reader(1), 0)) * qty
                    labor_t = labor_t + (If(IsNumeric(reader(2)) = True, reader(2), 0)) * qty
                End While
            End If

            reader.Close()


            '------------------------------------
            '------------ Get parts and qty ----------------
            Dim cmd2 As New MySqlCommand
            cmd2.Parameters.AddWithValue("@feature_code", my_Feature_code)
            cmd2.Parameters.AddWithValue("@solution", solution)
            cmd2.Parameters.AddWithValue("@type", type)
            cmd2.CommandText = "SELECT part_name, qty from quote_table.feature_parts where Feature_code = @feature_code and solution = @solution and type = @type"

            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read

                    'If my_assemblies.contains(reader2(0).ToString) = True Then
                    '    got_assemblies.Add(reader2(0).ToString)
                    'ElseIf my_kits.contains(reader2(0).ToString) = True Then
                    '    got_kits.Add(reader2(0).ToString)
                    'Else
                    Call Insert_into_dummy_table(reader2(0).ToString, Math.Ceiling(CType(reader2(1).ToString, Double) * qty), set_name, type)  'insert data to dummy datatable
                    '  End If

                End While
            End If

            reader2.Close()

            Call insert_into_mysql_dummy_Table()  'insert data to mysql dummy table
            scapegoat.Rows.Clear() 'clear dummy datatable


            '---------- get assemblies and kits parts --------------

            'For Each assembly In got_assemblies
            '    Dim cmd3 As New MySqlCommand
            '    cmd3.Parameters.AddWithValue("@ADA", assembly)
            '    cmd3.CommandText = "SELECT Part_Name, Qty from adv where Legacy_ADA = @ADA"
            '    cmd3.Connection = Login.Connection
            '    Dim reader3 As MySqlDataReader
            '    reader3 = cmd3.ExecuteReader

            '    If reader3.HasRows Then
            '        While reader3.Read
            '            Call Insert_into_dummy_table(reader3(0).ToString, CType(reader3(1).ToString, Double) * qty, set_name, type)
            '        End While
            '    End If

            '    reader3.Close()
            'Next

            'Call insert_into_mysql_dummy_Table() 'insert data to mysql dummy table
            'scapegoat.Rows.Clear() 'clear dummy datatable

            ''-------------
            'For Each kit In got_kits
            '    Dim cmd4 As New MySqlCommand
            '    cmd4.Parameters.AddWithValue("@ADA2", kit)
            '    cmd4.CommandText = "SELECT Part_Name, Qty from akn where Legacy_ADA = @ADA2"
            '    cmd4.Connection = Login.Connection
            '    Dim reader4 As MySqlDataReader
            '    reader4 = cmd4.ExecuteReader

            '    If reader4.HasRows Then
            '        While reader4.Read
            '            Call Insert_into_dummy_table(reader4(0).ToString, CType(reader4(1).ToString, Double) * qty, set_name, type)
            '        End While
            '    End If

            '    reader4.Close()

            'Next

            '================== FLA calculation ========================
            If String.Equals(type, "Panel") = True Or String.Equals(type, "Control Panel") = True Then


                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.AddWithValue("@feature", my_Feature_code)
                cmd5.Parameters.AddWithValue("@sol", solution)
                cmd5.CommandText = "SELECT FLA from quote_table.f_dimensions where feature_code = @feature and Solution = @sol"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader


                If reader5.HasRows Then
                    While reader5.Read
                        fla_data.Rows.Add(qty, reader5(0), set_name)
                    End While
                End If

                reader5.Close()

                '==============================================================================

                'Call insert_into_mysql_dummy_Table() 'insert data to mysql dummy table
                ' scapegoat.Rows.Clear() 'clear dummy datatable
            End If

            '----------------- feature code formulas -----------------
            If String.Equals(type, "Panel") = True Then

                Dim cmd6 As New MySqlCommand
                cmd6.Parameters.AddWithValue("@feature", my_Feature_code)
                cmd6.Parameters.AddWithValue("@sol", solution)
                cmd6.CommandText = "SELECT formula from quote_table.call_feature where feature_code = @feature and solution = @sol"
                cmd6.Connection = Login.Connection
                Dim reader6 As MySqlDataReader
                reader6 = cmd6.ExecuteReader

                If reader6.HasRows Then
                    While reader6.Read
                        Call Wired_feature_codes(reader6(0).ToString, qty, set_name)
                    End While
                End If

                reader6.Close()

                Call insert_into_mysql_dummy_Table()  'insert data to mysql dummy table
                scapegoat.Rows.Clear()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub


    Function find_solution(feature_code_desc As String) As String

        'find the solution in the sol_grid.

        find_solution = ""

        For i = 0 To datatable.Rows.Count - 1
            If String.Equals(feature_code_desc, datatable.Rows(i).Item(1).ToString) = True Then
                find_solution = datatable.Rows(i).Item(2).ToString
            End If
        Next

    End Function

    Private Sub Panel_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Panel_grid.CellValueChanged
        If start_recal = True Then

            Call General_calculation()

        End If
    End Sub

    Private Sub PLC_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles PLC_grid.CellValueChanged
        If start_recal = True Then

            Call General_calculation()

        End If
    End Sub

    Private Sub Field_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Field_grid.CellValueChanged
        If start_recal = True Then

            Call General_calculation()

        End If
    End Sub

    Private Sub Control_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Control_grid.CellValueChanged
        If start_recal = True Then

            Call General_calculation()

        End If
    End Sub


    Private Sub ToolStripMenuItem7_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        mySolutions.Visible = True
    End Sub

    Private Sub Field_grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Field_grid.CellDoubleClick

        If String.IsNullOrEmpty(Field_grid.CurrentCell.Value) = False Then

            If String.Equals(Field_grid.CurrentCell.Value.ToString, "") = False And String.Equals(Field_grid.CurrentCell.Value.ToString, " ") = False Then

                Dim component As String : component = Field_grid.CurrentCell.Value.ToString.Replace("/", "-")

                If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
                    System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
                End If

            End If

        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Allocation_parts.Visible = True
    End Sub


    Sub Insert_into_dummy_table(part As String, qty As Double, set_name As String, type As String)
        'INSERT INFO TO Datatable
        If qty > 0 Then
            scapegoat.Rows.Add(part, qty, set_name, type)
        End If

    End Sub

    Sub insert_into_mysql_dummy_Table()
        'insert data from dummy_datatable to mysql dummy table
        Try
            For i = 0 To scapegoat.Rows.Count - 1

                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@parts", scapegoat.Rows(i).Item(0).ToString)
                Create_cmd.Parameters.AddWithValue("@qty", scapegoat.Rows(i).Item(1).ToString)
                Create_cmd.Parameters.AddWithValue("@Set_name", scapegoat.Rows(i).Item(2).ToString)
                Create_cmd.Parameters.AddWithValue("@Type", scapegoat.Rows(i).Item(3).ToString)
                Create_cmd.Parameters.AddWithValue("@user", current_user)
                Create_cmd.CommandText = "INSERT INTO quote_table.Dummy_quote(part_name, qty, set_name, type_p, current_u) VALUES (@parts, @qty, @Set_name, @Type, @user)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    '------------- CREATE MAIN ALLOCATION TABLE WITH VENDORS AND PRICES -----------
    Sub Create_Main_table()

        Try
            Dim cmd2 As New MySqlCommand
            cmd2.Parameters.AddWithValue("@user", current_user)
            ' cmd2.CommandText = "SELECT part_name, SUM(qty), max(type_p) from quote_table.Dummy_quote where current_u = @user group by part_name" ', type_p" OLD QUERY
            cmd2.CommandText = "SELECT part_name, qty, type_p from quote_table.Dummy_quote where current_u = @user"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_alloc_table.Rows.Add(reader2(0).ToString, "", "", "", "", reader2(1).ToString, "", "", reader2(2).ToString)
                End While
            End If

            reader2.Close()


            '--- filter tha table avoid repetitive


            'get vendor and cost
            Dim cmd3 As New MySqlCommand
            For i = 0 To my_alloc_table.Rows.Count - 1

                cmd3.Parameters.Clear()
                cmd3.Parameters.AddWithValue("@part_name", my_alloc_table.Rows(i).Item(0).ToString)
                cmd3.CommandText = "SELECT Part_Description, Manufacturer, Primary_Vendor, MFG_type, Legacy_ADA_Number from parts_table where part_name = @part_name"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        my_alloc_table.Rows(i).Item(1) = reader3(0).ToString  'part desc
                        my_alloc_table.Rows(i).Item(2) = reader3(1).ToString   'manuf
                        my_alloc_table.Rows(i).Item(3) = reader3(2).ToString   'vendor
                        my_alloc_table.Rows(i).Item(7) = reader3(3).ToString   'mfg type
                        my_alloc_table.Rows(i).Item(9) = reader3(4).ToString   'ADA
                    End While
                End If

                reader3.Close()
            Next

            'get latest cost
            For i = 0 To my_alloc_table.Rows.Count - 1
                'if part is in assembly list get total material + bulk + labor
                If my_assemblies.contains(my_alloc_table.Rows(i).Item(0)) = True Then
                    my_alloc_table.Rows(i).Item(4) = "$" & Cost_of_Assem(my_alloc_table.Rows(i).Item(0))
                Else
                    my_alloc_table.Rows(i).Item(4) = "$" & Form1.Get_Latest_Cost(Login.Connection, my_alloc_table.Rows(i).Item(0), my_alloc_table.Rows(i).Item(3))
                End If

                my_alloc_table.Rows(i).Item(6) = my_alloc_table.Rows(i).Item(4) * my_alloc_table.Rows(i).Item(5)
            Next

            If Allocation_parts.Visible = True Then
                Allocation_parts.alloc_grid.Rows.Clear()

                For i = 0 To my_alloc_table.Rows.Count - 1
                    If CType(my_alloc_table.Rows(i).Item(5).ToString, Double) > 0 Then
                        Allocation_parts.alloc_grid.Rows.Add(New String() {my_alloc_table.Rows(i).Item(0).ToString, my_alloc_table.Rows(i).Item(1).ToString, my_alloc_table.Rows(i).Item(3).ToString, my_alloc_table.Rows(i).Item(5).ToString, my_alloc_table.Rows(i).Item(4).ToString, my_alloc_table.Rows(i).Item(6).ToString, my_alloc_table.Rows(i).Item(7).ToString})
                    End If
                Next


            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

    Private Sub OpenActiveQuoteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Quote_manager.Visible = True
        Call populate_quote_manager(1)
    End Sub

    Sub populate_quote_manager(index As Integer)

        If Quote_manager.Visible = True Then

            Quote_manager.TreeView1.Nodes.Clear()

            Dim root = New TreeNode("Quotes")
            Quote_manager.TreeView1.Nodes.Add(root)

            ' Dim my_jobs = New List(Of String)()

            If index = 1 Then

                Dim cmd2 As New MySqlCommand
                '  For Each job In my_jobs
                Try
                    cmd2.CommandText = "SELECT quote_name from quote_table.my_active_quote_table"
                    cmd2.Connection = Login.Connection
                    Dim reader2 As MySqlDataReader
                    reader2 = cmd2.ExecuteReader
                    ' Quote_manager.TreeView1.Nodes(0).Nodes.Add(New TreeNode(job))

                    If reader2.HasRows Then
                        While reader2.Read
                            Quote_manager.TreeView1.Nodes(0).Nodes.Add(New TreeNode(reader2(0).ToString))
                        End While
                    End If

                    reader2.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            ElseIf index = 2 Then


            Else index = 3

            End If

        End If

    End Sub

    Private Sub NewQuoteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewQuoteToolStripMenuItem.Click
        Enter_data.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click

        Dim result As DialogResult = MessageBox.Show("Current Quote changes will be deleted. Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then
            Me.Text = "My Quote"

            '---- resetsets ---------
            For i = Panel_grid.Columns.Count - 1 To 1 Step -1
                Panel_grid.Columns.RemoveAt(i)
                PLC_grid.Columns.RemoveAt(i)
                Field_grid.Columns.RemoveAt(i)
                Control_grid.Columns.RemoveAt(i)

            Next

            counter = 1
            my_Set_info.Clear()
            current_job = ""

            Try
                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@user", current_user)
                cmd.CommandText = "delete from quote_table.Dummy_quote where current_u = @user"
                cmd.Connection = Login.Connection
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            my_alloc_table.Rows.Clear()

            Allocation_parts.Close()
            Purchase_Request.Close()
            Summary_devices.Close()

            '============= zero out install table =========
            ' cal_ins = False

            For i = 1 To Install_grid.Rows.Count - 1
                For j = 1 To Install_grid.Columns.Count - 1
                    If i <> 33 And i <> 36 And i <> 42 And j <> 2 Then
                        Install_grid.Rows(i).Cells(j).Value = 0
                    End If
                Next
            Next

            Call Fill_installation()
            '  cal_ins = True
            '----------------------------------------------
            Quote_grid.Rows.Clear()
            Call load_total_quote()

            compare_grid.Rows.Clear()
            PR_grid.Rows.Clear()
        End If
    End Sub

    Private Sub RemoveThisQuoteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveThisQuoteToolStripMenuItem.Click


        If String.Equals(Me.Text, "My Quote") = False Then
            Dim result As DialogResult = MessageBox.Show("Are you completely sure, you want to delete " & Me.Text & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                '--- first check if its a release quote, if its not then delete it fro my_active_quote_table
                Try

                    Dim Create_cmd As New MySqlCommand   'my_active_quote_table
                    Dim Create_cmd2 As New MySqlCommand  'g_setup
                    Dim Create_cmd3 As New MySqlCommand  'my_solutions
                    Dim Create_cmd4 As New MySqlCommand  'my_inputs
                    Dim Create_cmd5 As New MySqlCommand  'my_spec_orders


                    'remove data from my_active_quote_Table
                    Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                    Create_cmd.CommandText = "delete from quote_table.my_active_quote_table where quote_name = @quote_n"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Create_cmd.Parameters.Clear()

                    'remove data from g_Setup
                    Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                    Create_cmd.CommandText = "delete from quote_table.g_setup where quote_name = @quote_n"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Create_cmd.Parameters.Clear()


                    'remove data from my_inputs
                    Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                    Create_cmd.CommandText = "delete from quote_table.my_inputs where quote_name = @quote_n"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Create_cmd.Parameters.Clear()

                    'remove data from install
                    Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                    Create_cmd.CommandText = "delete from quote_table.install_table where quote_name = @quote_n"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Create_cmd.Parameters.Clear()

                    'remove data from total_q_
                    Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                    Create_cmd.CommandText = "delete from quote_table.total_q_table where quote_name = @quote_n"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Create_cmd.Parameters.Clear()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try



                Me.Text = "My Quote"

                '---- reset sets ---------
                For i = Panel_grid.Columns.Count - 1 To 1 Step -1
                    Panel_grid.Columns.RemoveAt(i)
                    PLC_grid.Columns.RemoveAt(i)
                    Field_grid.Columns.RemoveAt(i)
                    Control_grid.Columns.RemoveAt(i)
                Next

                counter = 1
                my_Set_info.Clear()
                current_job = ""

                Allocation_parts.Close()
                Purchase_Request.Close()
                Summary_devices.Close()

                '============= zero out install table =========
                For i = 1 To Install_grid.Rows.Count - 1
                    For j = 1 To Install_grid.Columns.Count - 1
                        If i <> 33 And i <> 36 And i <> 42 And j <> 2 Then
                            Install_grid.Rows(i).Cells(j).Value = 0
                        End If
                    Next
                Next

                Call Fill_installation()

                '----------------------------------------------
                Quote_grid.Rows.Clear()
                Call load_total_quote()

                Call General_calculation()

            End If

        End If
    End Sub

    Sub compute()
        '///////////////////////////       TOTAL QUOTE FORM CALCULATIONS  //////////////////////////////////
        If donotcompute = True Then

            For i = 0 To 131
                'calculations from 0 to 93 row

                If i = 78 Or i = 85 Then Continue For
                If (IsNumeric(Quote_grid.Rows(i).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(3).Value())) Then
                    Quote_grid.Rows(i).Cells(2).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(2).Value()) = False, 0, Quote_grid.Rows(i).Cells(2).Value())
                    Quote_grid.Rows(i).Cells(3).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(3).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(3).Value()) = False, 0, Quote_grid.Rows(i).Cells(3).Value())

                    Quote_grid.Rows(i).Cells(4).Value() = Quote_grid.Rows(i).Cells(2).Value() * (Quote_grid.Rows(i).Cells(3).Value() / 100 + 1)
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value())) Then
                    Quote_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = False, 0, Quote_grid.Rows(i).Cells(4).Value())
                    Quote_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = False, 0, Quote_grid.Rows(i).Cells(5).Value())

                    Quote_grid.Rows(i).Cells(6).Value() = "$" & Decimal.Round(Quote_grid.Rows(i).Cells(4).Value() * Quote_grid.Rows(i).Cells(5).Value(), 2, MidpointRounding.AwayFromZero)
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(7).Value())) Then
                    If (i < 46 Or i > 68) Then
                        Quote_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = False, 0, Quote_grid.Rows(i).Cells(5).Value())
                        Quote_grid.Rows(i).Cells(7).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(7).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(7).Value()) = False, 0, Quote_grid.Rows(i).Cells(7).Value())
                        If i <> 72 Then
                            Quote_grid.Rows(i).Cells(8).Value() = "$" & Decimal.Round(((Quote_grid.Rows(i).Cells(7).Value() / 100) + 1) * Quote_grid.Rows(i).Cells(5).Value(), 2, MidpointRounding.AwayFromZero)
                        End If
                    End If
                    End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value())) Then
                    Quote_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = False, 0, Quote_grid.Rows(i).Cells(4).Value())
                    Quote_grid.Rows(i).Cells(8).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value()) = False, 0, Quote_grid.Rows(i).Cells(8).Value())

                    Quote_grid.Rows(i).Cells(9).Value() = "$" & Decimal.Round(Quote_grid.Rows(i).Cells(4).Value() * Quote_grid.Rows(i).Cells(8).Value(), 2, MidpointRounding.AwayFromZero)
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value())) Then
                    Quote_grid.Rows(i).Cells(6).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(6).Value()) = False, 0, Quote_grid.Rows(i).Cells(6).Value())
                    Quote_grid.Rows(i).Cells(9).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(9).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(9).Value()) = False, 0, Quote_grid.Rows(i).Cells(9).Value())

                    Quote_grid.Rows(i).Cells(10).Value() = Math.Round(If(Quote_grid.Rows(i).Cells(9).Value() > 0, Math.Round(100 * ((Quote_grid.Rows(i).Cells(9).Value() - Quote_grid.Rows(i).Cells(6).Value()) / Quote_grid.Rows(i).Cells(9).Value()), 1), 0)) & "%"
                End If

            Next



        End If


    End Sub

    Private Sub Quote_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Quote_grid.CellValueChanged
        Call compute()
    End Sub



    Private Sub Panel_grid_ColumnAdded(sender As Object, e As DataGridViewColumnEventArgs) Handles Panel_grid.ColumnAdded
        'if set is added recalculate
        If start_recal = True Then
            Call General_calculation()
        End If
    End Sub

    Sub alloc_Totals()

        Dim total_material_panel As Double : total_material_panel = 0 'total material cost panel
        Dim total_material_plc As Double : total_material_plc = 0 'total material cost panel
        Dim total_material_field As Double : total_material_field = 0 'total material cost panel
        Dim total_material_a_panel As Double : total_material_a_panel = 0 'total material cost panel
        Dim total_material_m12 As Double : total_material_m12 = 0 'total material cost panel
        Dim total_material_m12e As Double : total_material_m12e = 0 'total material cost panel

        '--------------------- Allocate  total cost in quote form ------------------
        For i = 0 To my_alloc_table.Rows.Count - 1

            If String.Equals(my_alloc_table.Rows(i).Item(8), "panel") = True Or String.Equals(my_alloc_table.Rows(i).Item(8), "Panel") = True Then
                'Total starter panel material cost
                total_material_panel = total_material_panel + my_alloc_table.Rows(i).Item(6)

            ElseIf String.Equals(my_alloc_table.Rows(i).Item(8), "PLC") = True Then
                'Total plc material cost
                total_material_plc = total_material_plc + my_alloc_table.Rows(i).Item(6)

            ElseIf String.Equals(my_alloc_table.Rows(i).Item(8), "field") = True Or String.Equals(my_alloc_table.Rows(i).Item(8), "Field") = True Then
                'Total field material cost
                total_material_field = total_material_field + my_alloc_table.Rows(i).Item(6)

            ElseIf String.Equals(my_alloc_table.Rows(i).Item(8), "A-Panel") = True Or String.Equals(my_alloc_table.Rows(i).Item(8), "Control Panel") = True Then
                'Total a panel material cost
                total_material_a_panel = total_material_a_panel + my_alloc_table.Rows(i).Item(6)

            ElseIf String.Equals(my_alloc_table.Rows(i).Item(8), "M12") = True Then
                'Total M12 material cost
                total_material_m12 = total_material_m12 + my_alloc_table.Rows(i).Item(6)

            ElseIf String.Equals(my_alloc_table.Rows(i).Item(8), "M12_E") = True Then
                'Total M12-E material cost
                total_material_m12e = total_material_m12e + my_alloc_table.Rows(i).Item(6)
            End If

        Next

        '--- put the data in the quote form
        Quote_grid.Rows(1).Cells(5).Value = "$" & total_material_panel
        Quote_grid.Rows(2).Cells(5).Value = "$" & total_material_a_panel 'now control panel also
        Quote_grid.Rows(3).Cells(5).Value = "$" & total_material_plc
        Quote_grid.Rows(19).Cells(5).Value = "$" & total_material_field
        Quote_grid.Rows(21).Cells(5).Value = "$" & total_material_m12
        Quote_grid.Rows(22).Cells(5).Value = "$" & total_material_m12e
        Quote_grid.Rows(4).Cells(5).Value = "$" & labor_t  'labor cost
        Quote_grid.Rows(5).Cells(5).Value = "$" & bulk_t  'bulk cost

    End Sub

    'Private Sub GeneratePurchaseRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GeneratePurchaseRequestToolStripMenuItem.Click
    '    Purchase_Request.Visible = True
    'End Sub

    Private Sub myQuote_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Allocation_parts.Close()
        Purchase_Request.Close()
        Summary_devices.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'cal totals in quote form
        Dim total_cost As Double : total_cost = 0
        Dim final_price As Double : final_price = 0

        Dim ADA_Panels As Double : ADA_Panels = 0
        Dim ADA_Panels_f As Double : ADA_Panels_f = 0
        Dim field_item As Double : field_item = 0
        Dim field_item_f As Double : field_item_f = 0
        Dim in_office_l As Double : in_office_l = 0
        Dim in_office_lf As Double : in_office_lf = 0
        Dim onsite_l As Double : onsite_l = 0
        Dim onsite_lf As Double : onsite_lf = 0
        Dim start_up As Double : start_up = 0
        Dim start_upf As Double : start_upf = 0
        Dim electrical As Double : electrical = 0
        Dim electricalf As Double : electricalf = 0
        Dim shipping As Double : shipping = 0
        Dim shippingf As Double : shippingf = 0


        'first total sum
        For i = 1 To 15
            ADA_Panels = ADA_Panels + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            ADA_Panels_f = ADA_Panels_f + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(16).Cells(6).Value = "$" & ADA_Panels.ToString("N")
        Quote_grid.Rows(16).Cells(9).Value = "$" & ADA_Panels_f.ToString("N")

        total_cost = total_cost + ADA_Panels
        final_price = final_price + ADA_Panels_f

        'second
        For i = 19 To 42
            field_item = field_item + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            field_item_f = field_item_f + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(43).Cells(6).Value = "$" & field_item.ToString("N")
        Quote_grid.Rows(43).Cells(9).Value = "$" & field_item_f.ToString("N")


        total_cost = total_cost + field_item
        final_price = final_price + field_item_f

        'third
        For i = 45 To 56
            in_office_l = in_office_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            in_office_lf = in_office_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(57).Cells(6).Value = "$" & in_office_l.ToString("N")
        Quote_grid.Rows(57).Cells(9).Value = "$" & in_office_lf.ToString("N")


        total_cost = total_cost + in_office_l
        final_price = final_price + in_office_lf

        'fourth
        For i = 59 To 68
            onsite_l = onsite_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            onsite_lf = onsite_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(69).Cells(6).Value = "$" & onsite_l.ToString("N")
        Quote_grid.Rows(69).Cells(9).Value = "$" & onsite_lf.ToString("N")

        total_cost = total_cost + onsite_l
        final_price = final_price + onsite_lf

        'fifth
        For i = 72 To 75
            start_up = start_up + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            start_upf = start_upf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(76).Cells(6).Value = "$" & start_up.ToString("N")
        Quote_grid.Rows(76).Cells(9).Value = "$" & start_upf.ToString("N")

        total_cost = total_cost + start_up
        final_price = final_price + start_upf

        'sixth
        For i = 79 To 82
            electrical = electrical + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            electricalf = electricalf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(83).Cells(6).Value = "$" & electrical.ToString("N")
        Quote_grid.Rows(83).Cells(9).Value = "$" & electricalf.ToString("N")

        total_cost = total_cost + electrical
        final_price = final_price + electricalf

        'seventh
        For i = 86 To 88
            shipping = shipping + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            shippingf = shippingf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(89).Cells(6).Value = "$" & shipping.ToString("N")
        Quote_grid.Rows(89).Cells(9).Value = "$" & shippingf.ToString("N")

        total_cost = total_cost + shipping
        final_price = final_price + shippingf

        Quote_grid.Rows(101).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(101).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(107).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(107).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(123).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(123).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(130).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(130).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")


        '--- cal expenses
        Dim days_site As Double : days_site = 0

        For i = 60 To 68
            days_site = days_site + If(IsNumeric(Quote_grid.Rows(i).Cells(2).Value) = True, Quote_grid.Rows(i).Cells(2).Value, 0)
        Next

        Dim on_site As Double : on_site = 1

        If IsNumeric(Quote_grid.Rows(138).Cells(2).Value) = True Then
            If Quote_grid.Rows(138).Cells(2).Value <= 0 Then
                on_site = 1
            Else
                on_site = Quote_grid.Rows(138).Cells(2).Value  'hours / day onsite
            End If
        End If

        Quote_grid.Rows(146).Cells(2).Value = Math.Ceiling(days_site / on_site)  'DAYS ON SITE
        '--------------------------

        '---- cal exp
        If (IsNumeric(Quote_grid.Rows(146).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(139).Cells(2).Value())) Then
            Quote_grid.Rows(146).Cells(2).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(146).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(146).Cells(2).Value()) = False, 0, Quote_grid.Rows(146).Cells(2).Value())
            Quote_grid.Rows(139).Cells(2).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(139).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(139).Cells(2).Value()) = False, 0, Quote_grid.Rows(139).Cells(2).Value())

            If Quote_grid.Rows(139).Cells(2).Value() > 0 Then
                Quote_grid.Rows(147).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value())  '#trips
                Quote_grid.Rows(151).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value())  '#Airfare
                Quote_grid.Rows(148).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value()) * If(String.IsNullOrEmpty(Quote_grid.Rows(140).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(140).Cells(2).Value()) = False, 0, Quote_grid.Rows(140).Cells(2).Value()) 'travel time
                Quote_grid.Rows(149).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value()) * If(String.IsNullOrEmpty(Quote_grid.Rows(141).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(141).Cells(2).Value()) = False, 0, Quote_grid.Rows(141).Cells(2).Value()) 'mileage
                Quote_grid.Rows(150).Cells(2).Value() = (Quote_grid.Rows(147).Cells(2).Value() * If(String.IsNullOrEmpty(Quote_grid.Rows(142).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(142).Cells(2).Value()) = False, 0, Quote_grid.Rows(142).Cells(2).Value())) + Quote_grid.Rows(146).Cells(2).Value

                Quote_grid.Rows(72).Cells(2).Value() = Quote_grid.Rows(148).Cells(2).Value()
                Quote_grid.Rows(73).Cells(2).Value() = Quote_grid.Rows(149).Cells(2).Value()
                Quote_grid.Rows(74).Cells(2).Value() = Quote_grid.Rows(150).Cells(2).Value()
                Quote_grid.Rows(75).Cells(2).Value() = Quote_grid.Rows(151).Cells(2).Value()


            End If
        End If

        '----((((( TOTALS - By component
        '--ADA Panels
        Quote_grid.Rows(94).Cells(6).Value = "$" & ADA_Panels
        Quote_grid.Rows(94).Cells(9).Value = "$" & ADA_Panels_f

        '--- Field items
        Quote_grid.Rows(95).Cells(6).Value = "$" & field_item
        Quote_grid.Rows(95).Cells(9).Value = "$" & field_item_f

        '--- In office labor
        Quote_grid.Rows(96).Cells(6).Value = "$" & in_office_l
        Quote_grid.Rows(96).Cells(9).Value = "$" & in_office_lf

        '---- onsite labor
        Quote_grid.Rows(97).Cells(6).Value = "$" & onsite_l
        Quote_grid.Rows(97).Cells(9).Value = "$" & onsite_lf

        '--- start up expenses
        Quote_grid.Rows(98).Cells(6).Value = "$" & start_up
        Quote_grid.Rows(98).Cells(9).Value = "$" & start_upf

        '-------electrical installation
        Quote_grid.Rows(99).Cells(6).Value = "$" & electrical
        Quote_grid.Rows(99).Cells(9).Value = "$" & electricalf

        '----- shipping -------
        Quote_grid.Rows(100).Cells(6).Value = "$" & shipping
        Quote_grid.Rows(100).Cells(9).Value = "$" & shippingf

        '----(((((( TOTALS Sum all but installation

        '-------- hardware and labor ---
        Quote_grid.Rows(104).Cells(6).Value = "$" & ADA_Panels + field_item + in_office_l + onsite_l + start_up
        Quote_grid.Rows(104).Cells(9).Value = "$" & ADA_Panels_f + field_item_f + in_office_lf + onsite_lf + start_upf

        '-------- electrical installation --
        Quote_grid.Rows(105).Cells(6).Value = "$" & electrical
        Quote_grid.Rows(105).Cells(9).Value = "$" & electricalf

        '--------- shipping ------
        Quote_grid.Rows(106).Cells(6).Value = "$" & shipping
        Quote_grid.Rows(106).Cells(9).Value = "$" & shippingf

        '---- (((( TOTAL-- By type

        '--------- Project labor 
        Dim project_l As Double : project_l = 0
        Dim project_lf As Double : project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(110).Cells(6).Value = project_l
        Quote_grid.Rows(110).Cells(9).Value = project_lf

        '------- Project Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(111).Cells(6).Value = project_l
        Quote_grid.Rows(111).Cells(9).Value = project_lf

        '------- Project Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(112).Cells(6).Value = project_l
        Quote_grid.Rows(112).Cells(9).Value = project_lf

        '--------- Manufacturing Labor

        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Manufacturing Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(113).Cells(6).Value = project_l
        Quote_grid.Rows(113).Cells(9).Value = project_lf



        '--------- Manufacturing Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Manufacturing Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(114).Cells(6).Value = project_l
        Quote_grid.Rows(114).Cells(9).Value = project_lf

        '--------- Startup Labor
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(115).Cells(6).Value = project_l
        Quote_grid.Rows(115).Cells(9).Value = project_lf

        '--------- Startup Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(116).Cells(6).Value = project_l
        Quote_grid.Rows(116).Cells(9).Value = project_lf

        '--------- Startup Expenses
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Expenses") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(117).Cells(6).Value = project_l
        Quote_grid.Rows(117).Cells(9).Value = project_lf

        '--------- Startup Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(118).Cells(6).Value = project_l
        Quote_grid.Rows(118).Cells(9).Value = project_lf

        '--------- Install Labor
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(119).Cells(6).Value = project_l
        Quote_grid.Rows(119).Cells(9).Value = project_lf

        '--------- Install Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(120).Cells(6).Value = project_l
        Quote_grid.Rows(120).Cells(9).Value = project_lf

        '--------- Install expenses
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Expenses") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(121).Cells(6).Value = project_l
        Quote_grid.Rows(121).Cells(9).Value = project_lf

        '--------- Install Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(122).Cells(6).Value = project_l
        Quote_grid.Rows(122).Cells(9).Value = project_lf


        '--- (((( TOTALS Summary ------
        Quote_grid.Rows(126).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(110).Cells(6).Value), Quote_grid.Rows(110).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(113).Cells(6).Value), Quote_grid.Rows(113).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(115).Cells(6).Value), Quote_grid.Rows(115).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(119).Cells(6).Value), Quote_grid.Rows(119).Cells(6).Value, 0))
        Quote_grid.Rows(126).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(110).Cells(9).Value), Quote_grid.Rows(110).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(113).Cells(9).Value), Quote_grid.Rows(113).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(115).Cells(9).Value), Quote_grid.Rows(115).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(119).Cells(9).Value), Quote_grid.Rows(119).Cells(9).Value, 0))

        Quote_grid.Rows(127).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(111).Cells(6).Value), Quote_grid.Rows(111).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(116).Cells(6).Value), Quote_grid.Rows(116).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(120).Cells(6).Value), Quote_grid.Rows(120).Cells(6).Value, 0))
        Quote_grid.Rows(127).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(111).Cells(9).Value), Quote_grid.Rows(111).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(116).Cells(9).Value), Quote_grid.Rows(116).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(120).Cells(9).Value), Quote_grid.Rows(120).Cells(9).Value, 0))

        Quote_grid.Rows(128).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(117).Cells(6).Value), Quote_grid.Rows(117).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(121).Cells(6).Value), Quote_grid.Rows(121).Cells(6).Value, 0))
        Quote_grid.Rows(128).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(117).Cells(9).Value), Quote_grid.Rows(117).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(121).Cells(9).Value), Quote_grid.Rows(121).Cells(9).Value, 0))

        Quote_grid.Rows(129).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(112).Cells(6).Value), Quote_grid.Rows(112).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(114).Cells(6).Value), Quote_grid.Rows(114).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(118).Cells(6).Value), Quote_grid.Rows(118).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(122).Cells(6).Value), Quote_grid.Rows(122).Cells(6).Value, 0))
        Quote_grid.Rows(129).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(112).Cells(9).Value), Quote_grid.Rows(112).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(114).Cells(9).Value), Quote_grid.Rows(114).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(118).Cells(9).Value), Quote_grid.Rows(118).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(122).Cells(9).Value), Quote_grid.Rows(122).Cells(9).Value, 0))


        For z = 110 To 122
            Quote_grid.Rows(z).Cells(6).Value = "$" & CType(Quote_grid.Rows(z).Cells(6).Value, Double).ToString("N")
            Quote_grid.Rows(z).Cells(9).Value = "$" & CType(Quote_grid.Rows(z).Cells(9).Value, Double).ToString("N")
        Next

    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Summary_devices.Visible = True
    End Sub



    Private Sub apanel_box_SelectedIndexChanged(sender As Object, e As EventArgs) Handles apanel_box.SelectedIndexChanged

        If start_recal = True Then

            Dim result As Boolean


            If change_sol = False Then
                Dim result_temp As DialogResult = MessageBox.Show("Current Quote changes will be deleted. Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                If result_temp = DialogResult.Yes Then
                    result = True
                Else
                    result = False
                End If

            Else
                result = True
            End If



            If (result = True) Then

                sol_label.Text = apanel_box.Text

                Dim vfd_type As String : vfd_type = "AB"

                If PF525_c.Checked = True Then
                    vfd_type = "AB"
                Else
                    vfd_type = "Eaton"
                End If

                Panel_grid.Rows.Clear()
                PLC_grid.Rows.Clear()
                Control_grid.Rows.Clear()

                Dim list_panel = New List(Of String)()
                Dim list_plc = New List(Of String)()
                Dim list_control = New List(Of String)()

                Dim c_1 As Integer : c_1 = 0
                Dim c_2 As Integer : c_2 = 0
                Dim c_3 As Integer : c_3 = 0

                Try
                    '-------- Panel Specific type -------------
                    Dim cmd_panel As New MySqlCommand
                    cmd_panel.Parameters.AddWithValue("@sol", sol_label.Text)
                    cmd_panel.CommandText = "SELECT distinct specific_type, type from quote_table.feature_codes where show_menu = 'Y' and Solution = @sol"
                    cmd_panel.Connection = Login.Connection
                    Dim reader_panel As MySqlDataReader
                    reader_panel = cmd_panel.ExecuteReader

                    If reader_panel.HasRows Then
                        While reader_panel.Read
                            If String.Equals(reader_panel(1).ToString, "Panel") = True Then
                                list_panel.Add(reader_panel(0).ToString)
                            ElseIf String.Equals(reader_panel(1).ToString, "PLC") = True Then
                                list_plc.Add(reader_panel(0).ToString)
                            ElseIf String.Equals(reader_panel(1).ToString, "Control Panel") = True Then
                                list_control.Add(reader_panel(0).ToString)
                            End If
                        End While
                    End If

                    reader_panel.Close()


                    Dim cmd_dt As New MySqlCommand
                    Dim reader_panel2 As MySqlDataReader
                    Dim reader_plc As MySqlDataReader
                    Dim reader_control As MySqlDataReader

                    '---------------- Panel -----------------
                    For j = 0 To list_panel.Count - 1

                        Panel_grid.Rows.Add(New String() {list_panel(j)})  'add type line
                        Panel_grid.Rows(c_1).DefaultCellStyle.BackColor = Color.Goldenrod
                        Panel_grid.Rows(c_1).ReadOnly = True

                        cmd_dt.Parameters.Clear()
                        cmd_dt.Parameters.AddWithValue("@s_type", list_panel(j))
                        cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                        cmd_dt.Parameters.AddWithValue("@vfd", vfd_type)
                        cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = @vfd or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'panel' and Solution = @sol"
                        cmd_dt.Connection = Login.Connection
                        reader_panel2 = cmd_dt.ExecuteReader

                        If reader_panel2.HasRows Then
                            While reader_panel2.Read
                                Panel_grid.Rows.Add(New String() {reader_panel2(0).ToString})
                                c_1 = c_1 + 1
                            End While
                        End If
                        reader_panel2.Close()

                        c_1 = c_1 + 1
                    Next

                    '-------------- PLC ------------
                    For j = 0 To list_plc.Count - 1

                        PLC_grid.Rows.Add(New String() {list_plc(j)})  'add type line
                        PLC_grid.Rows(c_2).DefaultCellStyle.BackColor = Color.Goldenrod
                        PLC_grid.Rows(c_2).ReadOnly = True

                        cmd_dt.Parameters.Clear()
                        cmd_dt.Parameters.AddWithValue("@s_type", list_plc(j))
                        cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                        cmd_dt.Parameters.AddWithValue("@vfd", vfd_type)
                        cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = @vfd or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'PLC' and Solution = @sol"
                        cmd_dt.Connection = Login.Connection
                        reader_plc = cmd_dt.ExecuteReader

                        If reader_plc.HasRows Then
                            While reader_plc.Read
                                PLC_grid.Rows.Add(New String() {reader_plc(0).ToString})
                                c_2 = c_2 + 1
                            End While
                        End If
                        reader_plc.Close()

                        c_2 = c_2 + 1
                    Next

                    '------------- Control Panel -----------------
                    For j = 0 To list_control.Count - 1

                        Control_grid.Rows.Add(New String() {list_control(j)})  'add type line
                        Control_grid.Rows(c_3).DefaultCellStyle.BackColor = Color.Goldenrod
                        Control_grid.Rows(c_3).ReadOnly = True

                        cmd_dt.Parameters.Clear()
                        cmd_dt.Parameters.AddWithValue("@s_type", list_control(j))
                        cmd_dt.Parameters.AddWithValue("@sol", sol_label.Text)
                        cmd_dt.Parameters.AddWithValue("@vfd", vfd_type)
                        cmd_dt.CommandText = "SELECT description from quote_table.feature_codes where (VFD_TYPE = @vfd or VFD_TYPE = 'none') and show_menu = 'Y' and specific_type = @s_type and type = 'Control Panel' and Solution = @sol"
                        cmd_dt.Connection = Login.Connection
                        reader_control = cmd_dt.ExecuteReader

                        If reader_control.HasRows Then
                            While reader_control.Read
                                Control_grid.Rows.Add(New String() {reader_control(0).ToString})
                                c_3 = c_3 + 1
                            End While
                        End If
                        reader_control.Close()

                        c_3 = c_3 + 1
                    Next


                    If Allocation_parts.Visible = True Then
                        Allocation_parts.alloc_grid.Rows.Clear()
                    End If

                    ' sol_label.Text = apanel_box.Text
                    Call General_calculation()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If

        End If
    End Sub

    Private Sub IncludeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Include_A.Click
        Include_A.Checked = Not Include_A.Checked
        Call General_calculation()
    End Sub

    Sub M12_process()

        'Calculate M12 and M12E cables
        'calculate the number of M12 cables we need
        Dim inputs As Integer : inputs = 0
        Dim outputs As Integer : outputs = 0
        Dim motion As Integer : motion = 0

        Dim limit_feet As Double : limit_feet = 15
        avoid_change_m12 = False

        Dim n_rows As Double : n_rows = 0
        Dim ran_feet As Integer : ran_feet = 3
        M12_grid.Rows.Clear()
        M12_ES_grid.Rows.Clear()

        'calculate total inputs

        For i = 1 To Field_grid.Columns.Count - 1
            inputs = inputs + Count_IO(Field_grid.Columns(i).HeaderText, 1)
        Next
        inputs_io = inputs

        'calculate total outputs

        For i = 1 To Field_grid.Columns.Count - 1
            outputs = outputs + Count_IO(Field_grid.Columns(i).HeaderText, 2)
        Next

        outputs_io = outputs

        'calculate total motion

        For i = 1 To Field_grid.Columns.Count - 1
            motion = motion + Count_IO(Field_grid.Columns(i).HeaderText, 3)
        Next
        motion_io = motion

        n_rows = inputs + outputs + motion
        n_rows = n_rows + n_rows * (If(IsNumeric(percentage_io.Text) = True, percentage_io.Text, 0) / 100)

        'write io labels
        total_io_l.Text = "Total IO: " & n_rows
        inputs_io_l.Text = "Digital Inputs: " & inputs
        output_io_l.Text = "Digital Outputs: " & outputs
        motion_io_l.Text = "Motion Outputs: " & motion

        If n_rows > 0 Then
            If RIO_press = False Then
                rio_label_w.Text = "Please DO NOT forget to allocate RIO Modules"
            Else
                rio_label_w.Text = ""
            End If
        End If

        For i = 0 To n_rows - 1
            If ran_feet > limit_feet Then
                ran_feet = 3
            End If
            M12_grid.Rows.Add(New String() {ran_feet})
            ran_feet = ran_feet + 3
        Next

        Dim fixed_l(5) As Double  'this will contain the qty of cable M12
        fixed_l(0) = 0
        fixed_l(1) = 0
        fixed_l(2) = 0
        fixed_l(3) = 0
        fixed_l(4) = 0
        fixed_l(5) = 0



        Dim temp As Double : temp = 0
        Dim temp_l(5) As Double 'this array is like fixed_l but doesnt accumulate value

        For i = 0 To M12_grid.Rows.Count - 1

            'clear temp_l
            temp_l(0) = 0 : temp_l(1) = 0 : temp_l(2) = 0 : temp_l(3) = 0 : temp_l(4) = 0 : temp_l(5) = 0


            fixed_l(5) = fixed_l(5) + Math.Round(M12_grid.Rows(i).Cells(0).Value / 65.6)
            temp_l(5) = Math.Round(M12_grid.Rows(i).Cells(0).Value / 65.6)

            temp = M12_grid.Rows(i).Cells(0).Value Mod 65.6

            If temp <= 6.56 Then
                fixed_l(0) = fixed_l(0) + 1
                temp_l(0) = 1
            End If

            If temp <= 9.84 And temp > 6.56 Then
                fixed_l(1) = fixed_l(1) + 1
                temp_l(1) = 1
            End If

            If temp <= 16.4 And temp > 9.84 Then
                fixed_l(2) = fixed_l(2) + 1
                temp_l(2) = 1
            End If

            If temp <= 32.8 And temp > 16.4 Then
                fixed_l(3) = fixed_l(3) + 1
                temp_l(3) = 1
            End If

            If temp <= 49.2 And temp > 32.8 Then
                fixed_l(4) = fixed_l(4) + 1
                temp_l(4) = 1
            End If

            M12_grid.Rows(i).Cells(1).Value = Math.Round(temp_l(0) * 6.56 + temp_l(1) * 9.84 + temp_l(2) * 16.4 + temp_l(3) * 32.8 + temp_l(4) * 49.2 + temp_l(5) * 65.6, 1)

        Next

        For j = 0 To 5
            fixed_l(j) = Math.Ceiling(fixed_l(j) * 1.1)
        Next

        '------------ M12_ES---------------
        Dim ran_feet_es As Double : ran_feet_es = 3
        Dim count_rows As Double : count_rows = 0


        count_rows = estops_count()

        For i = 0 To count_rows - 1
            If ran_feet_es > 55 Then
                ran_feet_es = 3
            End If
            M12_ES_grid.Rows.Add(New String() {ran_feet_es})

            If ran_feet_es > 44 Then
                ran_feet_es = ran_feet_es + 5
            Else
                ran_feet_es = ran_feet_es + 3
            End If
        Next

        Dim fixed_l2(5) As Double  'this will contain the qty of cable M12
        fixed_l2(0) = 0
        fixed_l2(1) = 0
        fixed_l2(2) = 0
        fixed_l2(3) = 0
        fixed_l2(4) = 0
        fixed_l2(5) = 0



        Dim temp2 As Double : temp2 = 0
        Dim temp_l2(5) As Double 'this array is like fixed_l but doesnt accumulate value

        For i = 0 To M12_ES_grid.Rows.Count - 1

            'clear temp_l
            temp_l2(0) = 0 : temp_l2(1) = 0 : temp_l2(2) = 0 : temp_l2(3) = 0 : temp_l2(4) = 0 : temp_l2(5) = 0


            fixed_l2(5) = fixed_l2(5) + Math.Round(M12_ES_grid.Rows(i).Cells(0).Value / 65.6)
            temp_l2(5) = Math.Round(M12_ES_grid.Rows(i).Cells(0).Value / 65.6)

            temp2 = M12_ES_grid.Rows(i).Cells(0).Value Mod 65.6

            If temp2 <= 6.56 Then
                fixed_l2(0) = fixed_l2(0) + 1
                temp_l2(0) = 1
            End If

            If temp2 <= 9.84 And temp2 > 6.56 Then
                fixed_l2(1) = fixed_l2(1) + 1  '0 index
                temp_l2(1) = 1   '0 index
            End If

            If temp2 <= 16.4 And temp2 > 9.84 Then
                fixed_l2(2) = fixed_l2(2) + 1
                temp_l2(2) = 1
            End If

            If temp2 <= 32.8 And temp2 > 16.4 Then
                fixed_l2(3) = fixed_l2(3) + 1
                temp_l2(3) = 1
            End If

            If temp2 <= 49.2 And temp2 > 32.8 Then
                fixed_l2(4) = fixed_l2(4) + 1
                temp_l2(4) = 1
            End If

            M12_ES_grid.Rows(i).Cells(1).Value = Math.Round(temp_l2(0) * 6.56 + temp_l2(1) * 9.84 + temp_l2(2) * 16.4 + temp_l2(3) * 32.8 + temp_l2(4) * 49.2 + temp_l2(5) * 65.6, 1)

        Next

        For j = 0 To 5
            fixed_l2(j) = Math.Ceiling(fixed_l2(j) * 1.1)
        Next

        '////////////////////   fill M12 Cables list  //////////


        M12_cables("V15-G-2M-PUR-V15-G") = fixed_l(0) + fixed_l(1)
        ' M12_cables("V15-G-3M-PUR-V15-G") = fixed_l(1)
        M12_cables("V15-G-5M-PUR-V15-G") = fixed_l(2)
        M12_cables("V15-G-10M-PUR-V15-G") = fixed_l(3)
        M12_cables("V15-G-15M-PUR-V15-G") = fixed_l(4)
        M12_cables("V15-G-20M-PUR-V15-G") = fixed_l(5)

        M12_ES_cables("V15-G-S-YE2M-PUR-A-V15-G") = fixed_l2(0)
        M12_ES_cables("7000-40041-0150300") = fixed_l2(1)
        M12_ES_cables("V15-G-S-YE5M-PUR-A-V15-G") = fixed_l2(2)
        M12_ES_cables("V15-G-S-YE10M-PUR-A-V15-G") = fixed_l2(3)
        M12_ES_cables("V15-G-S-YE15M-PUR-A-V15-G") = fixed_l2(4)
        M12_ES_cables("V15-G-S-YE20M-PUR-A-V15-G") = fixed_l2(5)

        'insert m12 into dummy datatable
        scapegoat.Rows.Clear()

        For Each kvp As KeyValuePair(Of String, Double) In M12_cables.ToArray
            If CType(M12_cables(kvp.Key), Double) > 0 And Panel_grid.Columns.Count > 1 Then
                scapegoat.Rows.Add(kvp.Key, M12_cables(kvp.Key), Panel_grid.Columns.Item(1).HeaderText, "M12")
            End If
        Next

        Call insert_into_mysql_dummy_Table()
        scapegoat.Rows.Clear()

        'insert m12 into dummy datatable
        For Each kvp As KeyValuePair(Of String, Double) In M12_ES_cables.ToArray
            If CType(M12_ES_cables(kvp.Key), Double) > 0 And Panel_grid.Columns.Count > 1 Then
                scapegoat.Rows.Add(kvp.Key, M12_ES_cables(kvp.Key), Panel_grid.Columns.Item(1).HeaderText, "M12_E")
            End If
        Next

        Call insert_into_mysql_dummy_Table()
        scapegoat.Rows.Clear()

        avoid_change_m12 = True

    End Sub

    Sub RIO_analyze()

        RIO_grid.Rows.Clear()
        rio_grid2.Rows.Clear()

        RIO_grid.Rows.Add(New String() {"Required", inputs_io, outputs_io, motion_io})
        RIO_grid.Rows.Add(New String() {"Available", 0, 0, 0})


        Dim f_code = From feature_c In table_remote_io Where feature_c.Field(Of String)("solutiom") = "EIP-MS/EIP-RIO" Select feature_c

        For Each row As DataRow In f_code
            rio_grid2.Rows.Add(New String() {row.Item("feature_code")})
        Next


        For i = 0 To rio_grid2.Rows.Count - 1
            Call Rio_needed(i)
        Next

    End Sub

    '--------- this function belonged to RIO wiz -----
    Sub Rio_needed(index As Integer)
        'calculate the needed and available RIO

        Dim n_inputs As Double : n_inputs = 0
        Dim n_outputs As Double : n_outputs = 0
        Dim n_motion As Double : n_motion = 0

        Dim iom = From iom_t In table_remote_io Where iom_t.Field(Of String)("solutiom") = "EIP-MS/EIP-RIO" And iom_t.Field(Of String)("feature_code") = rio_grid2.Rows(index).Cells(0).Value Select iom_t

        For Each row As DataRow In iom
            If row.Item("inputs") > 0 Then
                n_inputs = Math.Ceiling(inputs_io / row.Item("inputs"))
            End If

            If row.Item("output") > 0 Then
                n_outputs = Math.Ceiling(outputs_io / row.Item("output"))
            End If

            If row.Item("motion") > 0 Then
                n_motion = Math.Ceiling(motion_io / row.Item("motion"))
            End If
        Next

        rio_grid2.Rows(index).Cells(1).Value = Max_number(n_inputs, n_outputs, n_motion)
        rio_grid2.Rows(index).Cells(2).Value = Max_number(n_inputs, n_outputs, n_motion)

        'Try
        '    Dim cmd_panel As New MySqlCommand
        '    cmd_panel.Parameters.AddWithValue("@sol", "EIP-MS/EIP-RIO")
        '    cmd_panel.Parameters.AddWithValue("@feature", rio_grid2.Rows(index).Cells(0).Value)
        '    cmd_panel.CommandText = "SELECT inputs, output, motion from quote_table.Remote_IO where solutiom = @sol and feature_code = @feature"
        '    cmd_panel.Connection = Login.Connection
        '    Dim reader_panel As MySqlDataReader
        '    reader_panel = cmd_panel.ExecuteReader

        '    If reader_panel.HasRows Then
        '        While reader_panel.Read

        '            If reader_panel(0) > 0 Then
        '                n_inputs = Math.Ceiling(inputs_io / reader_panel(0))
        '            End If

        '            If reader_panel(1) > 0 Then
        '                n_outputs = Math.Ceiling(outputs_io / reader_panel(1))
        '            End If

        '            If reader_panel(2) > 0 Then
        '                n_motion = Math.Ceiling(motion_io / reader_panel(2))
        '            End If

        '        End While
        '    End If

        '    rio_grid2.Rows(index).Cells(1).Value = Max_number(n_inputs, n_outputs, n_motion)
        '    rio_grid2.Rows(index).Cells(2).Value = Max_number(n_inputs, n_outputs, n_motion)

        '    reader_panel.Close()

        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try


    End Sub

    Function Count_IO(set_m As String, io As Integer) As Integer

        'return the number of IO
        '
        ' 1 -- input
        ' 2 -- output
        ' 3 -- motion


        Dim t_io As Double : t_io = 0
        Dim index As Integer : index = 1

        'find column
        For i = 1 To Field_grid.Columns.Count - 1
            If String.Equals(Field_grid.Columns(i).HeaderText, set_m) = True Then
                index = i
                Exit For
            End If
        Next

        For i = 0 To Field_grid.Rows.Count - 1

            If io = 1 Then

                Dim desc As String : desc = Field_grid.Rows(i).Cells(0).Value
                Dim inputs_t = From inp In table_io_points Where inp.Field(Of String)("solution") = sol_label.Text And inp.Field(Of String)("description") = desc Select inp

                For Each row As DataRow In inputs_t
                    If IsNumeric(Field_grid.Rows(i).Cells(index).Value) = True Then
                        t_io = t_io + (Field_grid.Rows(i).Cells(index).Value * (row.Item("input")))
                    End If
                Next


            ElseIf io = 2 Then

                Dim desc As String : desc = Field_grid.Rows(i).Cells(0).Value
                Dim output_t = From inp In table_io_points Where inp.Field(Of String)("solution") = sol_label.Text And inp.Field(Of String)("description") = desc Select inp

                For Each row As DataRow In output_t
                    If IsNumeric(Field_grid.Rows(i).Cells(index).Value) = True Then
                        t_io = t_io + (Field_grid.Rows(i).Cells(index).Value * (row.Item("output")))
                    End If
                Next


            Else

                Dim desc As String : desc = Field_grid.Rows(i).Cells(0).Value
                Dim motion_t = From inp In table_io_points Where inp.Field(Of String)("solution") = sol_label.Text And inp.Field(Of String)("description") = desc Select inp

                For Each row As DataRow In motion_t
                    If IsNumeric(Field_grid.Rows(i).Cells(index).Value) = True Then
                        t_io = t_io + (Field_grid.Rows(i).Cells(index).Value * (row.Item("motion")))
                    End If
                Next

            End If
        Next

        Count_IO = t_io

    End Function

    Function estops_count() As Integer

        Dim estop As Integer : estop = 0

        For i = 0 To Field_grid.Rows.Count - 1
            If String.Equals(Field_grid.Rows(i).Cells(0).Value, "ADA-ASM-EPB101") = True Or String.Equals(Field_grid.Rows(i).Cells(0).Value, "ADA-ASM-EPB102") = True Or String.Equals(Field_grid.Rows(i).Cells(0).Value, "ADA-ASM-EPC101") = True Or String.Equals(Field_grid.Rows(i).Cells(0).Value, "ADA-ASM-EPC102") = True Then
                For j = 1 To Field_grid.Columns.Count - 1
                    estop = estop + If(IsNumeric(Field_grid.Rows(i).Cells(j).Value) = True, Field_grid.Rows(i).Cells(j).Value, 0)
                Next
            End If
        Next

        estops_count = estop

    End Function

    Private Sub EnableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles recal_enable.Click
        '------------ Enables recalculation ------------
        recal_enable.Checked = Not recal_enable.Checked

        If recal_enable.Checked = True Then
            go_on = False


            Call General_calculation()

        Else
            go_on = True
        End If


    End Sub

    Sub Wired_feature_codes(feature_code As String, qty As Integer, set_n As String)

        'These are the hardwire feature codes that process the qty value and are called by other feature code
        Dim myqty As Double : myqty = 0

        Select Case feature_code

            Case "ES-BUS"
                'process qty
                myqty = Math.Ceiling(qty / 3)

                Call Insert_into_dummy_table("FSFD 4.5-0.5", myqty, set_n, "Panel")  'insert data to dummy datatable
                Call Insert_into_dummy_table("FKFD 4.5-0.5", myqty, set_n, "panel")
                Call Insert_into_dummy_table("3210567", 2 * myqty, set_n, "panel")
                Call Insert_into_dummy_table("3211634", myqty, set_n, "panel")


            Case "ES-BUS-P"
                'process qty
                myqty = Panel_grid.Columns.Count - 1

                Call Insert_into_dummy_table("FSFD 4.5-0.5", myqty, set_n, "panel")  'insert data to dummy datatable
                Call Insert_into_dummy_table("FKFD 4.5-0.5", myqty, set_n, "panel")
                Call Insert_into_dummy_table("3210567", 2 * myqty, set_n, "panel")
                Call Insert_into_dummy_table("3211634", myqty, set_n, "panel")


            Case "SafeRly"

                'process qty
                myqty = Math.Ceiling(qty / 3)

                Call Insert_into_dummy_table("P7SA-10F", myqty, set_n, "panel")  'insert data to dummy datatable
                Call Insert_into_dummy_table("G7SA-3A1B-DC24", myqty, set_n, "panel")


            Case "MS-4-DE1"
                'process qty
                myqty = Math.Ceiling(qty / 3)
                Call Insert_into_dummy_table("XTSC010BBTB", myqty, set_n, "panel")

            Case "MDR_PS"
                Dim qty_40 As Double : qty_40 = Math.Ceiling(qty / 40)
                Dim qty_20 As Double : qty_20 = Math.Ceiling((qty - (qty_40 * 40)) / 20)

                Call Insert_into_dummy_table("QT20.241", qty_20, set_n, "panel")
                Call Insert_into_dummy_table("FAZ-C20/1-SP", qty_20, set_n, "panel")
                Call Insert_into_dummy_table("3211813", qty_20, set_n, "panel")
                Call Insert_into_dummy_table("XTSC010BBTB", Math.Ceiling(qty_20 / 6), set_n, "panel")

                Call Insert_into_dummy_table("QT20.241", qty_40, set_n, "panel")
                Call Insert_into_dummy_table("FAZ-C20/1-SP", 2 * qty_40, set_n, "panel")
                Call Insert_into_dummy_table("3211813", 2 * qty_40, set_n, "panel")
                Call Insert_into_dummy_table("XTSC010BBTB", Math.Ceiling(qty_20 / 4), set_n, "panel")

            Case "Class2"

                Dim qty_4 As Double : qty_4 = ((qty + 3) / 4) - 2 * ((qty + 3) / 8)
                Dim qty_8 As Double : qty_8 = ((qty + 3) / 8)

                '--channel 4

                Call Insert_into_dummy_table("QT20.241", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("9000-41064-0400000", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("3000715", qty_4 * 2, set_n, "panel")
                Call Insert_into_dummy_table("3211634", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("3030161", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("3211813", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("3048519", qty_4, set_n, "panel")
                Call Insert_into_dummy_table("ATDR3 600V CC TD", 3 * qty_4, set_n, "panel")
                Call Insert_into_dummy_table("FKFD 4.5-0.5", 4 * qty_4, set_n, "panel")

                '--channel 8

                Call Insert_into_dummy_table("QT20.241", qty_8, set_n, "panel")
                Call Insert_into_dummy_table("9000-41064-0400000", qty_8, set_n, "panel")
                Call Insert_into_dummy_table("3000715", qty_8 * 4, set_n, "panel")
                Call Insert_into_dummy_table("3030187", qty_8, set_n, "panel")
                Call Insert_into_dummy_table("3211634", qty_8, set_n, "panel")
                Call Insert_into_dummy_table("3211813", qty_8 * 2, set_n, "panel")
                Call Insert_into_dummy_table("3048519", qty_8 * 2, set_n, "panel")
                Call Insert_into_dummy_table("3ATDR3 600V CC TD", qty_8 * 3, set_n, "panel")
                Call Insert_into_dummy_table("FKFD 4.5-0.5", qty_8 * 8, set_n, "panel")
        End Select



    End Sub

    'Function Mode_panel(set_name As String) As String
    '    'return the mode of the set

    '    Mode_panel = "All 24' Panels"

    '    For i = 0 To dimen_table.Rows.Count - 1
    '        If String.Equals(set_name, dimen_table.Rows(i).Item(0).ToString) = True Then
    '            Mode_panel = dimen_table.Rows(i).Item(1).ToString
    '            Exit For
    '        End If
    '    Next

    'End Function

    'Function Mypanel_qty(set_name As String, mode As String) As Double
    '    'return new panel qty (24 or 30 inch)

    '    Mypanel_qty = 0

    '    For i = 0 To dimen_table.Rows.Count - 1
    '        If String.Equals(set_name, dimen_table.Rows(i).Item(0).ToString) = True Then
    '            If String.Equals(mode, "24") = True Then
    '                Mypanel_qty = dimen_table.Rows(i).Item(3).ToString
    '                Exit For
    '            Else
    '                Mypanel_qty = dimen_table.Rows(i).Item(5).ToString
    '                Exit For
    '            End If
    '        End If
    '    Next

    'End Function

    'Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
    '    RIO_wiz.ShowDialog()
    'End Sub

    Function Rio_qty(feature_desc As String) As Double
        'calculate the needed and available RIO

        Rio_qty = 0

        Dim n_inputs As Double : n_inputs = 0
        Dim n_outputs As Double : n_outputs = 0
        Dim n_motion As Double : n_motion = 0

        Dim iot = From inp In table_remote_io Where inp.Field(Of String)("solutiom") = sol_label.Text And inp.Field(Of String)("description") = feature_desc Select inp

        For Each row As DataRow In iot

            If row.Item("inputs") > 0 Then
                n_inputs = Math.Ceiling(inputs_io / row.Item("inputs"))
            End If

            If row.Item("output") > 0 Then
                n_outputs = Math.Ceiling(outputs_io / row.Item("output"))
            End If

            If row.Item("motion") > 0 Then
                n_motion = Math.Ceiling(motion_io / row.Item("motion"))
            End If

        Next

        Rio_qty = Max_number(n_inputs, n_outputs, n_motion)


    End Function

    Function Max_number(num1 As Double, num2 As Double, num3 As Double) As Double

        'Return the max number of the parameters passed
        Max_number = num1

        Dim array_t(4) As Double
        array_t(0) = num1
        array_t(1) = num2
        array_t(2) = num3

        Max_number = array_t.Max

    End Function



    Sub isinAlloc_update(part As String, mfg As String, vendor As String, n_qty As Double, cost As Double)

        'find the part in the alloc table and if it exist update qty

        Dim isinAlloc As Boolean : isinAlloc = False


        For i = 0 To my_alloc_table.Rows.Count - 1
            If String.Compare(part, my_alloc_table.Rows(i).Item(0)) = 0 Then
                my_alloc_table.Rows(i).Item(5) = n_qty + my_alloc_table.Rows(i).Item(5)
                isinAlloc = True
                Exit For
            End If
        Next

        If isinAlloc = False Then
            my_alloc_table.Rows.Add(part, "", mfg, vendor, cost, n_qty, cost * n_qty, "Project_specific", "Project_specific", "")
        End If


    End Sub

    Private Sub rio_grid2_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles rio_grid2.CellValueChanged

        '--- update cell in RIO needed grid (This sub used to be in RIO wiz)
        If RIO_grid.Rows.Count > 1 And do_not_IO_recal = False Then

            Dim av_inputs As Double = av_inputs = 0
            Dim av_outputs As Double = av_outputs = 0
            Dim av_motion As Double = av_motion = 0


            For i = 0 To rio_grid2.Rows.Count - 1

                If IsNumeric(rio_grid2.Rows(i).Cells(2).Value) = True Then

                    If rio_grid2.Rows(i).Cells(2).Value > 0 Then

                        Dim fc As String : fc = rio_grid2.Rows(i).Cells(0).Value

                        Dim iot = From inp In table_remote_io Where inp.Field(Of String)("solutiom") = "EIP-MS/EIP-RIO" And inp.Field(Of String)("feature_code") = fc Select inp

                        For Each row As DataRow In iot
                            av_inputs = av_inputs + CType(row.Item("inputs"), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)
                            av_outputs = av_outputs + CType(row.Item("output"), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)
                            av_motion = av_motion + CType(row.Item("motion"), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)
                        Next


                    End If
                End If

            Next

            RIO_grid.Rows(1).Cells(1).Value = Math.Ceiling(av_inputs + 1)
            RIO_grid.Rows(1).Cells(2).Value = Math.Ceiling(av_outputs + 1)
            RIO_grid.Rows(1).Cells(3).Value = Math.Ceiling(av_motion + 1)


            If RIO_grid.Rows(1).Cells(1).Value < RIO_grid.Rows(0).Cells(1).Value Then
                RIO_grid.Rows(1).Cells(1).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(1).Style.BackColor = Color.LightGray
            End If

            If RIO_grid.Rows(1).Cells(2).Value < RIO_grid.Rows(0).Cells(2).Value Then
                RIO_grid.Rows(1).Cells(2).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(2).Style.BackColor = Color.LightGray
            End If

            If RIO_grid.Rows(1).Cells(3).Value < RIO_grid.Rows(0).Cells(3).Value Then
                RIO_grid.Rows(1).Cells(3).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(3).Style.BackColor = Color.LightGray
            End If

        End If

    End Sub

    Private Sub SaveQuoteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveQuoteToolStripMenuItem.Click
        If String.Equals("My Quote", Me.Text) = False Then

            '--------- delete quote first ---------

            '--- first check if its a release quote, if its not then delete it fro my_active_quote_table
            Try

                Dim Create_cmd As New MySqlCommand   'my_active_quote_table

                'remove data from my_active_quote_Table
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.CommandText = "delete from quote_table.my_active_quote_table where quote_name = @quote_n"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Create_cmd.Parameters.Clear()

                'remove data from g_Setup
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.CommandText = "delete from quote_table.g_setup where quote_name = @quote_n"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Create_cmd.Parameters.Clear()

                'remove data from my_inputs
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.CommandText = "delete from quote_table.my_inputs where quote_name = @quote_n"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Create_cmd.Parameters.Clear()

                'remove data from install table
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.CommandText = "delete from quote_table.install_table where quote_name = @quote_n"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Create_cmd.Parameters.Clear()

                'remove data from quote table
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.CommandText = "delete from quote_table.total_q_table where quote_name = @quote_n"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Create_cmd.Parameters.Clear()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            '--save quote
            '---------- main active_quote table-------------
            Dim v_480 As Double : v_480 = 0
            Dim v_230 As Double : v_230 = 0
            Dim v_575 As Double : v_575 = 0
            Dim PF525_d As Double : PF525_d = 0
            Dim DE1_d As Double : DE1_d = 0
            Dim spare_io As Double
            Dim spare_panel_space As Double
            Dim Apanel As String
            Dim labor_Cost As Double
            Dim include_d As Double : include_d = 0
            Dim t_feet As Double : t_feet = 0

            Try
                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                main_cmd.Parameters.AddWithValue("@user", current_user)
                main_cmd.CommandText = "INSERT INTO quote_table.my_active_quote_table(quote_name, created_by, Date_Created) VALUES (@quote_n,@user,now())"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()


                '------ save global setups --------
                If volts_480.Checked = True Then
                    v_480 = 1
                ElseIf volts_575.Checked = True Then
                    v_575 = 1
                Else
                    v_230 = 1
                End If

                If PF525_c.Checked = True Then
                    PF525_d = 1
                Else
                    DE1_d = 1
                End If

                If include_a.Checked = True Then
                    include_d = 1
                Else
                    include_d = 0
                End If

                spare_io = If(IsNumeric(percentage_io.Text), percentage_io.Text, 0)
                spare_panel_space = If(IsNumeric(percentage_panel.Text), percentage_panel.Text, 0)
                Apanel = apanel_box.Text
                labor_Cost = If(IsNumeric(labor_box.Text), labor_box.Text, 0)

                '-- save total feet
                t_feet = If(IsNumeric(t_feet_box.Text), t_feet_box.Text, 0)


                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@quote_n", Me.Text)
                Create_cmd.Parameters.AddWithValue("@v_480", CType(v_480, Double))
                Create_cmd.Parameters.AddWithValue("@v_575", CType(v_575, Double))
                Create_cmd.Parameters.AddWithValue("@v_230", CType(v_230, Double))
                Create_cmd.Parameters.AddWithValue("@PF525_c", CType(PF525_d, Double))
                Create_cmd.Parameters.AddWithValue("@DE1_c", CType(DE1_d, Double))
                Create_cmd.Parameters.AddWithValue("@spare_io", CType(spare_io, Double))
                Create_cmd.Parameters.AddWithValue("@spare_panel_space", CType(spare_panel_space, Double))
                Create_cmd.Parameters.AddWithValue("@apanel", Apanel)
                Create_cmd.Parameters.AddWithValue("@labor_Cost", CType(labor_Cost, Double))
                Create_cmd.Parameters.AddWithValue("@include_a", CType(include_d, Double))
                Create_cmd.Parameters.AddWithValue("@t_feet", CType(t_feet, Double))
                Create_cmd.CommandText = "INSERT INTO quote_table.g_setup(quote_name, v_480, v_575, v_230, PF525_c, DE1_c, spare_io, spare_panel_space, A_panel ,labor_Cost, include_a, t_feet) VALUES (@quote_n, @v_480,@v_575, @v_230, @PF525_c,@DE1_c, @spare_io,@spare_panel_space, @apanel, @labor_Cost, @include_a, @t_feet)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

                '---------- save inputs ----------

                '============ Starter Panel =============
                For i = 1 To Panel_grid.Columns.Count - 1
                    For j = 0 To Panel_grid.Rows.Count - 1

                        If IsNumeric(Panel_grid.Rows(j).Cells(i).Value) = True Then
                            Dim Create_cmd3 As New MySqlCommand
                            Create_cmd3.Parameters.Clear()
                            Create_cmd3.Parameters.AddWithValue("@quote_n", Me.Text)
                            Create_cmd3.Parameters.AddWithValue("@set_name", Panel_grid.Columns(i).HeaderText.ToString)
                            Create_cmd3.Parameters.AddWithValue("@feature_desc", Panel_grid.Rows(j).Cells(0).Value.ToString)
                            Create_cmd3.Parameters.AddWithValue("@qty", CType(Panel_grid.Rows(j).Cells(i).Value.ToString, Double))
                            Create_cmd3.Parameters.AddWithValue("@type_p", "Panel")

                            Create_cmd3.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc, qty, type_p) VALUES (@quote_n, @set_name,@feature_desc,@qty,@type_p)"
                            Create_cmd3.Connection = Login.Connection
                            Create_cmd3.ExecuteNonQuery()
                        End If
                    Next
                Next
                '================== PLC===============
                For i = 1 To PLC_grid.Columns.Count - 1
                    For j = 0 To PLC_grid.Rows.Count - 1

                        If IsNumeric(PLC_grid.Rows(j).Cells(i).Value) = True Then
                            Dim Create_cmd2 As New MySqlCommand
                            Create_cmd2.Parameters.Clear()
                            Create_cmd2.Parameters.AddWithValue("@quote_n", Me.Text)
                            Create_cmd2.Parameters.AddWithValue("@set_name", PLC_grid.Columns(i).HeaderText)
                            Create_cmd2.Parameters.AddWithValue("@feature_desc", PLC_grid.Rows(j).Cells(0).Value.ToString)
                            Create_cmd2.Parameters.AddWithValue("@qty", CType(PLC_grid.Rows(j).Cells(i).Value.ToString(), Double))
                            Create_cmd2.Parameters.AddWithValue("@type_p", "PLC")

                            Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc,qty, type_p) VALUES (@quote_n,@set_name,@feature_desc,@qty,@type_p)"
                            Create_cmd2.Connection = Login.Connection
                            Create_cmd2.ExecuteNonQuery()
                        End If
                    Next
                Next

                '================== Control Panel ===============
                For i = 1 To Control_grid.Columns.Count - 1
                    For j = 0 To Control_grid.Rows.Count - 1

                        If IsNumeric(Control_grid.Rows(j).Cells(i).Value) = True Then
                            Dim Create_cmd2 As New MySqlCommand
                            Create_cmd2.Parameters.Clear()
                            Create_cmd2.Parameters.AddWithValue("@quote_n", Me.Text)
                            Create_cmd2.Parameters.AddWithValue("@set_name", Control_grid.Columns(i).HeaderText)
                            Create_cmd2.Parameters.AddWithValue("@feature_desc", Control_grid.Rows(j).Cells(0).Value.ToString)
                            Create_cmd2.Parameters.AddWithValue("@qty", CType(Control_grid.Rows(j).Cells(i).Value.ToString(), Double))
                            Create_cmd2.Parameters.AddWithValue("@type_p", "Control Panel")

                            Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc, qty, type_p) VALUES (@quote_n,@set_name,@feature_desc,@qty,@type_p)"
                            Create_cmd2.Connection = Login.Connection
                            Create_cmd2.ExecuteNonQuery()
                        End If
                    Next
                Next

                '=================== Field ============
                For i = 1 To Field_grid.Columns.Count - 1
                    For j = 0 To Field_grid.Rows.Count - 1

                        If IsNumeric(Field_grid.Rows(j).Cells(i).Value) = True Then
                            Dim Create_cmd2 As New MySqlCommand
                            Create_cmd2.Parameters.Clear()
                            Create_cmd2.Parameters.AddWithValue("@quote_n", Me.Text)
                            Create_cmd2.Parameters.AddWithValue("@set_name", Field_grid.Columns(i).HeaderText)
                            Create_cmd2.Parameters.AddWithValue("@feature_desc", Field_grid.Rows(j).Cells(0).Value.ToString)
                            Create_cmd2.Parameters.AddWithValue("@qty", CType(Field_grid.Rows(j).Cells(i).Value.ToString, Double))
                            Create_cmd2.Parameters.AddWithValue("@type_p", "Field")

                            Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name,  set_name, feature_desc,qty, type_p) VALUES (@quote_n, @set_name,@feature_desc,@qty,@type_p)"
                            Create_cmd2.Connection = Login.Connection
                            Create_cmd2.ExecuteNonQuery()
                        End If
                    Next
                Next

                '---------- save installations -------------
                For i = 0 To Install_grid.Rows.Count - 1

                    Dim Create_cmd2 As New MySqlCommand
                    Create_cmd2.Parameters.Clear()
                    Create_cmd2.Parameters.AddWithValue("@quote_name", Me.Text)
                    Create_cmd2.Parameters.AddWithValue("@description", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(0).Value) = True, "", Install_grid.Rows(i).Cells(0).Value))
                    Create_cmd2.Parameters.AddWithValue("@qty", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value) = True, "", Install_grid.Rows(i).Cells(1).Value))
                    Create_cmd2.Parameters.AddWithValue("@u_time", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(2).Value) = True, "", Install_grid.Rows(i).Cells(2).Value))
                    Create_cmd2.Parameters.AddWithValue("@t_time", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(3).Value) = True, "", Install_grid.Rows(i).Cells(3).Value))
                    Create_cmd2.Parameters.AddWithValue("@rate", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(4).Value) = True, "", Install_grid.Rows(i).Cells(4).Value))
                    Create_cmd2.Parameters.AddWithValue("@rate_t", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(5).Value) = True, "", Install_grid.Rows(i).Cells(5).Value))
                    Create_cmd2.Parameters.AddWithValue("@rate_l", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(6).Value) = True, "", Install_grid.Rows(i).Cells(6).Value))
                    Create_cmd2.Parameters.AddWithValue("@rate_tl", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(7).Value) = True, "", Install_grid.Rows(i).Cells(7).Value))
                    Create_cmd2.Parameters.AddWithValue("@mat_u", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(8).Value) = True, "", Install_grid.Rows(i).Cells(8).Value))
                    Create_cmd2.Parameters.AddWithValue("@mat_t", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(9).Value) = True, "", Install_grid.Rows(i).Cells(9).Value))
                    Create_cmd2.Parameters.AddWithValue("@exp_c", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(10).Value) = True, "", Install_grid.Rows(i).Cells(10).Value))
                    Create_cmd2.Parameters.AddWithValue("@exp_t", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(11).Value) = True, "", Install_grid.Rows(i).Cells(11).Value))
                    Create_cmd2.Parameters.AddWithValue("@total_t", If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(12).Value) = True, "", Install_grid.Rows(i).Cells(12).Value))


                    Create_cmd2.CommandText = "INSERT INTO quote_table.install_table(quote_name, description, qty, u_time, t_time, rate, rate_t, rate_l, rate_tl, mat_u, mat_t, exp_c, exp_t, total_t) VALUES (@quote_name, @description, @qty, @u_time, @t_time, @rate, @rate_t, @rate_l, @rate_tl, @mat_u, @mat_t, @exp_c, @exp_t, @total_t)"
                    Create_cmd2.Connection = Login.Connection
                    Create_cmd2.ExecuteNonQuery()

                Next


                '---------- save Quote totals --------
                For i = 0 To Quote_grid.Rows.Count - 1

                    Dim Create_cmd2 As New MySqlCommand
                    Create_cmd2.Parameters.Clear()
                    Create_cmd2.Parameters.AddWithValue("@quote_name", Me.Text)
                    Create_cmd2.Parameters.AddWithValue("@blank", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(0).Value) = True, "", Quote_grid.Rows(i).Cells(0).Value))
                    Create_cmd2.Parameters.AddWithValue("@description", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(1).Value) = True, "", Quote_grid.Rows(i).Cells(1).Value))
                    Create_cmd2.Parameters.AddWithValue("@qty", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(2).Value) = True, "", Quote_grid.Rows(i).Cells(2).Value))
                    Create_cmd2.Parameters.AddWithValue("@risk", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(3).Value) = True, "", Quote_grid.Rows(i).Cells(3).Value))
                    Create_cmd2.Parameters.AddWithValue("@qty_w", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(4).Value) = True, "", Quote_grid.Rows(i).Cells(4).Value))
                    Create_cmd2.Parameters.AddWithValue("@unit_c", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(5).Value) = True, "", Quote_grid.Rows(i).Cells(5).Value))
                    Create_cmd2.Parameters.AddWithValue("@total_c", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(6).Value) = True, "", Quote_grid.Rows(i).Cells(6).Value))
                    Create_cmd2.Parameters.AddWithValue("@markup", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(7).Value) = True, "", Quote_grid.Rows(i).Cells(7).Value))
                    Create_cmd2.Parameters.AddWithValue("@u_price", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(8).Value) = True, "", Quote_grid.Rows(i).Cells(8).Value))
                    Create_cmd2.Parameters.AddWithValue("@f_price", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(9).Value) = True, "", Quote_grid.Rows(i).Cells(9).Value))
                    Create_cmd2.Parameters.AddWithValue("@margin", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(10).Value) = True, "", Quote_grid.Rows(i).Cells(10).Value))
                    Create_cmd2.Parameters.AddWithValue("@type_p", If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(11).Value) = True, "", Quote_grid.Rows(i).Cells(11).Value))
                    Create_cmd2.Parameters.AddWithValue("@id", i + 1)


                    Create_cmd2.CommandText = "INSERT INTO quote_table.total_q_table(quote_name, blank, description, qty, risk, qty_w, unit_c, total_c, markup, u_price, f_price, margin, type_p, id ) VALUES (@quote_name, @blank, @description, @qty, @risk, @qty_w, @unit_c, @total_c, @markup, @u_price, @f_price, @margin, @type_p, @id)"
                    Create_cmd2.Connection = Login.Connection
                    Create_cmd2.ExecuteNonQuery()

                Next
                '======================================
                '======================================

                MessageBox.Show("Changes Saved")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Else
            MessageBox.Show("Please, Open a Quote")
        End If
    End Sub

    'this function returns the material + bulk + labor cost of an assembly
    Function Cost_of_Assem(assembly As String) As Double

        Cost_of_Assem = 0

        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("primary_vendor", GetType(String))
        datatable.Columns.Add("total_cost", GetType(Double))

        Dim labor_c As Double : labor_c = 0
        Dim bulk_c As Double : bulk_c = 0
        Dim mat_c As Double : mat_c = 0


        Try
            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@ADA", assembly)
            cmd3.CommandText = "SELECT p1.Part_Name, adv.Qty, p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.Legacy_ADA  = @ADA"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    datatable.Rows.Add(reader3(0).ToString, reader3(1).ToString, reader3(2).ToString, 0)
                End While
            End If

            reader3.Close()

            For i = 0 To datatable.Rows.Count - 1
                datatable.Rows(i).Item(3) = Form1.Get_Latest_Cost(Login.Connection, datatable.Rows(i).Item(0), datatable.Rows(i).Item(2)) * datatable.Rows(i).Item(1)
                mat_c = mat_c + datatable.Rows(i).Item(3)
            Next

            '--get labor and bulk
            Dim cmd As New MySqlCommand
            Dim reader As MySqlDataReader

            cmd.Parameters.AddWithValue("@Legacy_ADA", assembly)
            cmd.CommandText = "SELECT Labor_Cost, Bulk_Cost from devices where Legacy_ADA_Number = @Legacy_ADA"
            cmd.Connection = Login.Connection
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    labor_c = reader(0)
                    bulk_c = reader(1)
                End While
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Cost_of_Assem = labor_c + bulk_c + mat_c

    End Function

    '--- get total qty of feature codes ----
    Function get_field_qty(feature_code As String, grid As DataGridView)

        get_field_qty = 0

        For i = 0 To grid.Rows.Count - 1

            If String.Equals(grid.Rows(i).Cells(0).Value.ToString, feature_code) = True Then
                For j = 1 To grid.Columns.Count - 1
                    get_field_qty = get_field_qty + If(IsNumeric(grid.Rows(i).Cells(j).Value) = True, grid.Rows(i).Cells(j).Value, 0)
                Next
            End If
        Next


    End Function



    Sub Fill_installation()


        Install_grid.Rows(1).Cells(2).Value = "1" 'PE
        Install_grid.Rows(2).Cells(2).Value = "1" 'Pull Cord Stations
        Install_grid.Rows(3).Cells(2).Value = "0.125" 'Pull Cord Cable/eyelit (per foot)
        Install_grid.Rows(4).Cells(2).Value = "0.5" 'Pull Cord Pulley Kits
        Install_grid.Rows(5).Cells(2).Value = "1" 'OS Station
        Install_grid.Rows(6).Cells(2).Value = "1" 'Limit/Prox Switch
        Install_grid.Rows(7).Cells(2).Value = "1" 'Pressure Switch
        Install_grid.Rows(8).Cells(2).Value = "1" 'Beacons
        Install_grid.Rows(9).Cells(2).Value = "1" 'Horns
        Install_grid.Rows(10).Cells(2).Value = "1" 'Solenoids
        Install_grid.Rows(11).Cells(2).Value = "2" 'Interface form-Robot Controller
        Install_grid.Rows(12).Cells(2).Value = "1" 'Interface Box
        Install_grid.Rows(13).Cells(2).Value = "1" 'Clutch Brake J-Box
        Install_grid.Rows(14).Cells(2).Value = "0.85" 'I/O Receptacle
        Install_grid.Rows(15).Cells(2).Value = "1" 'Encoder
        Install_grid.Rows(16).Cells(2).Value = "8" 'Scanners w/ Frame ( No commissioning)
        Install_grid.Rows(17).Cells(2).Value = "4" 'Scale Controller Install with frame
        Install_grid.Rows(18).Cells(2).Value = "4" 'Print & Apply (Self Supported)
        Install_grid.Rows(19).Cells(2).Value = "12" 'Print & Apply (With 80/20 Frame)
        Install_grid.Rows(20).Cells(2).Value = "1.5" 'ADA Box Mounting
        Install_grid.Rows(21).Cells(2).Value = "1" '24VDC Motor Drops
        Install_grid.Rows(22).Cells(2).Value = "0.125" 'Open wire basket cable tray (Per Foot)
        Install_grid.Rows(23).Cells(2).Value = "0.75" '24VDC Interfaced Zones
        Install_grid.Rows(24).Cells(2).Value = "6" '2” EMT Chase - Personnel Gate
        Install_grid.Rows(25).Cells(2).Value = "1" '2” EMT Chase
        Install_grid.Rows(26).Cells(2).Value = "0.75" 'Sorter Terminations ( Control Only / Includes Cables )
        Install_grid.Rows(27).Cells(2).Value = "1" 'Sorter Terminations ( VFD & Controls )
        Install_grid.Rows(28).Cells(2).Value = "1" 'Sorter Terminations ( Device Net )
        Install_grid.Rows(29).Cells(2).Value = "0.05" 'Ethernet cabling (Surface Mounted)
        Install_grid.Rows(30).Cells(2).Value = "0.25" 'Ethernet Terminations
        Install_grid.Rows(31).Cells(2).Value = "1" 'Man Lift per week
        Install_grid.Rows(34).Cells(2).Value = "50" 'Stand-by & Supervision
        Install_grid.Rows(35).Cells(2).Value = "8" 'Travel Time (per  round trip) 35
        Install_grid.Rows(44).Cells(2).Value = "12" '400amp 480v 3PH PDP Install
        Install_grid.Rows(45).Cells(2).Value = "8" '200amp 480v 3PH PDP Install
        Install_grid.Rows(46).Cells(2).Value = "8" '125amp 120/240 PDP Install
        Install_grid.Rows(47).Cells(2).Value = "8" '50KVA 480v-240v XFMR Install
        Install_grid.Rows(49).Cells(2).Value = "8" 'Air Compressor Wiring W/ Disc. 15HP Max.
        Install_grid.Rows(50).Cells(2).Value = "0.035" '480v 30a 3-phase 3 wire Ckt. (L.F.)
        Install_grid.Rows(53).Cells(2).Value = "0.5" '120v 20a Quad Receptacles
        Install_grid.Rows(54).Cells(2).Value = "0.035" '120v 20a Ckt (L.F.)
        Install_grid.Rows(55).Cells(2).Value = "0.5" '120v 30a Twist Lock Receptacles
        Install_grid.Rows(56).Cells(2).Value = "0.035" '120v 30a 1-PH Ckt. (L.F.)
        Install_grid.Rows(60).Cells(2).Value = "48" 'Demo of existing system
        Install_grid.Rows(61).Cells(2).Value = "40" 'Retro fit exisisting system to new

        For i = 1 To Install_grid.Rows.Count - 2
            If i <> 33 And i <> 36 And i <> 42 Then
                Install_grid.Rows(i).Cells(4).Value = 80  'subcontract rate $
            End If
        Next

        '-----material unit cost
        Install_grid.Rows(1).Cells(8).Value = "5" 'PE
        Install_grid.Rows(2).Cells(8).Value = "5" 'Pull Cord Stations
        Install_grid.Rows(3).Cells(8).Value = "0.5" 'Pull Cord Cable/eyelit (per foot)
        Install_grid.Rows(4).Cells(8).Value = "0.5" 'Pull Cord Pulley Kits
        Install_grid.Rows(5).Cells(8).Value = "5" 'OS Station
        Install_grid.Rows(6).Cells(8).Value = "3.75" 'Limit/Prox Switch
        Install_grid.Rows(7).Cells(8).Value = "3.75" 'Pressure Switch
        Install_grid.Rows(8).Cells(8).Value = "20" 'Beacons
        Install_grid.Rows(9).Cells(8).Value = "20" 'Horns
        Install_grid.Rows(10).Cells(8).Value = "3.75" 'Solenoids
        Install_grid.Rows(11).Cells(8).Value = "6.75" 'Interface form-Robot Controller
        Install_grid.Rows(12).Cells(8).Value = "6.75" 'Interface Box
        Install_grid.Rows(13).Cells(8).Value = "6.75" 'Clutch Brake J-Box
        Install_grid.Rows(14).Cells(8).Value = "5" 'I/O Receptacle
        Install_grid.Rows(15).Cells(8).Value = "15" 'Encoder
        Install_grid.Rows(16).Cells(8).Value = "800" 'Scanners w/ Frame ( No commissioning)
        Install_grid.Rows(17).Cells(8).Value = "400" 'Scale Controller Install with frame
        Install_grid.Rows(18).Cells(8).Value = "25" 'Print & Apply (Self Supported)
        Install_grid.Rows(19).Cells(8).Value = "800" 'Print & Apply (With 80/20 Frame)
        Install_grid.Rows(20).Cells(8).Value = "75" 'ADA Box Mounting
        Install_grid.Rows(21).Cells(8).Value = "15" '24VDC Motor Drops
        Install_grid.Rows(22).Cells(8).Value = "11.5" 'Open wire basket cable tray (Per Foot)
        Install_grid.Rows(23).Cells(8).Value = "0.75" '24VDC Interfaced Zones
        Install_grid.Rows(24).Cells(8).Value = "375" '2” EMT Chase - Personnel Gate
        Install_grid.Rows(25).Cells(8).Value = "150" '2” EMT Chase
        Install_grid.Rows(26).Cells(8).Value = "35" 'Sorter Terminations ( Control Only / Includes Cables )
        Install_grid.Rows(27).Cells(8).Value = "25" 'Sorter Terminations ( VFD & Controls )
        Install_grid.Rows(28).Cells(8).Value = "115" 'Sorter Terminations ( Device Net )
        Install_grid.Rows(29).Cells(8).Value = "0.35" 'Ethernet cabling (Surface Mounted)
        Install_grid.Rows(30).Cells(8).Value = "7.50" 'Ethernet Terminations
        Install_grid.Rows(31).Cells(8).Value = "1000" 'Man Lift per week
        Install_grid.Rows(34).Cells(8).Value = "0" 'Stand-by & Supervision
        Install_grid.Rows(35).Cells(8).Value = "0" 'Travel Time (per  round trip) 
        Install_grid.Rows(44).Cells(8).Value = "8750" '400amp 480v 3PH PDP Install
        Install_grid.Rows(45).Cells(8).Value = "6500" '200amp 480v 3PH PDP Install
        Install_grid.Rows(46).Cells(8).Value = "2500" '125amp 120/240 PDP Install
        Install_grid.Rows(47).Cells(8).Value = "1750" '50KVA 480v-240v XFMR Install
        Install_grid.Rows(48).Cells(8).Value = "600" 'ADA 30a 480v Power drops (per price)
        Install_grid.Rows(49).Cells(8).Value = "600" 'Air Compressor Wiring W/ Disc. 15HP Max.
        Install_grid.Rows(50).Cells(8).Value = "4.10" '480v 30a 3-phase 3 wire Ckt. (L.F.)
        Install_grid.Rows(51).Cells(8).Value = "500" ' Motors(price per)
        Install_grid.Rows(52).Cells(8).Value = "110" 'Motor disconnects (price per)
        Install_grid.Rows(53).Cells(8).Value = "75" '120v 20a Quad Receptacles
        Install_grid.Rows(54).Cells(8).Value = "3.75" '120v 20a Ckt (L.F.)
        Install_grid.Rows(55).Cells(8).Value = "50" '120v 30a Twist Lock Receptacles
        Install_grid.Rows(56).Cells(8).Value = "3.85" '120v 30a 1-PH Ckt. (L.F.)
        Install_grid.Rows(57).Cells(8).Value = "500" 'Power supply feeds
        Install_grid.Rows(60).Cells(8).Value = "0" 'Demo of existing system
        Install_grid.Rows(61).Cells(8).Value = "0" 'Retro fit exisisting system to new

        '---labor rate
        Install_grid.Rows(34).Cells(6).Value = "70" 'Stand-by & Supervision
        Install_grid.Rows(35).Cells(6).Value = "70" 'Travel Time (per  round trip) 35

        '------ expenses
        Install_grid.Rows(37).Cells(10).Value = "200" 'travel expenses per day
        Install_grid.Rows(38).Cells(10).Value = "400" 'Travel expenses per day low voltage
        Install_grid.Rows(39).Cells(10).Value = "0.55" 'mileage
        Install_grid.Rows(40).Cells(10).Value = "750" 'airfare
        Install_grid.Rows(41).Cells(10).Value = "65" 'Tcar renta;

    End Sub

    Private Sub markup_p_TextChanged(sender As Object, e As EventArgs) Handles markup_p.TextChanged
        '--change markup for panels
        If start_recal = True Then
            For i = 1 To 15
                If i <> 7 Then
                    Quote_grid.Rows(i).Cells(7).Value = markup_p.Text
                End If
            Next
        End If
    End Sub

    Private Sub markup_f_TextChanged(sender As Object, e As EventArgs) Handles markup_f.TextChanged
        '--change markup for field
        If start_recal = True Then
            For i = 19 To 42
                If i <> 27 Then
                    Quote_grid.Rows(i).Cells(7).Value = markup_f.Text
                End If
            Next
        End If
    End Sub

    Private Sub markup_e_TextChanged(sender As Object, e As EventArgs) Handles markup_e.TextChanged
        '--change markup for install
        If start_recal = True Then
            For i = 79 To 82
                Quote_grid.Rows(i).Cells(7).Value = markup_e.Text
            Next
        End If
    End Sub

    Private Sub Control_grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Control_grid.CellDoubleClick

        If String.IsNullOrEmpty(Control_grid.CurrentCell.Value) = False Then

            If String.Equals(Control_grid.CurrentCell.Value.ToString, "") = False And String.Equals(Control_grid.CurrentCell.Value.ToString, " ") = False Then

                Dim component As String : component = Control_grid.CurrentCell.Value.ToString.Replace("/", "-")

                If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
                    System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
                End If

            End If

        End If



    End Sub

    Private Sub InstallationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstallationToolStripMenuItem.Click
        '-- zero out install table

        For i = 1 To Install_grid.Rows.Count - 1
            For j = 1 To Install_grid.Columns.Count - 1
                If i <> 33 And i <> 36 And i <> 42 And j <> 2 Then
                    Install_grid.Rows(i).Cells(j).Value = 0
                End If
            Next
        Next

        Call Fill_installation()

    End Sub



    Private Sub sub_rate_box_TextChanged(sender As Object, e As EventArgs) Handles sub_rate_box.TextChanged

        For i = 1 To Install_grid.Rows.Count - 2
            If i <> 33 And i <> 36 And i <> 42 Then
                Install_grid.Rows(i).Cells(4).Value = sub_rate_box.Text 'subcontract rate $
            End If
        Next
    End Sub

    Private Sub labor_rate_box_TextChanged(sender As Object, e As EventArgs) Handles labor_rate_box.TextChanged

        'For i = 1 To Install_grid.Rows.Count - 2
        '    If i <> 33 And i <> 36 And i <> 42 Then
        '        Install_grid.Rows(i).Cells(6).Value = labor_rate_box.Text 'labor rate $
        '    End If
        'Next
    End Sub

    Private Sub Material_rate_box_TextChanged(sender As Object, e As EventArgs) Handles Material_rate_box.TextChanged
        For i = 1 To Install_grid.Rows.Count - 2
            If i <> 33 And i <> 36 And i <> 42 Then
                Install_grid.Rows(i).Cells(8).Value = Material_rate_box.Text 'labor rate $
            End If
        Next
    End Sub

    Private Sub Exp_rate_box_TextChanged(sender As Object, e As EventArgs) Handles Exp_rate_box.TextChanged
        For i = 1 To Install_grid.Rows.Count - 2
            If i <> 33 And i <> 36 And i <> 42 Then
                Install_grid.Rows(i).Cells(10).Value = Exp_rate_box.Text 'labor rate $
            End If
        Next
    End Sub

    Sub load_total_quote()
        '--refresh entire total quote table
        donotcompute = False

        For i = 5 To 6
            Quote_grid.Columns(i).HeaderCell.Style.BackColor = Color.DarkSalmon
        Next
        For i = 8 To 9
            Quote_grid.Columns(i).HeaderCell.Style.BackColor = Color.PowderBlue
        Next

        Quote_grid.Rows.Add(New String() {"", "ADA Panels and options"})
        Quote_grid.Rows(0).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(0).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(0).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(0).ReadOnly = True


        Quote_grid.Rows.Add(New String() {"", "Starter Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Control Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "PLC Panel – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Bulk - Material Cost", "", "", "1", "INFO HERE", "", "20", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "Labor Cost", "", "", "1", "INFO HERE", "", "20", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "", "", "", "", "", "", "", ""})

        For i = 1 To 6
            Quote_grid.Rows(i).Cells(1).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {"", "Project Specific Items below"})
        Quote_grid.Rows(7).DefaultCellStyle.BackColor = Color.SlateGray
        Quote_grid.Rows(7).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(7).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20", "", "", "", "Project Materials"})

        For i = 8 To 15
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next


        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(16).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(16).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(16).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(17).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(17).ReadOnly = True


        Quote_grid.Rows.Add(New String() {"", "Field Items"})
        Quote_grid.Rows(18).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(18).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(18).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(18).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Field Parts – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "1", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "M12 Cables – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "M12 Estop Cables – Materials", "", "", "1", "INFO HERE", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})

        'this just to be camera and m12 labor
        Quote_grid.Rows(23).ReadOnly = True
        Quote_grid.Rows(24).ReadOnly = True
        Quote_grid.Rows(25).ReadOnly = True
        Quote_grid.Rows(26).ReadOnly = True

        For i = 19 To 26
            Quote_grid.Rows(i).Cells(1).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {"", "Project Specific Items below"})
        Quote_grid.Rows(27).DefaultCellStyle.BackColor = Color.SlateGray
        Quote_grid.Rows(27).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(27).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "ZOE HMI PC (Linux)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "ZOE HMI Enclosure (Global)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanners (Top)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanners Side (Single)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Scanner Frame (Top)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Scanner Frame (Side)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Scanner Tech (#Days)", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Subcontract"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "$0", "", "20", "", "", "", "Project Materials"})

        For i = 28 To 42
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(43).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(43).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(43).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(44).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(44).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "In-Office Labor"})
        Quote_grid.Rows(45).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(45).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(45).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(45).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Engineering", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Layouts", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Panel Sets", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "PLC Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "HMI Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Services Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Reports Programming", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "FAT", "", "", "0", "$62", "", "", "$105", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Documentation", "", "", "0", "$62", "", "", "$75", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "Project Management", "", "", "0", "$62", "", "", "$95", "", "", "Project Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Project Labor"})

        For i = 46 To 56
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 46 To 56
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(57).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(57).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(57).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(58).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(58).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Onsite Labor"})
        Quote_grid.Rows(59).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(59).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(59).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(59).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "I/O Checkout", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "PLC Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "HMI Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "WCS Commissioning", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Standby", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "SAT", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Training", "", "", "0", "$62", "", "", "$95", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "", "", "", "", "Startup Labor"})

        For i = 60 To 68
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 60 To 68
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(69).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(69).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(69).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(70).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(70).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Start Up Expenses And Travel"})
        Quote_grid.Rows(71).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(71).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(71).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(71).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Travel Time (TT)", "0", "", "0", "$62", "", "", "$75", "", "", "Startup Labor"})
        Quote_grid.Rows.Add(New String() {"", "Mileage", "0", "", "0", "$0.57", "", "", "$0.57", "", "", "Startup Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Start-Up Expenses (SE)", "0", "", "0", "$250", "", "", "$250", "", "", "Startup Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Airfare (SE)", "0", "", "0", "$750", "", "", "$750", "", "", "Startup Expenses"})

        For i = 72 To 75
            Quote_grid.Rows(i).Cells(2).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        For i = 72 To 75
            Quote_grid.Rows(i).Cells(7).ReadOnly = True
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(76).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(76).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(76).Cells(9).Style.BackColor = Color.PowderBlue

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(77).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(77).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "1", "include?"})
        Quote_grid.Rows(78).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(78).Cells(2).Style.BackColor = Color.White
        Quote_grid.Rows(78).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(78).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(78).ReadOnly = True
        Quote_grid.Rows(78).Cells(2).ReadOnly = False

        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Labor", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Labor"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Materials", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Materials"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Expenses", "1", "", "1", "INFO HERE", "", "0%", "", "", "", "Install Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Subcontract", "1", "", "1", "INFO HERE", "", "15%", "", "", "", "Install Subcontract"})


        For i = 79 To 82
            Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.White
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(83).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(83).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(83).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(83).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(84).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(84).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Est. Shipping", "1", "include?"})
        Quote_grid.Rows(85).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(85).Cells(2).Style.BackColor = Color.White
        Quote_grid.Rows(85).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(85).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(85).ReadOnly = True
        Quote_grid.Rows(85).Cells(2).ReadOnly = False

        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING I/O", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING Motor Big", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING – Scanners", "", "", "", "$300", "", "0", "", "", "", "Project Materials"})

        For i = 86 To 88
            Quote_grid.Rows(i).Cells(2).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153)
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.White
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(89).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(89).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(89).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(89).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(90).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(90).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(91).DefaultCellStyle.BackColor = Color.DarkSlateGray
        Quote_grid.Rows(91).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(92).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(92).ReadOnly = True

        '-------------- TOTAL SECTION --------------

        Quote_grid.Rows.Add(New String() {"", "TOTALS – By Component"})
        Quote_grid.Rows(93).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(93).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(93).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(93).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "ADA Panels", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Field Items", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "In-Office Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Onsite Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Start Up Expenses And Travel", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Shipping", "", "", "", "", "", "", "", "", ""})

        For i = 94 To 100
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(101).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(101).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(101).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(101).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(101).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(102).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(102).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – Sum all but Installation"})
        Quote_grid.Rows(103).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(103).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(103).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(103).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Hardware and Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Shipping", "", "", "", "", "", "", "", "", ""})

        For i = 104 To 106
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255) : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(107).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(107).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(107).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(107).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(107).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(108).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(108).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – By Type"})
        Quote_grid.Rows(109).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(109).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(109).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(109).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Project Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Project Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Project Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Manufacturing Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Manufacturing Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Startup Subcontract", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Install Subcontract", "", "", "", "", "", "", "", "", ""})

        For i = 110 To 122
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(123).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(123).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(123).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(123).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(123).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(124).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(124).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "TOTALS – Summary"})
        Quote_grid.Rows(125).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        Quote_grid.Rows(125).Cells(12).Style.BackColor = Color.Silver
        Quote_grid.Rows(125).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows(125).ReadOnly = True

        Quote_grid.Rows.Add(New String() {"", "Labor", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Materials", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Expenses", "", "", "", "", "", "", "", "", ""})
        Quote_grid.Rows.Add(New String() {"", "Subcontract", "", "", "", "", "", "", "", "", ""})

        For i = 126 To 129
            Quote_grid.Rows(i).Cells(2).ReadOnly = True : Quote_grid.Rows(i).Cells(3).ReadOnly = True : Quote_grid.Rows(i).Cells(4).ReadOnly = True : Quote_grid.Rows(i).Cells(5).ReadOnly = True
            Quote_grid.Rows(i).Cells(0).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(1).Style.BackColor = Color.White : Quote_grid.Rows(i).Cells(2).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(3).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(5).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(4).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(6).Style.BackColor = Color.DarkSalmon : Quote_grid.Rows(i).Cells(7).Style.BackColor = Color.Gainsboro
            Quote_grid.Rows(i).Cells(8).Style.BackColor = Color.Gainsboro : Quote_grid.Rows(i).Cells(9).Style.BackColor = Color.PowderBlue
            Quote_grid.Rows(i).Cells(10).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(11).Style.BackColor = Color.Silver : Quote_grid.Rows(i).Cells(12).Style.BackColor = Color.Silver
        Next

        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(130).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(130).Cells(6).Style.BackColor = Color.DarkSalmon
        Quote_grid.Rows(130).Cells(9).Style.BackColor = Color.PowderBlue
        Quote_grid.Rows(130).Cells(10).Style.BackColor = Color.Yellow
        Quote_grid.Rows(130).ReadOnly = True


        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows(131).DefaultCellStyle.BackColor = Color.Silver
        Quote_grid.Rows(131).ReadOnly = True

        Quote_grid.Rows.Add(New String() {""}) : Quote_grid.Rows.Add(New String() {""}) : Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "EXPENSE CALCULATIONS"}) : Quote_grid.Rows(135).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {"", "Config"}) : Quote_grid.Rows(136).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "hours/ day onsite", "10"}) : Quote_grid.Rows(138).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(138).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "days onsite / trip", "12"}) : Quote_grid.Rows(139).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(139).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Travel time per trip", "12"}) : Quote_grid.Rows(140).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(140).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Miles per trip", "70"}) : Quote_grid.Rows(141).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(141).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {"", "Extra days of expenses per trip (travel days)", "0"}) : Quote_grid.Rows(142).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(142).Cells(2).Style.BackColor = Color.SpringGreen
        Quote_grid.Rows.Add(New String() {""})


        Quote_grid.Rows.Add(New String() {"", "Output"}) : Quote_grid.Rows(144).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {"", "DAYS ONSITE", ""}) : Quote_grid.Rows(146).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(146).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(146).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "# TRIPS", ""}) : Quote_grid.Rows(147).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(147).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(147).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Travel Time (TT)", ""}) : Quote_grid.Rows(148).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(148).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(148).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Mileage", ""}) : Quote_grid.Rows(149).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(149).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(149).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Start-Up Expenses (SE)", ""}) : Quote_grid.Rows(150).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(150).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(150).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {"", "Airfare (SE)", ""}) : Quote_grid.Rows(151).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold) : Quote_grid.Rows(151).Cells(2).Style.BackColor = Color.FromArgb(255, 255, 153) : Quote_grid.Rows(151).Cells(1).Style.BackColor = Color.FromArgb(204, 204, 255)
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})
        Quote_grid.Rows.Add(New String() {""})

        donotcompute = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '---- export quote form ------

        Label4.Visible = True
        Dim appPath As String = Application.StartupPath()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet2 As Excel.Worksheet
        Dim xlWorkSheet3 As Excel.Worksheet
        Dim xlWorkSheet4 As Excel.Worksheet
        Dim xlWorkSheet5 As Excel.Worksheet

        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            Try
                Dim wb As Excel.Workbook = xlApp.Workbooks.Open("O:\atlanta\APL\Template.xlsx")   ' Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\Template.xlsx")
                xlWorkSheet = wb.Sheets("TOTALS")

                'Start filling the vaules from TOTAL datagridview

                '----------- ADA Panels and options -------
                For i = 1 To 16
                    For j = 1 To 11
                        If j <> 6 And j <> 8 And j <> 9 And j <> 10 Then
                            If j = 7 Then
                                xlWorkSheet.Cells(i + 3, j + 1) = If(IsNumeric(Quote_grid.Rows(i).Cells(j).Value()) = True, Quote_grid.Rows(i).Cells(j).Value() / 100, Quote_grid.Rows(i).Cells(j).Value())
                            Else
                                xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                            End If
                        End If
                    Next
                Next
                '--------- Field Items -------
                For i = 19 To 43
                    For j = 1 To 11
                        If j <> 6 And j <> 8 And j <> 9 And j <> 10 Then
                            If j = 7 Then
                                xlWorkSheet.Cells(i + 3, j + 1) = If(IsNumeric(Quote_grid.Rows(i).Cells(j).Value()) = True, Quote_grid.Rows(i).Cells(j).Value() / 100, Quote_grid.Rows(i).Cells(j).Value())
                            Else
                                xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                            End If
                        End If
                    Next
                Next

                '-------- In Office Labor --------
                For i = 46 To 57
                    For j = 1 To 11
                        If j <> 6 And j <> 5 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '-------- Onsite Labor --------
                For i = 60 To 69
                    For j = 1 To 11
                        If j <> 6 And j <> 5 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '------ Start Up Expenses and Travel ----
                For i = 72 To 76
                    For j = 1 To 11
                        If j <> 2 And j <> 6 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '--------- Electrical installation -------
                For i = 79 To 83
                    For j = 1 To 11
                        If j <> 6 And j <> 8 And j <> 9 And j <> 10 Then
                            If j = 7 Then
                                xlWorkSheet.Cells(i + 3, j + 1) = If(IsNumeric(Quote_grid.Rows(i).Cells(j).Value()) = True, Quote_grid.Rows(i).Cells(j).Value() / 100, Quote_grid.Rows(i).Cells(j).Value())
                            Else
                                xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                            End If
                        End If
                    Next
                Next

                '----- Est Shipping -----------
                For i = 86 To 89
                    For j = 1 To 11
                        If j <> 6 And j <> 8 And j <> 9 And j <> 10 Then
                            If j = 7 Then
                                xlWorkSheet.Cells(i + 3, j + 1) = If(IsNumeric(Quote_grid.Rows(i).Cells(j).Value()) = True, Quote_grid.Rows(i).Cells(j).Value() / 100, Quote_grid.Rows(i).Cells(j).Value())
                            Else
                                xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                            End If
                        End If
                    Next
                Next

                '------ Totals - By component
                For i = 94 To 101
                    For j = 1 To 11
                        If j <> 6 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '----- Totals -- sum all but installation
                For i = 104 To 107
                    For j = 1 To 11
                        If j <> 6 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '----- Totals -- by type
                For i = 110 To 123
                    For j = 1 To 11
                        If j <> 6 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                '----- Totals -- Summary
                For i = 126 To 130
                    For j = 1 To 11
                        If j <> 6 And j <> 9 And j <> 10 Then
                            xlWorkSheet.Cells(i + 3, j + 1) = Quote_grid.Rows(i).Cells(j).Value()
                        End If
                    Next
                Next

                xlWorkSheet.Cells(145, 3) = Quote_grid.Rows(138).Cells(2).Value()
                xlWorkSheet.Cells(146, 3) = Quote_grid.Rows(139).Cells(2).Value()
                xlWorkSheet.Cells(147, 3) = Quote_grid.Rows(140).Cells(2).Value()
                xlWorkSheet.Cells(148, 3) = Quote_grid.Rows(141).Cells(2).Value()
                xlWorkSheet.Cells(149, 3) = Quote_grid.Rows(142).Cells(2).Value()



                '------- copy BOM -----
                xlWorkSheet2 = wb.Sheets("BOM")

                xlWorkSheet2.Range("A:B").ColumnWidth = 40
                xlWorkSheet2.Range("C:C").ColumnWidth = 30
                xlWorkSheet2.Range("D:I").ColumnWidth = 20
                xlWorkSheet2.Range("A:I").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel

                Purchase_Request.Visible = True
                Purchase_Request.Visible = False

                For i As Integer = 0 To Purchase_Request.PR_grid.ColumnCount - 1

                    xlWorkSheet2.Cells(1, i + 1) = Purchase_Request.PR_grid.Columns(i).HeaderText
                    For j As Integer = 0 To Purchase_Request.PR_grid.RowCount - 1
                        xlWorkSheet2.Cells(j + 2, i + 1) = Purchase_Request.PR_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                '  Purchase_Request.Visible = False
                '-----------------------------

                '----------- Copy feature codes used--------------
                xlWorkSheet3 = wb.Sheets("Feature Codes")
                xlWorkSheet3.Range("A:B").ColumnWidth = 40
                xlWorkSheet3.Range("A:B").HorizontalAlignment = Excel.Constants.xlCenter


                Summary_devices.Visible = True
                Summary_devices.Visible = False

                For i As Integer = 0 To Summary_devices.setup_grid.ColumnCount - 1

                    xlWorkSheet3.Cells(1, i + 1) = Summary_devices.setup_grid.Columns(i).HeaderText
                    For j As Integer = 0 To Summary_devices.setup_grid.RowCount - 1
                        xlWorkSheet3.Cells(j + 2, i + 1) = Summary_devices.setup_grid.Rows(j).Cells(i).Value
                    Next j

                Next i


                '-------------- Copy Assemblies used ---------------
                xlWorkSheet4 = wb.Sheets("Assemblies used")
                xlWorkSheet4.Range("A:A").ColumnWidth = 40
                xlWorkSheet4.Range("B:G").ColumnWidth = 15
                xlWorkSheet4.Range("A:I").HorizontalAlignment = Excel.Constants.xlCenter

                Dim counter As Integer : counter = 1

                For i = 0 To table_assem.Rows.Count - 1

                    xlWorkSheet4.Cells(counter, 1) = table_assem.Rows(i).Item(0)  'legacy number or device name
                    xlWorkSheet4.Cells(counter, 1).interior.color = Color.Yellow

                    xlWorkSheet4.Cells(counter, 3) = "Labor Cost"
                    xlWorkSheet4.Cells(counter, 4) = table_assem.Rows(i).Item(1)  'legacy number or device name

                    xlWorkSheet4.Cells(counter, 5) = "Bulk Cost"
                    xlWorkSheet4.Cells(counter, 6) = table_assem.Rows(i).Item(2)  'legacy number or device name

                    counter = counter + 2

                    For j = 0 To table_adv.Rows.Count - 1
                        If String.Equals(table_adv.Rows(j).Item(3).ToString, table_assem.Rows(i).Item(0).ToString) = True Then
                            xlWorkSheet4.Cells(counter, 1) = table_adv.Rows(j).Item(0)
                            xlWorkSheet4.Cells(counter, 2) = table_adv.Rows(j).Item(2)
                            counter = counter + 1
                        End If
                    Next

                    counter = counter + 2

                Next

                '--------------- copy Installation ------------------
                xlWorkSheet5 = wb.Sheets("Install")

                '---- installation site
                Dim z As Integer : z = 7

                For i = 1 To 32
                    xlWorkSheet5.Cells(z, 2) = Install_grid.Rows(i).Cells(0).Value
                    xlWorkSheet5.Cells(z, 3) = Install_grid.Rows(i).Cells(1).Value
                    xlWorkSheet5.Cells(z, 4) = Install_grid.Rows(i).Cells(2).Value
                    xlWorkSheet5.Cells(z, 5) = Install_grid.Rows(i).Cells(3).Value
                    xlWorkSheet5.Cells(z, 6) = Install_grid.Rows(i).Cells(4).Value
                    xlWorkSheet5.Cells(z, 7) = Install_grid.Rows(i).Cells(5).Value
                    xlWorkSheet5.Cells(z, 8) = Install_grid.Rows(i).Cells(6).Value
                    xlWorkSheet5.Cells(z, 9) = Install_grid.Rows(i).Cells(7).Value
                    xlWorkSheet5.Cells(z, 10) = Install_grid.Rows(i).Cells(8).Value
                    xlWorkSheet5.Cells(z, 11) = Install_grid.Rows(i).Cells(9).Value
                    xlWorkSheet5.Cells(z, 12) = Install_grid.Rows(i).Cells(10).Value
                    xlWorkSheet5.Cells(z, 13) = Install_grid.Rows(i).Cells(11).Value
                    xlWorkSheet5.Cells(z, 14) = Install_grid.Rows(i).Cells(12).Value

                    z = z + 1
                Next

                z = z + 1

                'supervision
                For i = 34 To 35
                    xlWorkSheet5.Cells(z, 2) = Install_grid.Rows(i).Cells(0).Value
                    xlWorkSheet5.Cells(z, 3) = Install_grid.Rows(i).Cells(1).Value
                    xlWorkSheet5.Cells(z, 4) = Install_grid.Rows(i).Cells(2).Value
                    xlWorkSheet5.Cells(z, 5) = Install_grid.Rows(i).Cells(3).Value
                    xlWorkSheet5.Cells(z, 6) = Install_grid.Rows(i).Cells(4).Value
                    xlWorkSheet5.Cells(z, 7) = Install_grid.Rows(i).Cells(5).Value
                    xlWorkSheet5.Cells(z, 8) = Install_grid.Rows(i).Cells(6).Value
                    xlWorkSheet5.Cells(z, 9) = Install_grid.Rows(i).Cells(7).Value
                    xlWorkSheet5.Cells(z, 10) = Install_grid.Rows(i).Cells(8).Value
                    xlWorkSheet5.Cells(z, 11) = Install_grid.Rows(i).Cells(9).Value
                    xlWorkSheet5.Cells(z, 12) = Install_grid.Rows(i).Cells(10).Value
                    xlWorkSheet5.Cells(z, 13) = Install_grid.Rows(i).Cells(11).Value
                    xlWorkSheet5.Cells(z, 14) = Install_grid.Rows(i).Cells(12).Value

                    z = z + 1
                Next

                z = z + 1
                'expenses
                For i = 37 To 41
                    xlWorkSheet5.Cells(z, 2) = Install_grid.Rows(i).Cells(0).Value
                    xlWorkSheet5.Cells(z, 3) = Install_grid.Rows(i).Cells(1).Value
                    xlWorkSheet5.Cells(z, 4) = Install_grid.Rows(i).Cells(2).Value
                    xlWorkSheet5.Cells(z, 5) = Install_grid.Rows(i).Cells(3).Value
                    xlWorkSheet5.Cells(z, 6) = Install_grid.Rows(i).Cells(4).Value
                    xlWorkSheet5.Cells(z, 7) = Install_grid.Rows(i).Cells(5).Value
                    xlWorkSheet5.Cells(z, 8) = Install_grid.Rows(i).Cells(6).Value
                    xlWorkSheet5.Cells(z, 9) = Install_grid.Rows(i).Cells(7).Value
                    xlWorkSheet5.Cells(z, 10) = Install_grid.Rows(i).Cells(8).Value
                    xlWorkSheet5.Cells(z, 11) = Install_grid.Rows(i).Cells(9).Value
                    xlWorkSheet5.Cells(z, 12) = Install_grid.Rows(i).Cells(10).Value
                    xlWorkSheet5.Cells(z, 13) = Install_grid.Rows(i).Cells(11).Value
                    xlWorkSheet5.Cells(z, 14) = Install_grid.Rows(i).Cells(12).Value

                    z = z + 1
                Next

                z = z + 1
                'power sub
                For i = 43 To Install_grid.Rows.Count - 1
                    xlWorkSheet5.Cells(z, 2) = Install_grid.Rows(i).Cells(0).Value
                    xlWorkSheet5.Cells(z, 3) = Install_grid.Rows(i).Cells(1).Value
                    xlWorkSheet5.Cells(z, 4) = Install_grid.Rows(i).Cells(2).Value
                    xlWorkSheet5.Cells(z, 5) = Install_grid.Rows(i).Cells(3).Value
                    xlWorkSheet5.Cells(z, 6) = Install_grid.Rows(i).Cells(4).Value
                    xlWorkSheet5.Cells(z, 7) = Install_grid.Rows(i).Cells(5).Value
                    xlWorkSheet5.Cells(z, 8) = Install_grid.Rows(i).Cells(6).Value
                    xlWorkSheet5.Cells(z, 9) = Install_grid.Rows(i).Cells(7).Value
                    xlWorkSheet5.Cells(z, 10) = Install_grid.Rows(i).Cells(8).Value
                    xlWorkSheet5.Cells(z, 11) = Install_grid.Rows(i).Cells(9).Value
                    xlWorkSheet5.Cells(z, 12) = Install_grid.Rows(i).Cells(10).Value
                    xlWorkSheet5.Cells(z, 13) = Install_grid.Rows(i).Cells(11).Value
                    xlWorkSheet5.Cells(z, 14) = Install_grid.Rows(i).Cells(12).Value

                    z = z + 1
                Next


                '-----------------------------------------------

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                wb.Close(False)


                Marshal.ReleaseComObject(xlApp)
                Label4.Visible = False

                MessageBox.Show("Quote generated successfully!")


            Catch ex As Exception
                MessageBox.Show("File not found or corrupted")
            End Try

        End If

    End Sub

    Sub Estop_feet()
        '-- convert to meters when textbox change

        Dim m As String : m = t_feet_box.Text
        Dim av As Double : av = 0

        If IsNumeric(t_feet_box.Text) = True Then
            m = Math.Round(t_feet_box.Text / 3.28, 2)
        End If

        meters_l.Text = "meters: " & m

        If IsNumeric(TextBox2.Text) = True And IsNumeric(m) = True Then
            av = Math.Round(m / TextBox2.Text, 2)
            Label7.Text = "Average Distance: " & av
        End If

        If av < 6 Then
            my_av.Text = "2"
        ElseIf av > 6 And av < 11 Then
            my_av.Text = "5"
        ElseIf av > 11 And av < 16 Then
            my_av.Text = "10"
        ElseIf av > 16 And av < 21 Then
            my_av.Text = "15"
        ElseIf av > 21 And av < 35 Then
            my_av.Text = "20"
        Else
            my_av.Text = "err"
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles t_feet_box.TextChanged
        '-- convert to meters when textbox change
        Call Estop_feet()
        Call cables_percentage()


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        '-- convert to meters when textbox change
        Call Estop_feet()
        Call cables_percentage()

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Call cables_percentage()

    End Sub

    Sub cables_percentage()

        If IsNumeric(TextBox3.Text) = True And IsNumeric(TextBox2.Text) = True Then

            c_l.Text = TextBox3.Text * TextBox2.Text

            If IsNumeric(my_av.Text) = True Then
                m_l.Text = (c_l.Text * my_av.Text)

                r_d.Text = If(IsNumeric(t_feet_box.Text), Math.Round(t_feet_box.Text / 3.28, 2), 0) - m_l.Text

            End If
        End If

        Call cable_cal()

    End Sub

    Sub cable_cal()

        If donotcompute = True Then

            For i = 0 To cables_grid.Rows.Count - 1

                cables_grid.Rows(i).Cells(2).Value = Math.Ceiling(((If(IsNumeric(r_d.Text) = True, r_d.Text, 0) / cables_grid.Rows(i).Cells(0).Value)) * cables_grid.Rows(i).Cells(1).Value)
                cables_grid.Rows(i).Cells(3).Value = cables_grid.Rows(i).Cells(0).Value * cables_grid.Rows(i).Cells(2).Value
            Next

            Dim total_c As Double : total_c = 0
            Dim total_m As Double : total_m = 0


            For i = 0 To cables_grid.Rows.Count - 1
                total_c = total_c + If(IsNumeric(cables_grid.Rows(i).Cells(2).Value), cables_grid.Rows(i).Cells(2).Value, 0)
                total_m = total_m + If(IsNumeric(cables_grid.Rows(i).Cells(3).Value), cables_grid.Rows(i).Cells(3).Value, 0)
            Next

            Label18.Text = "Total Cables: " & c_l.Text + total_c
            Label19.Text = "Total m: " & m_l.Text + total_m
        End If
    End Sub

    Private Sub cables_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles cables_grid.CellValueChanged
        Call cable_cal()
    End Sub

    '----- alloc m12 estops cables -----
    Sub Estops_alloc(ESB As Double)

        If ESB > 0 Then

            TextBox2.Text = ESB
            Call Estop_feet()
            Call cables_percentage()

            '--20   V15-G-S-YE20M-PUR-A-V15-G
            If IsNumeric(cables_grid.Rows(0).Cells(2).Value) = True Then
                If cables_grid.Rows(0).Cells(2).Value > 0 Then
                    scapegoat.Rows.Add("V15-G-S-YE20M-PUR-A-V15-G", cables_grid.Rows(0).Cells(2).Value, Panel_grid.Columns.Item(1).HeaderText, "M12_E")
                End If
            End If

            '--15  V15-G-S-YE15M-PUR-A-V15-G
            If IsNumeric(cables_grid.Rows(1).Cells(2).Value) = True Then
                If cables_grid.Rows(1).Cells(2).Value > 0 Then
                    scapegoat.Rows.Add("V15-G-S-YE15M-PUR-A-V15-G", cables_grid.Rows(1).Cells(2).Value, Panel_grid.Columns.Item(1).HeaderText, "M12_E")
                End If
            End If


            '--10  V15-G-S-YE10M-PUR-A-V15-G
            If IsNumeric(cables_grid.Rows(2).Cells(2).Value) = True Then
                If cables_grid.Rows(2).Cells(2).Value > 0 Then
                    scapegoat.Rows.Add("V15-G-S-YE10M-PUR-A-V15-G", cables_grid.Rows(2).Cells(2).Value, Panel_grid.Columns.Item(1).HeaderText, "M12_E")
                End If
            End If


            '--5  V15-G-S-YE5M-PUR-A-V15-G
            If IsNumeric(cables_grid.Rows(3).Cells(2).Value) = True Then
                If cables_grid.Rows(3).Cells(2).Value > 0 Then
                    scapegoat.Rows.Add("V15-G-S-YE5M-PUR-A-V15-G", cables_grid.Rows(3).Cells(2).Value, Panel_grid.Columns.Item(1).HeaderText, "M12_E")
                End If
            End If


            '--2  V15-G-S-YE2M-PUR-A-V15-G
            If IsNumeric(cables_grid.Rows(4).Cells(2).Value) = True Then
                If cables_grid.Rows(4).Cells(2).Value > 0 Then
                    scapegoat.Rows.Add("V15-G-S-YE2M-PUR-A-V15-G", cables_grid.Rows(4).Cells(2).Value, Panel_grid.Columns.Item(1).HeaderText, "M12_E")
                End If
            End If

            Call insert_into_mysql_dummy_Table()
            scapegoat.Rows.Clear()

        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ' Quote_manager.Visible = True
        '  Call populate_quote_manager(1)
        Quote_open.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call General_calculation()
    End Sub

    Sub FLA_load()

        fla_set_grid.Rows.Clear()
        fla_set_grid.Rows.Add(New String() {"FLA (Amps)"})
        '--- fill FLA table
        For i = fla_set_grid.Columns.Count - 1 To 1 Step -1
            fla_set_grid.Columns.RemoveAt(i)
        Next

        For i = 1 To Panel_grid.Columns.Count - 1
            fla_set_grid.Columns.Add(Panel_grid.Columns(i).HeaderText, Panel_grid.Columns(i).HeaderText)
        Next

        Dim fla_array = New ArrayList()
        Dim total_fla As Double : total_fla = 0
        Dim ms_panel As Double : ms_panel = 0

        '//////////////  iterate per sets //////////////////
        For j = 1 To Panel_grid.Columns.Count - 1

            ms_panel = 0
            total_fla = 0
            fla_array.Clear()

            'find number of MS panels in that set

            For i = 0 To Panel_grid.Rows.Count - 1
                If String.Equals(Panel_grid.Rows(i).Cells(0).Value.ToString, "MS_Panel (24x15x10 Green)") = True Then
                    ms_panel = If(IsNumeric(Panel_grid.Rows(i).Cells(j).Value) = True, Panel_grid.Rows(i).Cells(j).Value, 0)
                    'Exit For
                End If

                If String.Equals(Panel_grid.Rows(i).Cells(0).Value.ToString, "MS_Panel (24x16x10 Gray)") = True Then
                    ms_panel = ms_panel + If(IsNumeric(Panel_grid.Rows(i).Cells(j).Value) = True, Panel_grid.Rows(i).Cells(j).Value, 0)
                    ' Exit For
                End If
            Next


            For z = 0 To fla_data.Rows.Count - 1
                If String.Equals(fla_data.Rows(z).Item(2), Panel_grid.Columns(j).HeaderText) = True Then
                    total_fla = total_fla + (fla_data.Rows(z).Item(0) * fla_data.Rows(z).Item(1))

                    For k = 0 To fla_data.Rows(z).Item(0)
                        fla_array.Add(fla_data.Rows(z).Item(1))
                    Next
                End If
            Next

            'sort and descending order
            fla_array.Sort()
            fla_array.Reverse()


            If ms_panel > 0 Then

                For l = 0 To ms_panel
                    If l < fla_array.Count Then
                        total_fla = total_fla + (fla_array(l) * 0.25)
                    End If
                Next

                fla_set_grid.Rows(0).Cells(j).Value = Math.Round(total_fla, 2)

            Else
                fla_set_grid.Rows(0).Cells(j).Value = 0
            End If

        Next

    End Sub

    Private Sub fla_set_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles fla_set_grid.CellValueChanged

        Dim fla_t As Double : fla_t = 0

        For i = 0 To fla_set_grid.Columns.Count - 1
            If fla_set_grid.Rows.Count > 0 Then
                If IsNumeric(fla_set_grid.Rows(0).Cells(i).Value) = True Then
                    fla_t = fla_t + fla_set_grid.Rows(0).Cells(i).Value
                End If
            End If
        Next

        Label6.Text = "Total FLA: " & fla_t & " amps"

    End Sub


    Sub Feature_code_solve2(feature_desc As String, qty As Double, solution As String, type As String, set_name As String)

        Dim my_Feature_code As String : my_Feature_code = ""

        '--------------- Start collecting parts according to feature code ------------

        Dim feature_c = From feature_code In table_Feature_code Where feature_code.Field(Of String)("description") = feature_desc And feature_code.Field(Of String)("Solution") = solution Select feature_code

        If feature_c.Any() = True Then
            my_Feature_code = feature_c(0).Field(Of String)("Feature_code")
            bulk_t = bulk_t + feature_c(0).Field(Of Double)("labor_cost") * qty
            labor_t = labor_t + feature_c(0).Field(Of Double)("bulk_cost") * qty
        End If

        '    '------------ Get parts and qty ----------------

        Dim part_qty = From feature_part In table_feature_parts Where feature_part.Field(Of String)("Feature_code") = my_Feature_code And feature_part.Field(Of String)("solution") = solution And feature_part.Field(Of String)("type") = type Select feature_part

        For Each row As DataRow In part_qty
            If qty > 0 Then
                my_alloc_table.Rows.Add(row.Item("part_name"), "", "", "", "", Math.Ceiling(row.Item("qty")) * qty, "", "", row.Item("type"))
            End If
        Next


        '    '================== FLA calculation ========================
        If String.Equals(type, "Panel") = True Or String.Equals(type, "Control Panel") = True Then

            Dim fla_return = From fla In table_f_dimensions Where fla.Field(Of String)("feature_code") = my_Feature_code And fla.Field(Of String)("Solution") = solution Select fla

            If fla_return.Any() = True Then
                fla_data.Rows.Add(qty, fla_return(0).Field(Of Double)("FLA"), set_name)
            End If

        End If

        '    '----------------- feature code formulas -----------------
        If String.Equals(type, "Panel") = True Then

            Dim formula_t = From formu In table_call_feature Where formu.Field(Of String)("feature_code") = my_Feature_code And formu.Field(Of String)("solution") = solution Select formu


            ' If formula_t.Any() = True Then
            For Each row As DataRow In formula_t
                Call Wired_feature_codes2(row.Item("formula"), qty, set_name)
                ' End If
            Next
        End If

    End Sub



    Sub M12_process2()

        'Calculate M12 and M12E cables
        'calculate the number of M12 cables we need
        Dim inputs As Integer : inputs = 0
        Dim outputs As Integer : outputs = 0
        Dim motion As Integer : motion = 0

        Dim limit_feet As Double : limit_feet = 15
        avoid_change_m12 = False

        Dim n_rows As Double : n_rows = 0
        Dim ran_feet As Integer : ran_feet = 3
        M12_grid.Rows.Clear()
        M12_ES_grid.Rows.Clear()

        'calculate total inputs, output, motion

        For i = 1 To Field_grid.Columns.Count - 1
            inputs = inputs + Count_IO(Field_grid.Columns(i).HeaderText, 1)
            outputs = outputs + Count_IO(Field_grid.Columns(i).HeaderText, 2)
            motion = motion + Count_IO(Field_grid.Columns(i).HeaderText, 3)
        Next

        inputs_io = inputs
        outputs_io = outputs
        motion_io = motion



        n_rows = inputs + outputs + motion
        n_rows = n_rows + n_rows * (If(IsNumeric(percentage_io.Text) = True, percentage_io.Text, 0) / 100)

        'write io labels
        total_io_l.Text = "Total IO: " & n_rows
        inputs_io_l.Text = "Digital Inputs: " & inputs
        output_io_l.Text = "Digital Outputs: " & outputs
        motion_io_l.Text = "Motion Outputs: " & motion

        If n_rows > 0 Then
            If RIO_press = False Then
                rio_label_w.Text = "Please DO NOT forget to allocate RIO Modules"
            Else
                rio_label_w.Text = ""
            End If
        End If

        For i = 0 To n_rows - 1
            If ran_feet > limit_feet Then
                ran_feet = 3
            End If
            M12_grid.Rows.Add(New String() {ran_feet})
            ran_feet = ran_feet + 3
        Next

        Dim fixed_l(5) As Double  'this will contain the qty of cable M12
        fixed_l(0) = 0
        fixed_l(1) = 0
        fixed_l(2) = 0
        fixed_l(3) = 0
        fixed_l(4) = 0
        fixed_l(5) = 0



        Dim temp As Double : temp = 0
        Dim temp_l(5) As Double 'this array is like fixed_l but doesnt accumulate value

        For i = 0 To M12_grid.Rows.Count - 1

            'clear temp_l
            temp_l(0) = 0 : temp_l(1) = 0 : temp_l(2) = 0 : temp_l(3) = 0 : temp_l(4) = 0 : temp_l(5) = 0


            fixed_l(5) = fixed_l(5) + Math.Round(M12_grid.Rows(i).Cells(0).Value / 65.6)
            temp_l(5) = Math.Round(M12_grid.Rows(i).Cells(0).Value / 65.6)

            temp = M12_grid.Rows(i).Cells(0).Value Mod 65.6

            If temp <= 6.56 Then
                fixed_l(0) = fixed_l(0) + 1
                temp_l(0) = 1
            End If

            If temp <= 9.84 And temp > 6.56 Then
                fixed_l(1) = fixed_l(1) + 1
                temp_l(1) = 1
            End If

            If temp <= 16.4 And temp > 9.84 Then
                fixed_l(2) = fixed_l(2) + 1
                temp_l(2) = 1
            End If

            If temp <= 32.8 And temp > 16.4 Then
                fixed_l(3) = fixed_l(3) + 1
                temp_l(3) = 1
            End If

            If temp <= 49.2 And temp > 32.8 Then
                fixed_l(4) = fixed_l(4) + 1
                temp_l(4) = 1
            End If

            M12_grid.Rows(i).Cells(1).Value = Math.Round(temp_l(0) * 6.56 + temp_l(1) * 9.84 + temp_l(2) * 16.4 + temp_l(3) * 32.8 + temp_l(4) * 49.2 + temp_l(5) * 65.6, 1)

        Next

        For j = 0 To 5
            fixed_l(j) = Math.Ceiling(fixed_l(j) * 1.1)
        Next

        ''------------ M12_ES---------------
        'Dim ran_feet_es As Double : ran_feet_es = 3
        'Dim count_rows As Double : count_rows = 0


        'count_rows = estops_count()

        'For i = 0 To count_rows - 1
        '    If ran_feet_es > 55 Then
        '        ran_feet_es = 3
        '    End If
        '    M12_ES_grid.Rows.Add(New String() {ran_feet_es})

        '    If ran_feet_es > 44 Then
        '        ran_feet_es = ran_feet_es + 5
        '    Else
        '        ran_feet_es = ran_feet_es + 3
        '    End If
        'Next

        'Dim fixed_l2(5) As Double  'this will contain the qty of cable M12
        'fixed_l2(0) = 0
        'fixed_l2(1) = 0
        'fixed_l2(2) = 0
        'fixed_l2(3) = 0
        'fixed_l2(4) = 0
        'fixed_l2(5) = 0



        'Dim temp2 As Double : temp2 = 0
        'Dim temp_l2(5) As Double 'this array is like fixed_l but doesnt accumulate value

        'For i = 0 To M12_ES_grid.Rows.Count - 1

        '    'clear temp_l
        '    temp_l2(0) = 0 : temp_l2(1) = 0 : temp_l2(2) = 0 : temp_l2(3) = 0 : temp_l2(4) = 0 : temp_l2(5) = 0


        '    fixed_l2(5) = fixed_l2(5) + Math.Round(M12_ES_grid.Rows(i).Cells(0).Value / 65.6)
        '    temp_l2(5) = Math.Round(M12_ES_grid.Rows(i).Cells(0).Value / 65.6)

        '    temp2 = M12_ES_grid.Rows(i).Cells(0).Value Mod 65.6

        '    If temp2 <= 6.56 Then
        '        fixed_l2(0) = fixed_l2(0) + 1
        '        temp_l2(0) = 1
        '    End If

        '    If temp2 <= 9.84 And temp2 > 6.56 Then
        '        fixed_l2(1) = fixed_l2(1) + 1  '0 index
        '        temp_l2(1) = 1   '0 index
        '    End If

        '    If temp2 <= 16.4 And temp2 > 9.84 Then
        '        fixed_l2(2) = fixed_l2(2) + 1
        '        temp_l2(2) = 1
        '    End If

        '    If temp2 <= 32.8 And temp2 > 16.4 Then
        '        fixed_l2(3) = fixed_l2(3) + 1
        '        temp_l2(3) = 1
        '    End If

        '    If temp2 <= 49.2 And temp2 > 32.8 Then
        '        fixed_l2(4) = fixed_l2(4) + 1
        '        temp_l2(4) = 1
        '    End If

        '    M12_ES_grid.Rows(i).Cells(1).Value = Math.Round(temp_l2(0) * 6.56 + temp_l2(1) * 9.84 + temp_l2(2) * 16.4 + temp_l2(3) * 32.8 + temp_l2(4) * 49.2 + temp_l2(5) * 65.6, 1)

        'Next

        'For j = 0 To 5
        '    fixed_l2(j) = Math.Ceiling(fixed_l2(j) * 1.1)
        'Next

        '////////////////////   fill M12 Cables list  //////////


        M12_cables("V15-G-2M-PUR-V15-G") = fixed_l(0) + fixed_l(1)
        M12_cables("V15-G-5M-PUR-V15-G") = fixed_l(2)
        M12_cables("V15-G-10M-PUR-V15-G") = fixed_l(3)
        M12_cables("V15-G-15M-PUR-V15-G") = fixed_l(4)
        M12_cables("V15-G-20M-PUR-V15-G") = fixed_l(5)

        'M12_ES_cables("V15-G-S-YE2M-PUR-A-V15-G") = fixed_l2(0)
        'M12_ES_cables("7000-40041-0150300") = fixed_l2(1)
        'M12_ES_cables("V15-G-S-YE5M-PUR-A-V15-G") = fixed_l2(2)
        'M12_ES_cables("V15-G-S-YE10M-PUR-A-V15-G") = fixed_l2(3)
        'M12_ES_cables("V15-G-S-YE15M-PUR-A-V15-G") = fixed_l2(4)
        'M12_ES_cables("V15-G-S-YE20M-PUR-A-V15-G") = fixed_l2(5)



        For Each kvp As KeyValuePair(Of String, Double) In M12_cables.ToArray
            If CType(M12_cables(kvp.Key), Double) > 0 And Panel_grid.Columns.Count > 1 Then
                my_alloc_table.Rows.Add(kvp.Key, "", "", "", "", M12_cables(kvp.Key), "", "", "M12")
            End If
        Next


        'insert m12 into dummy datatable
        'For Each kvp As KeyValuePair(Of String, Double) In M12_ES_cables.ToArray
        '    If CType(M12_ES_cables(kvp.Key), Double) > 0 And Panel_grid.Columns.Count > 1 Then
        '        my_alloc_table.Rows.Add(kvp.Key, "", "", "", "", M12_ES_cables(kvp.Key), "", "", "M12_E")
        '    End If
        'Next


        avoid_change_m12 = True

    End Sub

    Sub Estops_alloc2(ESB As Double)

        If ESB > 0 Then

            TextBox2.Text = ESB
            'Call Estop_feet()
            'Call cables_percentage()

            ''------ old way of cal ESB with percenatages ---------

            ''--20   V15-G-S-YE20M-PUR-A-V15-G
            'If IsNumeric(cables_grid.Rows(0).Cells(2).Value) = True Then
            '    If cables_grid.Rows(0).Cells(2).Value > 0 Then
            '        my_alloc_table.Rows.Add("V15-G-S-YE20M-PUR-A-V15-G", "", "", "", "", cables_grid.Rows(0).Cells(2).Value, "", "", "M12_E")
            '    End If
            'End If

            ''--15  V15-G-S-YE15M-PUR-A-V15-G
            'If IsNumeric(cables_grid.Rows(1).Cells(2).Value) = True Then
            '    If cables_grid.Rows(1).Cells(2).Value > 0 Then
            '        my_alloc_table.Rows.Add("V15-G-S-YE15M-PUR-A-V15-G", "", "", "", "", cables_grid.Rows(1).Cells(2).Value, "", "", "M12_E")
            '    End If
            'End If


            ''--10  V15-G-S-YE10M-PUR-A-V15-G
            'If IsNumeric(cables_grid.Rows(2).Cells(2).Value) = True Then
            '    If cables_grid.Rows(2).Cells(2).Value > 0 Then
            '        my_alloc_table.Rows.Add("V15-G-S-YE10M-PUR-A-V15-G", "", "", "", "", cables_grid.Rows(2).Cells(2).Value, "", "", "M12_E")

            '    End If
            'End If


            ''--5  V15-G-S-YE5M-PUR-A-V15-G
            'If IsNumeric(cables_grid.Rows(3).Cells(2).Value) = True Then
            '    If cables_grid.Rows(3).Cells(2).Value > 0 Then
            '        my_alloc_table.Rows.Add("V15-G-S-YE5M-PUR-A-V15-G", "", "", "", "", cables_grid.Rows(3).Cells(2).Value, "", "", "M12_E")

            '    End If
            'End If


            ''--2  V15-G-S-YE2M-PUR-A-V15-G
            'If IsNumeric(cables_grid.Rows(4).Cells(2).Value) = True Then
            '    If cables_grid.Rows(4).Cells(2).Value > 0 Then
            '        my_alloc_table.Rows.Add("V15-G-S-YE2M-PUR-A-V15-G", "", "", "", "", cables_grid.Rows(4).Cells(2).Value, "", "", "M12_E")

            '    End If
            'End If

            '----- simple way of cal ES cables -------
            my_alloc_table.Rows.Add("V15-G-S-YE20M-PUR-A-V15-G", "", "", "", "", Math.Ceiling(ESB), "", "", "M12_E")
            my_alloc_table.Rows.Add("V15-G-S-YE15M-PUR-A-V15-G", "", "", "", "", Math.Ceiling(ESB * 0.15), "", "", "M12_E")
            my_alloc_table.Rows.Add("V15-G-S-YE10M-PUR-A-V15-G", "", "", "", "", Math.Ceiling(ESB * 0.25), "", "", "M12_E")
            my_alloc_table.Rows.Add("V15-G-S-YE5M-PUR-A-V15-G", "", "", "", "", Math.Ceiling(ESB * 0.25), "", "", "M12_E")
            my_alloc_table.Rows.Add("V15-G-S-YE2M-PUR-A-V15-G", "", "", "", "", Math.Ceiling(ESB * 0.35), "", "", "M12_E")


        End If
    End Sub

    Sub Create_Main_table2()

        Try

            'get vendor and cost
            Dim cmd3 As New MySqlCommand
            For i = 0 To my_alloc_table.Rows.Count - 1

                Dim part_n As String : part_n = my_alloc_table.Rows(i).Item(0).ToString
                Dim allparts = From parts In table_parts Where parts.Field(Of String)("Part_Name") = part_n Select parts

                For Each row As DataRow In allparts
                    my_alloc_table.Rows(i).Item(1) = row.Item("Part_Description")  'part desc
                    my_alloc_table.Rows(i).Item(2) = row.Item("Manufacturer")  'manuf
                    my_alloc_table.Rows(i).Item(3) = row.Item("Primary_Vendor")   'vendor
                    my_alloc_table.Rows(i).Item(7) = row.Item("MFG_type") 'mfg type
                    my_alloc_table.Rows(i).Item(9) = row.Item("Legacy_ADA_Number")   'ADA
                Next

                'get latest cost of part or assembly
                If my_assemblies.contains(part_n) = True Then
                    my_alloc_table.Rows(i).Item(4) = "$" & Cost_of_Assem2(part_n)
                Else
                    my_alloc_table.Rows(i).Item(4) = "$" & Get_Latest_Cost2(part_n, my_alloc_table.Rows(i).Item(3))
                End If
                'subtotal calculation
                my_alloc_table.Rows(i).Item(6) = my_alloc_table.Rows(i).Item(4) * my_alloc_table.Rows(i).Item(5)
            Next


            If Allocation_parts.Visible = True Then
                Allocation_parts.alloc_grid.Rows.Clear()

                For i = 0 To my_alloc_table.Rows.Count - 1
                    If CType(my_alloc_table.Rows(i).Item(5).ToString, Double) > 0 Then
                        Allocation_parts.alloc_grid.Rows.Add(New String() {my_alloc_table.Rows(i).Item(0).ToString, my_alloc_table.Rows(i).Item(1).ToString, my_alloc_table.Rows(i).Item(3).ToString, my_alloc_table.Rows(i).Item(5).ToString, my_alloc_table.Rows(i).Item(4).ToString, my_alloc_table.Rows(i).Item(6).ToString, my_alloc_table.Rows(i).Item(7).ToString})
                    End If
                Next


            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Function Get_Latest_Cost2(part As String, vendor As String) As Decimal

        Get_Latest_Cost2 = 0

        Dim costs_r = From cost In table_vendors Where cost.Field(Of String)("Part_Name") = part And cost.Field(Of String)("Vendor_Name") = vendor Select cost Order By cost.Field(Of Date)("Purchase_Date")

        For Each row As DataRow In costs_r
            Get_Latest_Cost2 = Decimal.Round(row.Item("cost"), 2, MidpointRounding.AwayFromZero)
        Next


    End Function

    Function Cost_of_Assem2(assem As String) As Decimal

        Cost_of_Assem2 = 0

        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("primary_vendor", GetType(String))
        datatable.Columns.Add("total_cost", GetType(Double))

        Dim labor_c As Double : labor_c = 0
        Dim bulk_c As Double : bulk_c = 0
        Dim mat_c As Double : mat_c = 0

        Dim mixdata = From parts In table_parts Join adv In table_adv On parts.Field(Of String)("Part_Name") Equals adv.Field(Of String)("Part_Name") Where adv.Field(Of String)("Legacy_ADA") = assem Select parts = parts.Field(Of String)("Part_Name"), qty = adv.Field(Of Double)("Qty"), primary_vendor = parts.Field(Of String)("Primary_Vendor")

        For Each row In mixdata
            datatable.Rows.Add(row.parts, row.qty, row.primary_vendor, 0)
        Next

        For i = 0 To datatable.Rows.Count - 1
            datatable.Rows(i).Item(3) = Get_Latest_Cost2(datatable.Rows(i).Item(0), datatable.Rows(i).Item(2)) * datatable.Rows(i).Item(1)

            mat_c = mat_c + datatable.Rows(i).Item(3)
        Next

        '--get labor and bulk

        Dim bulk_lab = From bl In table_assem Where bl.Field(Of String)("Legacy_ADA_Number") = assem Select bl

        For Each row As DataRow In bulk_lab
            labor_c = row.Item("Labor_Cost")
            bulk_c = row.Item("Bulk_Cost")
        Next

        Cost_of_Assem2 = labor_c + bulk_c + mat_c

    End Function

    Sub Wired_feature_codes2(feature_code As String, qty As Integer, set_n As String)

        'These are the hardwire feature codes that process the qty value and are called by other feature code
        Dim myqty As Double : myqty = 0

        Select Case feature_code

            Case "ES-BUS"
                'process qty
                myqty = Math.Ceiling(qty / 3)

                my_alloc_table.Rows.Add("FSFD 4.5-0.5", "", "", "", "", myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("FKFD 4.5-0.5", "", "", "", "", myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("3210567", "", "", "", "", 2 * myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("3211634", "", "", "", "", myqty, "", "", "Panel")


            Case "ES-BUS-P"
                'process qty
                myqty = Panel_grid.Columns.Count - 1

                my_alloc_table.Rows.Add("FSFD 4.5-0.5", "", "", "", "", myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("FKFD 4.5-0.5", "", "", "", "", myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("3210567", "", "", "", "", 2 * myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("3211634", "", "", "", "", myqty, "", "", "Panel")


            Case "SafeRly"

                'process qty
                myqty = Math.Ceiling(qty / 3)
                my_alloc_table.Rows.Add("P7SA-10F", "", "", "", "", myqty, "", "", "Panel")
                my_alloc_table.Rows.Add("G7SA-3A1B-DC24", "", "", "", "", myqty, "", "", "Panel")


            Case "MS-4-DE1"
                'process qty
                myqty = Math.Ceiling(qty / 3)
                my_alloc_table.Rows.Add("XTSC010BBTB", "", "", "", "", myqty, "", "", "Panel")


            Case "MDR_PS"
                Dim qty_40 As Double : qty_40 = Math.Ceiling(qty / 40)
                Dim qty_20 As Double : qty_20 = Math.Ceiling((qty - (qty_40 * 40)) / 20)


                my_alloc_table.Rows.Add("QT20.241", "", "", "", "", qty_20, "", "", "Panel")
                my_alloc_table.Rows.Add("FAZ-C20/1-SP", "", "", "", "", qty_20, "", "", "Panel")
                my_alloc_table.Rows.Add("3211813", "", "", "", "", qty_20, "", "", "Panel")
                my_alloc_table.Rows.Add("XTSC010BBTB", "", "", "", "", Math.Ceiling(qty_20 / 6), "", "", "Panel")

                my_alloc_table.Rows.Add("QT20.241", "", "", "", "", qty_40, "", "", "Panel")
                my_alloc_table.Rows.Add("FAZ-C20/1-SP", "", "", "", "", 2 * qty_40, "", "", "Panel")
                my_alloc_table.Rows.Add("3211813", "", "", "", "", 2 * qty_40, "", "", "Panel")
                my_alloc_table.Rows.Add("XTSC010BBTB", "", "", "", "", Math.Ceiling(qty_20 / 4), "", "", "Panel")

            Case "Class2"

                Dim qty_4 As Double : qty_4 = ((qty + 3) / 4) - 2 * ((qty + 3) / 8)
                Dim qty_8 As Double : qty_8 = ((qty + 3) / 8)

                '--channel 4


                my_alloc_table.Rows.Add("QT20.241", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("9000-41064-0400000", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("3000715", "", "", "", "", qty_4 * 2, "", "", "Panel")
                my_alloc_table.Rows.Add("3211634", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("3030161", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("3211813", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("3048519", "", "", "", "", qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("ATDR3 600V CC TD", "", "", "", "", 3 * qty_4, "", "", "Panel")
                my_alloc_table.Rows.Add("FKFD 4.5-0.5", "", "", "", "", 4 * qty_4, "", "", "Panel")

                '--channel 8

                my_alloc_table.Rows.Add("QT20.241", "", "", "", "", qty_8, "", "", "Panel")
                my_alloc_table.Rows.Add("9000-41064-0400000", "", "", "", "", qty_8, "", "", "Panel")
                my_alloc_table.Rows.Add("3000715", "", "", "", "", qty_8 * 4, "", "", "Panel")
                my_alloc_table.Rows.Add("3030187", "", "", "", "", qty_8, "", "", "Panel")
                my_alloc_table.Rows.Add("3211634", "", "", "", "", qty_8, "", "", "Panel")
                my_alloc_table.Rows.Add("3211813", "", "", "", "", qty_8 * 2, "", "", "Panel")
                my_alloc_table.Rows.Add("3048519", "", "", "", "", qty_8 * 2, "", "", "Panel")
                my_alloc_table.Rows.Add("3ATDR3 600V CC TD", "", "", "", "", qty_8 * 3, "", "", "Panel")
                my_alloc_table.Rows.Add("FKFD 4.5-0.5", "", "", "", "", qty_8 * 8, "", "", "Panel")
        End Select



    End Sub

    Private Sub myQuote_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

        'call panel wiz ctrl + P
        If (e.Control AndAlso e.KeyCode = Keys.P) Then
            Set_wizard.ShowDialog()
        End If
    End Sub

    Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MessageBox.Show(Cost_of_Assem2("ADA-ASM-ANN101"))
    End Sub

    Sub totals_bt()

        '--- calculate totals
        'cal totals in quote form
        Dim total_cost As Double : total_cost = 0
        Dim final_price As Double : final_price = 0

        Dim ADA_Panels As Double : ADA_Panels = 0
        Dim ADA_Panels_f As Double : ADA_Panels_f = 0
        Dim field_item As Double : field_item = 0
        Dim field_item_f As Double : field_item_f = 0
        Dim in_office_l As Double : in_office_l = 0
        Dim in_office_lf As Double : in_office_lf = 0
        Dim onsite_l As Double : onsite_l = 0
        Dim onsite_lf As Double : onsite_lf = 0
        Dim start_up As Double : start_up = 0
        Dim start_upf As Double : start_upf = 0
        Dim electrical As Double : electrical = 0
        Dim electricalf As Double : electricalf = 0
        Dim shipping As Double : shipping = 0
        Dim shippingf As Double : shippingf = 0


        'first total sum
        For i = 1 To 15
            ADA_Panels = ADA_Panels + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            ADA_Panels_f = ADA_Panels_f + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(16).Cells(6).Value = "$" & ADA_Panels.ToString("N")
        Quote_grid.Rows(16).Cells(9).Value = "$" & ADA_Panels_f.ToString("N")

        total_cost = total_cost + ADA_Panels
        final_price = final_price + ADA_Panels_f

        'second
        For i = 19 To 42
            field_item = field_item + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            field_item_f = field_item_f + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(43).Cells(6).Value = "$" & field_item.ToString("N")
        Quote_grid.Rows(43).Cells(9).Value = "$" & field_item_f.ToString("N")


        total_cost = total_cost + field_item
        final_price = final_price + field_item_f

        'third
        For i = 45 To 56
            in_office_l = in_office_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            in_office_lf = in_office_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(57).Cells(6).Value = "$" & in_office_l.ToString("N")
        Quote_grid.Rows(57).Cells(9).Value = "$" & in_office_lf.ToString("N")


        total_cost = total_cost + in_office_l
        final_price = final_price + in_office_lf

        'fourth
        For i = 59 To 68
            onsite_l = onsite_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            onsite_lf = onsite_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(69).Cells(6).Value = "$" & onsite_l.ToString("N")
        Quote_grid.Rows(69).Cells(9).Value = "$" & onsite_lf.ToString("N")

        total_cost = total_cost + onsite_l
        final_price = final_price + onsite_lf

        'fifth
        For i = 72 To 75
            start_up = start_up + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            start_upf = start_upf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(76).Cells(6).Value = "$" & start_up.ToString("N")
        Quote_grid.Rows(76).Cells(9).Value = "$" & start_upf.ToString("N")

        total_cost = total_cost + start_up
        final_price = final_price + start_upf

        'sixth
        For i = 79 To 82
            electrical = electrical + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            electricalf = electricalf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(83).Cells(6).Value = "$" & electrical.ToString("N")
        Quote_grid.Rows(83).Cells(9).Value = "$" & electricalf.ToString("N")

        total_cost = total_cost + electrical
        final_price = final_price + electricalf

        'seventh
        For i = 86 To 88
            shipping = shipping + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
            shippingf = shippingf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
        Next
        Quote_grid.Rows(89).Cells(6).Value = "$" & shipping.ToString("N")
        Quote_grid.Rows(89).Cells(9).Value = "$" & shippingf.ToString("N")

        total_cost = total_cost + shipping
        final_price = final_price + shippingf

        Quote_grid.Rows(101).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(101).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(107).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(107).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(123).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(123).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(130).Cells(6).Value = "$" & Math.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
        Quote_grid.Rows(130).Cells(9).Value = "$" & Math.Round(final_price, 2, MidpointRounding.AwayFromZero).ToString("N")


        '--- cal expenses
        Dim days_site As Double : days_site = 0

        For i = 60 To 68
            days_site = days_site + If(IsNumeric(Quote_grid.Rows(i).Cells(2).Value) = True, Quote_grid.Rows(i).Cells(2).Value, 0)
        Next

        Dim on_site As Double : on_site = 1

        If IsNumeric(Quote_grid.Rows(138).Cells(2).Value) = True Then
            If Quote_grid.Rows(138).Cells(2).Value <= 0 Then
                on_site = 1
            Else
                on_site = Quote_grid.Rows(138).Cells(2).Value  'hours / day onsite
            End If
        End If

        Quote_grid.Rows(146).Cells(2).Value = Math.Ceiling(days_site / on_site)  'DAYS ON SITE
        '--------------------------

        '---- cal exp
        If (IsNumeric(Quote_grid.Rows(146).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(139).Cells(2).Value())) Then
            Quote_grid.Rows(146).Cells(2).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(146).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(146).Cells(2).Value()) = False, 0, Quote_grid.Rows(146).Cells(2).Value())
            Quote_grid.Rows(139).Cells(2).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(139).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(139).Cells(2).Value()) = False, 0, Quote_grid.Rows(139).Cells(2).Value())

            If Quote_grid.Rows(139).Cells(2).Value() > 0 Then
                Quote_grid.Rows(147).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value())  '#trips
                Quote_grid.Rows(151).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value())  '#Airfare
                Quote_grid.Rows(148).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value()) * If(String.IsNullOrEmpty(Quote_grid.Rows(140).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(140).Cells(2).Value()) = False, 0, Quote_grid.Rows(140).Cells(2).Value()) 'travel time
                Quote_grid.Rows(149).Cells(2).Value() = Math.Ceiling(Quote_grid.Rows(146).Cells(2).Value() / Quote_grid.Rows(139).Cells(2).Value()) * If(String.IsNullOrEmpty(Quote_grid.Rows(141).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(141).Cells(2).Value()) = False, 0, Quote_grid.Rows(141).Cells(2).Value()) 'mileage
                Quote_grid.Rows(150).Cells(2).Value() = (Quote_grid.Rows(147).Cells(2).Value() * If(String.IsNullOrEmpty(Quote_grid.Rows(142).Cells(2).Value()) = True Or IsNumeric(Quote_grid.Rows(142).Cells(2).Value()) = False, 0, Quote_grid.Rows(142).Cells(2).Value())) + Quote_grid.Rows(146).Cells(2).Value

                Quote_grid.Rows(72).Cells(2).Value() = Quote_grid.Rows(148).Cells(2).Value()
                Quote_grid.Rows(73).Cells(2).Value() = Quote_grid.Rows(149).Cells(2).Value()
                Quote_grid.Rows(74).Cells(2).Value() = Quote_grid.Rows(150).Cells(2).Value()
                Quote_grid.Rows(75).Cells(2).Value() = Quote_grid.Rows(151).Cells(2).Value()


            End If
        End If

        '----((((( TOTALS - By component
        '--ADA Panels
        Quote_grid.Rows(94).Cells(6).Value = "$" & ADA_Panels
        Quote_grid.Rows(94).Cells(9).Value = "$" & ADA_Panels_f

        '--- Field items
        Quote_grid.Rows(95).Cells(6).Value = "$" & field_item
        Quote_grid.Rows(95).Cells(9).Value = "$" & field_item_f

        '--- In office labor
        Quote_grid.Rows(96).Cells(6).Value = "$" & in_office_l
        Quote_grid.Rows(96).Cells(9).Value = "$" & in_office_lf

        '---- onsite labor
        Quote_grid.Rows(97).Cells(6).Value = "$" & onsite_l
        Quote_grid.Rows(97).Cells(9).Value = "$" & onsite_lf

        '--- start up expenses
        Quote_grid.Rows(98).Cells(6).Value = "$" & start_up
        Quote_grid.Rows(98).Cells(9).Value = "$" & start_upf

        '-------electrical installation
        Quote_grid.Rows(99).Cells(6).Value = "$" & electrical
        Quote_grid.Rows(99).Cells(9).Value = "$" & electricalf

        '----- shipping -------
        Quote_grid.Rows(100).Cells(6).Value = "$" & shipping
        Quote_grid.Rows(100).Cells(9).Value = "$" & shippingf

        '----(((((( TOTALS Sum all but installation

        '-------- hardware and labor ---
        Quote_grid.Rows(104).Cells(6).Value = "$" & ADA_Panels + field_item + in_office_l + onsite_l + start_up
        Quote_grid.Rows(104).Cells(9).Value = "$" & ADA_Panels_f + field_item_f + in_office_lf + onsite_lf + start_upf

        '-------- electrical installation --
        Quote_grid.Rows(105).Cells(6).Value = "$" & electrical
        Quote_grid.Rows(105).Cells(9).Value = "$" & electricalf

        '--------- shipping ------
        Quote_grid.Rows(106).Cells(6).Value = "$" & shipping
        Quote_grid.Rows(106).Cells(9).Value = "$" & shippingf

        '---- (((( TOTAL-- By type

        '--------- Project labor 
        Dim project_l As Double : project_l = 0
        Dim project_lf As Double : project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(110).Cells(6).Value = project_l
        Quote_grid.Rows(110).Cells(9).Value = project_lf

        '------- Project Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(111).Cells(6).Value = project_l
        Quote_grid.Rows(111).Cells(9).Value = project_lf

        '------- Project Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Project Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(112).Cells(6).Value = project_l
        Quote_grid.Rows(112).Cells(9).Value = project_lf

        '--------- Manufacturing Labor

        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Manufacturing Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(113).Cells(6).Value = project_l
        Quote_grid.Rows(113).Cells(9).Value = project_lf



        '--------- Manufacturing Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Manufacturing Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(114).Cells(6).Value = project_l
        Quote_grid.Rows(114).Cells(9).Value = project_lf

        '--------- Startup Labor
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(115).Cells(6).Value = project_l
        Quote_grid.Rows(115).Cells(9).Value = project_lf

        '--------- Startup Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(116).Cells(6).Value = project_l
        Quote_grid.Rows(116).Cells(9).Value = project_lf

        '--------- Startup Expenses
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Expenses") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(117).Cells(6).Value = project_l
        Quote_grid.Rows(117).Cells(9).Value = project_lf

        '--------- Startup Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Startup Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(118).Cells(6).Value = project_l
        Quote_grid.Rows(118).Cells(9).Value = project_lf

        '--------- Install Labor
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Labor") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(119).Cells(6).Value = project_l
        Quote_grid.Rows(119).Cells(9).Value = project_lf

        '--------- Install Materials
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Materials") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(120).Cells(6).Value = project_l
        Quote_grid.Rows(120).Cells(9).Value = project_lf

        '--------- Install expenses
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Expenses") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(121).Cells(6).Value = project_l
        Quote_grid.Rows(121).Cells(9).Value = project_lf

        '--------- Install Subcontract
        project_l = 0
        project_lf = 0

        For i = 1 To 131
            If String.Equals(Quote_grid.Rows(i).Cells(11).Value, "Install Subcontract") = True Then
                project_l = project_l + If(IsNumeric(Quote_grid.Rows(i).Cells(6).Value) = True, Quote_grid.Rows(i).Cells(6).Value, 0)
                project_lf = project_lf + If(IsNumeric(Quote_grid.Rows(i).Cells(9).Value) = True, Quote_grid.Rows(i).Cells(9).Value, 0)
            End If
        Next

        Quote_grid.Rows(122).Cells(6).Value = project_l
        Quote_grid.Rows(122).Cells(9).Value = project_lf


        '--- (((( TOTALS Summary ------
        Quote_grid.Rows(126).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(110).Cells(6).Value), Quote_grid.Rows(110).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(113).Cells(6).Value), Quote_grid.Rows(113).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(115).Cells(6).Value), Quote_grid.Rows(115).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(119).Cells(6).Value), Quote_grid.Rows(119).Cells(6).Value, 0))
        Quote_grid.Rows(126).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(110).Cells(9).Value), Quote_grid.Rows(110).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(113).Cells(9).Value), Quote_grid.Rows(113).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(115).Cells(9).Value), Quote_grid.Rows(115).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(119).Cells(9).Value), Quote_grid.Rows(119).Cells(9).Value, 0))

        Quote_grid.Rows(127).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(111).Cells(6).Value), Quote_grid.Rows(111).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(116).Cells(6).Value), Quote_grid.Rows(116).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(120).Cells(6).Value), Quote_grid.Rows(120).Cells(6).Value, 0))
        Quote_grid.Rows(127).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(111).Cells(9).Value), Quote_grid.Rows(111).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(116).Cells(9).Value), Quote_grid.Rows(116).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(120).Cells(9).Value), Quote_grid.Rows(120).Cells(9).Value, 0))

        Quote_grid.Rows(128).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(117).Cells(6).Value), Quote_grid.Rows(117).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(121).Cells(6).Value), Quote_grid.Rows(121).Cells(6).Value, 0))
        Quote_grid.Rows(128).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(117).Cells(9).Value), Quote_grid.Rows(117).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(121).Cells(9).Value), Quote_grid.Rows(121).Cells(9).Value, 0))

        Quote_grid.Rows(129).Cells(6).Value = "$" & (If(IsNumeric(Quote_grid.Rows(112).Cells(6).Value), Quote_grid.Rows(112).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(114).Cells(6).Value), Quote_grid.Rows(114).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(118).Cells(6).Value), Quote_grid.Rows(118).Cells(6).Value, 0) + If(IsNumeric(Quote_grid.Rows(122).Cells(6).Value), Quote_grid.Rows(122).Cells(6).Value, 0))
        Quote_grid.Rows(129).Cells(9).Value = "$" & (If(IsNumeric(Quote_grid.Rows(112).Cells(9).Value), Quote_grid.Rows(112).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(114).Cells(9).Value), Quote_grid.Rows(114).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(118).Cells(9).Value), Quote_grid.Rows(118).Cells(9).Value, 0) + If(IsNumeric(Quote_grid.Rows(122).Cells(9).Value), Quote_grid.Rows(122).Cells(9).Value, 0))


        For z = 110 To 122
            Quote_grid.Rows(z).Cells(6).Value = "$" & CType(Quote_grid.Rows(z).Cells(6).Value, Double).ToString("N")
            Quote_grid.Rows(z).Cells(9).Value = "$" & CType(Quote_grid.Rows(z).Cells(9).Value, Double).ToString("N")
        Next
    End Sub


    Sub BOM_refresh()
        '---------- refresh BOM in BOM tab ---------------

        PR_grid.Rows.Clear()


        Dim datatable As DataTable
        datatable = New DataTable
        datatable.Columns.Add("qty", GetType(String))
        datatable.Columns.Add("vendor", GetType(String))
        datatable.Columns.Add("parts", GetType(String))
        datatable.Columns.Add("desc", GetType(String))


        For i = 0 To my_alloc_table.Rows.Count - 1

            If isintable(my_alloc_table.Rows(i).Item(0).ToString) = True Then
                Call update_Qty(my_alloc_table.Rows(i).Item(0).ToString, my_alloc_table.Rows(i).Item(5).ToString)
            Else
                PR_grid.Rows.Add(New String() {my_alloc_table.Rows(i).Item(0).ToString, my_alloc_table.Rows(i).Item(1).ToString, my_alloc_table.Rows(i).Item(9).ToString, my_alloc_table.Rows(i).Item(2).ToString, my_alloc_table.Rows(i).Item(3).ToString, If(IsNumeric(my_alloc_table.Rows(i).Item(4).ToString) = True, my_alloc_table.Rows(i).Item(4).ToString, 0), If(IsNumeric(my_alloc_table.Rows(i).Item(5).ToString) = True, my_alloc_table.Rows(i).Item(5).ToString, 0), If(IsNumeric(my_alloc_table.Rows(i).Item(6).ToString) = True, my_alloc_table.Rows(i).Item(6).ToString, 0), my_alloc_table.Rows(i).Item(7).ToString})
            End If

        Next

        Call Total_rows()

    End Sub




    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call BOM_refresh()
    End Sub

    Private Sub PR_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles PR_grid.CellValueChanged
        Call Total_rows()
    End Sub

    Sub Total_rows()

        'Calculate total row number

        Dim total_parts As Double : total_parts = 0

        For i = 0 To PR_grid.Rows.Count - 1
            total_parts = total_parts + PR_grid.Rows(i).Cells(6).Value()
        Next

        Label10.Text = "# Of Parts: " & total_parts


        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(5).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = row.Cells(6).Value * row.Cells(5).Value
            End If
        Next

    End Sub

    Function Find_index(name As String) As Integer

        'Find and return row index of ADA part in datagrid

        Find_index = -1

        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If String.Compare(row.Cells.Item("Column10").Value.ToString, name) = 0 Then
                Find_index = row.Index
                Exit For
            End If
        Next

    End Function

    Function isintable(part_name As String) As Boolean
        isintable = False

        For i = 0 To PR_grid.Rows.Count - 1
            If String.Equals(part_name, PR_grid.Rows(i).Cells(0).Value) = True Then
                isintable = True
                Exit For
            End If
        Next

    End Function


    Sub update_Qty(part_name As String, qty As Double)

        For i = 0 To PR_grid.Rows.Count - 1
            If String.Equals(part_name, PR_grid.Rows(i).Cells(0).Value) = True Then
                PR_grid.Rows(i).Cells(6).Value = If(IsNumeric(PR_grid.Rows(i).Cells(6).Value), PR_grid.Rows(i).Cells(6).Value, 0) + qty
                Exit For
            End If
        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:B").ColumnWidth = 40
                xlWorkSheet.Range("C:C").ColumnWidth = 30
                xlWorkSheet.Range("D:I").ColumnWidth = 20
                xlWorkSheet.Range("A:I").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To PR_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = PR_grid.Columns(i).HeaderText
                    For j As Integer = 0 To PR_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = PR_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Purchase_Request.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("BOM exported successfully!")
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

        If TabControl1.SelectedTab Is TabPage10 Then

            If PR_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(PR_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage11 Then

            If compare_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(compare_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        '---------- Open old excel quote ---------------
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*ods;"

            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

                compare_grid.Rows.Clear()

                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)

                Dim i As Integer : i = 2

                Try
                    While (wb.Sheets("BOM").Cells(i, 1).Value IsNot Nothing)
                        compare_grid.Rows.Add(New String() {Trim(wb.Sheets("BOM").Cells(i, 1).value), wb.Sheets("BOM").Cells(i, 2).value, wb.Sheets("BOM").Cells(i, 7).value, 0, 0, wb.Sheets("BOM").Cells(i, 6).value, 0, 0, wb.Sheets("BOM").Cells(i, 8).value, 0, 0})
                        i = i + 1
                    End While

                    wb.Close(False)


                    '----------- copy current BOM ---------
                    If PR_grid.Rows.Count > 0 Then

                        Dim index As Integer : index = 0

                        For j = 0 To PR_grid.Rows.Count - 1

                            index = isintable_compare(PR_grid.Rows(j).Cells(0).Value.ToString)

                            If index > -1 Then  'part has been found
                                compare_grid.Rows(index).Cells(3).Value = PR_grid.Rows(j).Cells(6).Value
                                compare_grid.Rows(index).Cells(6).Value = PR_grid.Rows(j).Cells(5).Value
                                compare_grid.Rows(index).Cells(9).Value = PR_grid.Rows(j).Cells(7).Value

                            Else
                                compare_grid.Rows.Add(New String() {PR_grid.Rows(j).Cells(0).Value, PR_grid.Rows(j).Cells(1).Value, 0, PR_grid.Rows(j).Cells(6).Value, 0, 0, PR_grid.Rows(j).Cells(5).Value, 0, 0, PR_grid.Rows(j).Cells(7).Value, 0})
                            End If

                        Next
                    End If
                    '-------------------------------------


                    Call Total_compare_rows()
                    MessageBox.Show("Done")

                Catch ex As Exception
                    MessageBox.Show("Incorrect Quote format!")
                End Try

            End If
        End If

    End Sub

    Sub Total_compare_rows()

        'Calculate total row number

        Dim total_delta_Qty As Double : total_delta_Qty = 0
        Dim total_delta_sub As Double : total_delta_sub = 0

        For i = 0 To compare_grid.Rows.Count - 1
            total_delta_Qty = total_delta_Qty + compare_grid.Rows(i).Cells(4).Value()
            total_delta_sub = total_delta_sub + compare_grid.Rows(i).Cells(10).Value()
        Next

        Label14.Text = "Total Delta Qty: " & total_delta_Qty
        Label15.Text = "Total Delta Subtotal:  $" & Decimal.Round(total_delta_sub, 2, MidpointRounding.AwayFromZero).ToString("N")


        For Each row As DataGridViewRow In compare_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(3).Value) = True And IsNumeric(row.Cells(2).Value)) Then
                row.Cells(4).Value = CType(row.Cells(3).Value - row.Cells(2).Value, Double)


            End If

            If (IsNumeric(row.Cells(5).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = CType(row.Cells(6).Value - row.Cells(5).Value, Double)


            End If

            If (IsNumeric(row.Cells(9).Value) = True And IsNumeric(row.Cells(8).Value)) Then
                row.Cells(10).Value = Math.Round(CType(row.Cells(9).Value - row.Cells(8).Value, Double))

            End If
        Next



    End Sub

    Private Sub compare_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles compare_grid.CellValueChanged
        Call Total_compare_rows()

    End Sub

    Function isintable_compare(part_name As String) As Integer
        isintable_compare = -1

        For i = 0 To compare_grid.Rows.Count - 1
            If String.Equals(part_name, compare_grid.Rows(i).Cells(0).Value) = True Then
                isintable_compare = i
                Exit For
            End If
        Next

    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        '-- export compare
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:B").ColumnWidth = 40
                xlWorkSheet.Range("C:C").ColumnWidth = 30
                xlWorkSheet.Range("D:I").ColumnWidth = 20
                xlWorkSheet.Range("A:I").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To compare_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = compare_grid.Columns(i).HeaderText
                    For j As Integer = 0 To compare_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = compare_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Compare_quote.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Table exported successfully!")
            End If
        End If
    End Sub
End Class