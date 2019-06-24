Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ADA_Setup

    '--------- Creating Global setup variables----------------------

    Public Voltage As String
    Public efficiency As String
    Public Power_factor As String
    Public include_ADA_brackets As String
    Public Use_4pts_IO As String
    Public Use_8pts_IO As String
    Public Use_16pts_IO As String
    Public removal_msbox_t As String
    Public M23_at_IO_Receptacle As String
    Public M23_Bulkhead_at_ADA As String
    Public Use_4pts_IO_RB As String
    Public Use_8pts_IO_RB As String
    Public Use_16pts_IO_RB_Inputs As String
    Public Use_16pts_IO_RB_Outputs As String
    Public Single_Channel_IO_1_Input_per_RB As String
    Public Single_Channel_IO_1_Output_per_RB As String
    Public Splitter_Perc_dual_ch As String
    Public Valve_solenoid_adapter As String
    Public Smart_Wire_Darwin_Limited As String
    Public Smart_Wire_Darwin_Full As String
    Public EthernetIP_Full As String
    Public Percent_overage_inputs As String
    Public Percent_overage_nm_outputs As String
    Public Percent_overage_m_outputs As String

    '-----------  Variables use by the class    -----------------------------------------------------

    Public ADA_set_list  ' list that contains sets
    Public temp_name As String
    Public row_in As Integer ' this variable keep the row index of the PR grid selected when (Find Alternative part option is used)
    Public myQTY As Double 'store PR qty temporarily


    '---------- Create arrays that will contain inputs -----------------
    Public Motor_array(47, 3) As String
    Public inputs_array(18, 3) As String
    Public push_array(18, 3) As String
    Public lights_array(14, 3) As String
    Public field_array(11, 3) As String
    Public bus_array(11, 3) As String
    Public scanner_array(3, 3) As String
    Public brakes_array(14, 3) As String

    '------------- Store Total "inputs" (setup values) --------- CELL (BD)

    ' Public t_panel_present As Double
    Public t_NEMA12 As Double
    Public t_NEMA4X As Double
    Public t_PLC_Gatewat_PGW As Double
    Public t_control_logic As Double
    Public t_compact_logic As Double
    Public t_PLC_UPS As Double
    Public t_Amp_Mon As Double
    Public t_Remote_IO As Double
    Public t_total_current_group As Double
    '  Public t_amps_30 As Double
    Public t_mdr As Double
    Public t_ADA_NEMA12_BOX24_No_Starters As Double
    Public t_ADA_NEMA12_BOX30_No_Starters As Double
    Public t_motor_disconnect As Double
    Public t_Motor_array(47) As Double
    Public t_inputs_array(18) As Double
    Public t_push_array(18) As Double
    Public t_lights_array(14) As Double
    Public t_field_array(11) As Double
    Public t_bus_array(11) As Double
    Public t_scanner_array(3) As Double
    Public t_brakes_array(14) As Double

    '--------- Installation variables --------
    Public totales As Double
    Public labor As Double
    Public materials As Double
    Public expenses As Double
    Public subcontract As Double


    '---------- Horse Power Tables (Starter Panel tab) -------------------------
    '--- First index is horsepower, second is voltage

    Public Motor_horsepower(11, 2) As Double  'NFPA FLA for Motor Horse Power  
    Public VFD_horsepower(11, 2) As Double  'NFPA FLA for VFD Horse Power
    Public Starters_values(15) As Double  'amps * 0.85 * 1.25 in starter Panel tab

    '--------------------------------------------------
    Public value_to_index_480(50) As Integer  'going to store amps need. use to map value to index
    Public value_to_index_575(50) As Integer
    Public Power_supply_480(50, 3) As Double
    Public Power_supply_575(50, 3) As Double

    '------------------------------------------
    ' Cables
    Public M12_cables As New Dictionary(Of String, Double)
    Public M12_ES_cables As New Dictionary(Of String, Double)
    Public avoid_change_m12 As Boolean

    '-------------------------------------------------

    Private Sub ADA_Setup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '=================================  INITIALIZE  =======================================================
        '--------- Creating Global setup variables ----------------------

        Voltage = 480   'Initialize voltage as 480V
        efficiency = 0  'Initilialize efficiency
        Power_factor = 0  'Power factor
        include_ADA_brackets = 0 'Include ADA brackets
        Use_4pts_IO = 0 'Use 4pts  I/O?
        Use_8pts_IO = 0 'Use 8pts I/O?
        Use_16pts_IO = 0  'Use 16pts I/O?
        removal_msbox_t = 0 'I dont know what ??/ stand for
        M23_at_IO_Receptacle = 0 'M23 at I/O Receptacle?
        M23_Bulkhead_at_ADA = 0  'M23 Bulkhead at ADA?  (BOOL 1 or 0)
        Use_4pts_IO_RB = 0  'Use 4pts I/O  R/B?
        Use_8pts_IO_RB = 0  'Use 8pts I/O R/B?
        Use_16pts_IO_RB_Inputs = 0  'Use 16pts I/O Receptacle Block for Inputs? (BOOL 1 or 0)
        Use_16pts_IO_RB_Outputs = 0 'Use 16pts I/O Receptacle Block for Outputs? (BOOL 1 or 0)
        Single_Channel_IO_1_Input_per_RB = 0 'Single Channel I/O? ,1 Input Point per Receptacle Port
        Single_Channel_IO_1_Output_per_RB = 0 'Single Channel I/O? ,1 Output Point per Receptacle Port
        Splitter_Perc_dual_ch = 0 'Splitter Percentage for dual channel (1 - 100)
        Valve_solenoid_adapter = 0 'Valve Solenoid Adapters?
        Smart_Wire_Darwin_Limited = 0 'Smart Wire Darwin (SWD) on Limited Function VFDs, SoftSarts
        Smart_Wire_Darwin_Full = 0  'Smart Wire Darwin (SWD) on Full Function VFDs
        EthernetIP_Full = 0    'EthernetIP  on Full Function VFDs
        Percent_overage_inputs = 0  'Percent overage - inputs
        Percent_overage_nm_outputs = 0  'Percent overage - n.m. outputs
        Percent_overage_m_outputs = 0  'Percent overage - m. outputs

        '------------------------------ Create list of ADA_sets ---------------------------------

        ADA_set_list = New List(Of ADA_Set)()  'This list will store all our ADA sets


        ' NFPA FLA for Motors Horse Power
        '  ==== [ 230                               480                               575] =============           HP
        '------------------------------------------------------------------------------------------------
        Motor_horsepower(0, 0) = 1.4 : Motor_horsepower(0, 1) = 0.7 : Motor_horsepower(0, 2) = 0.6       '0.25
        Motor_horsepower(1, 0) = 2.2 : Motor_horsepower(1, 1) = 1.1 : Motor_horsepower(1, 2) = 0.9       '0.5
        Motor_horsepower(2, 0) = 3.2 : Motor_horsepower(2, 1) = 1.6 : Motor_horsepower(2, 2) = 1.3       '0.75
        Motor_horsepower(3, 0) = 4.2 : Motor_horsepower(3, 1) = 2.1 : Motor_horsepower(3, 2) = 1.7       '1
        Motor_horsepower(4, 0) = 6 : Motor_horsepower(4, 1) = 3 : Motor_horsepower(4, 2) = 2.4           '1.5
        Motor_horsepower(5, 0) = 6.8 : Motor_horsepower(5, 1) = 3.4 : Motor_horsepower(5, 2) = 2.7       '2
        Motor_horsepower(6, 0) = 8.2 : Motor_horsepower(6, 1) = 4.1 : Motor_horsepower(6, 2) = 3.3       '2.5
        Motor_horsepower(7, 0) = 9.6 : Motor_horsepower(7, 1) = 4.8 : Motor_horsepower(7, 2) = 3.9       '3
        Motor_horsepower(8, 0) = 12.4 : Motor_horsepower(8, 1) = 6.2 : Motor_horsepower(8, 2) = 5        '4
        Motor_horsepower(9, 0) = 15.2 : Motor_horsepower(9, 1) = 7.6 : Motor_horsepower(9, 2) = 6.1      '5
        Motor_horsepower(10, 0) = 22 : Motor_horsepower(10, 1) = 11 : Motor_horsepower(10, 2) = 9        '7.5
        Motor_horsepower(11, 0) = 28 : Motor_horsepower(11, 1) = 14 : Motor_horsepower(11, 2) = 11        '10


        ' NFPA FLA for VFD Horse Power
        '  ==== [ 230                               480                               575] =============
        '------------------------------------------------------------------------------------------------
        VFD_horsepower(0, 0) = 2.8 : VFD_horsepower(0, 1) = 1.4 : VFD_horsepower(0, 2) = 2.3          '0.25
        VFD_horsepower(1, 0) = 4.4 : VFD_horsepower(1, 1) = 2.2 : VFD_horsepower(1, 2) = 2.3          '0.5
        VFD_horsepower(2, 0) = 6.4 : VFD_horsepower(2, 1) = 3.2 : VFD_horsepower(2, 2) = 2.3          '0.75
        VFD_horsepower(3, 0) = 8.4 : VFD_horsepower(3, 1) = 4.2 : VFD_horsepower(3, 2) = 2.3          '1
        VFD_horsepower(4, 0) = 12 : VFD_horsepower(4, 1) = 6 : VFD_horsepower(4, 2) = 3.8             '1.5
        VFD_horsepower(5, 0) = 13.6 : VFD_horsepower(5, 1) = 6.8 : VFD_horsepower(5, 2) = 3.8         '2
        VFD_horsepower(6, 0) = 16.4 : VFD_horsepower(6, 1) = 8.2 : VFD_horsepower(6, 2) = 5.3         '2.5
        VFD_horsepower(7, 0) = 19.2 : VFD_horsepower(7, 1) = 9.6 : VFD_horsepower(7, 2) = 5.3          '3
        VFD_horsepower(8, 0) = 24.8 : VFD_horsepower(8, 1) = 12.4 : VFD_horsepower(8, 2) = 8.3         '4
        VFD_horsepower(9, 0) = 30.4 : VFD_horsepower(9, 1) = 15.2 : VFD_horsepower(9, 2) = 8.3         '5
        VFD_horsepower(10, 0) = 44 : VFD_horsepower(10, 1) = 22 : VFD_horsepower(10, 2) = 11.2         '7.5
        VFD_horsepower(11, 0) = 56 : VFD_horsepower(11, 1) = 28 : VFD_horsepower(11, 2) = 13.7         '10


        '-------- Array of amp  motors values ------------------  yellow row in Starter Panel tab Row 220

        Starters_values(0) = 0 : Starters_values(1) = 0.16 : Starters_values(2) = 0.25 : Starters_values(3) = 0.4 : Starters_values(4) = 0.63 : Starters_values(5) = 1
        Starters_values(6) = 1.6 : Starters_values(7) = 2.5 : Starters_values(8) = 4 : Starters_values(9) = 6.3 : Starters_values(10) = 10 : Starters_values(11) = 12
        Starters_values(12) = 16 : Starters_values(13) = 20 : Starters_values(14) = 25 : Starters_values(15) = 32

        value_to_index_480 = {420, 400, 380, 360, 340, 320, 300, 280, 260, 240, 220, 200, 180, 160, 140, 120, 110, 100, 90, 80, 70, 60, 50, 40, 30, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0}
        value_to_index_575 = {270, 260, 250, 240, 230, 220, 210, 200, 190, 180, 170, 160, 150, 140, 130, 120, 110, 100, 90, 80, 70, 60, 50, 40, 30, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0}

        '-------------------- Power Supply array --------------------------

        '-------- initialize variables ------------------
        For i = 0 To 50
            For j = 0 To 3
                Power_supply_480(i, j) = 0
                Power_supply_575(i, j) = 0
            Next
        Next


        Power_supply_480(0, 3) = 11 : Power_supply_480(1, 3) = 10 : Power_supply_480(2, 3) = 10 : Power_supply_480(3, 3) = 9 : Power_supply_480(4, 3) = 9 : Power_supply_480(5, 3) = 8
        Power_supply_480(6, 3) = 8 : Power_supply_480(7, 3) = 7 : Power_supply_480(8, 3) = 7 : Power_supply_480(9, 3) = 6 : Power_supply_480(10, 3) = 6 : Power_supply_480(11, 3) = 5
        Power_supply_480(12, 3) = 5 : Power_supply_480(13, 3) = 4 : Power_supply_480(14, 3) = 4 : Power_supply_480(15, 3) = 3 : Power_supply_480(16, 3) = 3 : Power_supply_480(17, 3) = 3
        Power_supply_480(18, 3) = 3 : Power_supply_480(19, 3) = 2 : Power_supply_480(20, 3) = 2 : Power_supply_480(21, 3) = 2 : Power_supply_480(22, 3) = 2 : Power_supply_480(23, 3) = 1
        Power_supply_480(24, 3) = 1 : Power_supply_480(25, 3) = 1 : Power_supply_480(26, 3) = 1 : Power_supply_480(27, 3) = 1 : Power_supply_480(28, 3) = 1 : Power_supply_480(29, 3) = 1
        Power_supply_480(30, 2) = 1 : Power_supply_480(31, 2) = 1 : Power_supply_480(32, 2) = 1 : Power_supply_480(33, 2) = 1 : Power_supply_480(34, 2) = 1 : Power_supply_480(35, 2) = 1
        Power_supply_480(36, 2) = 1 : Power_supply_480(37, 2) = 1 : Power_supply_480(38, 2) = 1 : Power_supply_480(39, 2) = 1 : Power_supply_480(40, 1) = 1 : Power_supply_480(41, 1) = 1
        Power_supply_480(42, 1) = 1 : Power_supply_480(43, 1) = 1 : Power_supply_480(44, 1) = 1 : Power_supply_480(45, 0) = 1 : Power_supply_480(46, 0) = 1 : Power_supply_480(47, 0) = 1
        Power_supply_480(48, 0) = 1 : Power_supply_480(49, 0) = 1

        '-----------------------------

        Power_supply_575(0, 3) = 10 : Power_supply_575(1, 3) = 9 : Power_supply_575(2, 3) = 9 : Power_supply_575(3, 3) = 8 : Power_supply_575(4, 3) = 8 : Power_supply_575(5, 3) = 7
        Power_supply_575(6, 3) = 7 : Power_supply_575(7, 3) = 6 : Power_supply_575(8, 3) = 6 : Power_supply_575(9, 3) = 5 : Power_supply_575(10, 3) = 5 : Power_supply_575(11, 3) = 4
        Power_supply_575(12, 3) = 4 : Power_supply_575(13, 3) = 4 : Power_supply_575(14, 3) = 4 : Power_supply_575(15, 3) = 3 : Power_supply_575(16, 3) = 3 : Power_supply_575(17, 3) = 3
        Power_supply_575(18, 3) = 2 : Power_supply_575(18, 2) = 1 : Power_supply_575(19, 3) = 2 : Power_supply_575(20, 3) = 2 : Power_supply_575(21, 3) = 1 : Power_supply_575(21, 2) = 1
        Power_supply_575(22, 3) = 1 : Power_supply_575(22, 2) = 1 : Power_supply_575(23, 3) = 1
        Power_supply_575(24, 3) = 1 : Power_supply_575(25, 2) = 1 : Power_supply_575(25, 1) = 1 : Power_supply_575(26, 2) = 1 : Power_supply_575(26, 1) = 1 : Power_supply_575(27, 2) = 1
        Power_supply_575(27, 1) = 1 : Power_supply_575(28, 2) = 1 : Power_supply_575(28, 1) = 1 : Power_supply_575(29, 2) = 1 : Power_supply_575(29, 1) = 1
        Power_supply_575(30, 2) = 1 : Power_supply_575(31, 2) = 1 : Power_supply_575(32, 2) = 1 : Power_supply_575(33, 2) = 1 : Power_supply_575(34, 2) = 1 : Power_supply_575(35, 2) = 1
        Power_supply_575(36, 2) = 1 : Power_supply_575(37, 2) = 1 : Power_supply_575(38, 2) = 1 : Power_supply_575(39, 2) = 1 : Power_supply_575(40, 1) = 1 : Power_supply_575(41, 1) = 1
        Power_supply_575(42, 1) = 1 : Power_supply_575(43, 1) = 1 : Power_supply_575(44, 1) = 1 : Power_supply_575(45, 0) = 1 : Power_supply_575(46, 0) = 1 : Power_supply_575(47, 0) = 1
        Power_supply_575(48, 0) = 1 : Power_supply_575(49, 0) = 1



        '------------------ Load motors table ---------------------------------
        Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.25"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.50"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.75"})
        Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  1.00"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  1.50"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  2.00"})
        Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  2.50"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  3.00"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  4.00"})
        Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  5.00"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  7.50"}) : Motors_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  10.0"})

        Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.25"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.50"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.75"})
        Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  1.00"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  1.50"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  2.00"})
        Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  2.50"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  3.00"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  4.00"})
        Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  5.00"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  7.50"}) : Motors_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  10.0"})

        Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.25"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.50"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.75"})
        Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  1.00"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  1.50"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  2.00"})
        Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  2.50"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  3.00"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  4.00"})
        Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  5.00"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  7.50"}) : Motors_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  10.0"})

        Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.25"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.50"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.75"})
        Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  1.00"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  1.50"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  2.00"})
        Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  2.50"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  3.00"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  4.00"})
        Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  5.00"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  7.50"}) : Motors_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  10.0"})

        For z = 12 To 23
            Motors_grid.Rows(z).Cells(0).Style.BackColor = Color.Peru
        Next

        For z = 36 To 47
            Motors_grid.Rows(z).Cells(0).Style.BackColor = Color.Peru
        Next

        '---------------- Load Inputs table -------------------------------------

        inputs_grid.Rows.Add(New String() {"Motor Reset PB Required?", "", "1"}) : inputs_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 1 end", "", "1"}) : inputs_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 2 ended", "", "1"})
        inputs_grid.Rows.Add(New String() {"E-Stop Cable - Total Feet"}) : inputs_grid.Rows.Add(New String() {"E-Stop Cable - Turns"}) : inputs_grid.Rows.Add(New String() {"E-Stop PB Station", "", "1"})
        inputs_grid.Rows.Add(New String() {"E-Stop PB Station Guarded", "", "1"}) : inputs_grid.Rows.Add(New String() {"Interlock Box", "", "6", "1", "1"}) : inputs_grid.Rows.Add(New String() {"Pressure Switch", "", "1"})
        inputs_grid.Rows.Add(New String() {"Prox Switch", "", "1"}) : inputs_grid.Rows.Add(New String() {"Limit Switch", "", "1"}) : inputs_grid.Rows.Add(New String() {"Photoeye 2 Hole Mnt", "", "1"})
        inputs_grid.Rows.Add(New String() {"Photoeye 18mm Mnt", "", "1"}) : inputs_grid.Rows.Add(New String() {"PhotoEye Emitter Receiver", "", "1"}) : inputs_grid.Rows.Add(New String() {"PhotoEye Diffuse 100mm range", "", "1"})
        inputs_grid.Rows.Add(New String() {"PhotoEye Diffuse 450mm range", "", "1"}) : inputs_grid.Rows.Add(New String() {"PhotoEye Diffuse 1000mm range", "", "1"}) : inputs_grid.Rows.Add(New String() {"UltraSonic Prox", "", "1"})
        inputs_grid.Rows.Add(New String() {"ADA Photoeye Brackets"})

        '------------------ Load pushbuttons table ---------------------------------
        push_grid.Rows.Add(New String() {"ADA-ASM-2CS101", "", "2"}) : push_grid.Rows.Add(New String() {"ADA-ASM-3CS101", "", "3"}) : push_grid.Rows.Add(New String() {"ADA-ASM-3CS102", "", "3"})
        push_grid.Rows.Add(New String() {"ADA-ASM-3CS103", "", "3"}) : push_grid.Rows.Add(New String() {"ADA-ASM-3CS104", "", "3"}) : push_grid.Rows.Add(New String() {"ADA-ASM-3CS105", "", "4"})
        push_grid.Rows.Add(New String() {"ADA-ASM-1CS102", "", "1"}) : push_grid.Rows.Add(New String() {"ADA-ASM-1CS103", "", "1", "1"}) : push_grid.Rows.Add(New String() {"ADA-ASM-1CS104"})
        push_grid.Rows.Add(New String() {"ADA-ASM-1CS101", "", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "1"})
        push_grid.Rows.Add(New String() {"User Defined I/O", "", "1", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "1", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "1", "1"})
        push_grid.Rows.Add(New String() {"User Defined I/O", "", "1", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "1"}) : push_grid.Rows.Add(New String() {"User Defined I/O", "", "2"})
        push_grid.Rows.Add(New String() {"User Defined I/O", "", "2"})

        '------------ load lights and motion --------------------
        Lights_grid.Rows.Add(New String() {"User Defined I/O", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"User Defined I/O", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"User Defined I/O", "", "", "1"})
        Lights_grid.Rows.Add(New String() {"User Defined I/O", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"User Defined I/O", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Stack Light Amber", "", "", "1"})
        Lights_grid.Rows.Add(New String() {"Stack Light Blue", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Stack Light Green", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Stack Light Red", "", "", "1"})
        Lights_grid.Rows.Add(New String() {"Stack Light Horn 85 db", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Stack Light Horn 100 db 8 tone", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"MZ Interlock (1 in 1 motion out)", "", "1", "", "1"})
        Lights_grid.Rows.Add(New String() {"MZ Output (1 motion out)", "", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Solenoid", "", "", "", "1"}) : Lights_grid.Rows.Add(New String() {"Double-Ended Solenoid", "", "", "", "2"})

        '--------------- load field devices ----------------------------
        Field_grid.Rows.Add(New String() {"VFD hardwired provided by others", "", "1", "", "1"}) : Field_grid.Rows.Add(New String() {"EZLogic PhotoEye Zone", "", "1"}) : Field_grid.Rows.Add(New String() {"EZLogic Controlled Zone", "", "1", "", "1"})
        Field_grid.Rows.Add(New String() {"ZP Accum Photo Eyes (EZLogic)"}) : Field_grid.Rows.Add(New String() {"Controls MDR non reversing", "", "1", "", "1"}) : Field_grid.Rows.Add(New String() {"Controls MDR reversing", "", "1", "", "2"})
        Field_grid.Rows.Add(New String() {"Release"}) : Field_grid.Rows.Add(New String() {"ZP Accum Power Supply (DC)", "", "1"}) : Field_grid.Rows.Add(New String() {"Additional amps drawn off of main 24vdc power supply"})
        Field_grid.Rows.Add(New String() {"12mm shaft Encoder"}) : Field_grid.Rows.Add(New String() {"Wheel Encoder"}) : Field_grid.Rows.Add(New String() {"Analog device"})

        '--------------- load bus extenders ----------------------------------
        bus_grid.Rows.Add(New String() {"B&R X2X Bus Extender  Module", "", "", "", ""}) : bus_grid.Rows.Add(New String() {"B&R X2X 8 Port 16 I/O IP67 Module", "", "-8", "-8", ""}) : bus_grid.Rows.Add(New String() {"SWD 1 Port 2 I/O Module", "", "-2", "", ""})
        bus_grid.Rows.Add(New String() {"SWD 2 Port 4 I/O Module", "", "-3", "-1", ""}) : bus_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as in", "", "-16", "", ""}) : bus_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as out", "", "", "-16", ""})
        bus_grid.Rows.Add(New String() {"SWD 8 Port 8 In/8 Mot Module", "", "-8", "-1", "-8"}) : bus_grid.Rows.Add(New String() {"SWD 8 Port 16 Mot Module", "", "", "", "-16"}) : bus_grid.Rows.Add(New String() {"IO Link Master Module (4 port)IP20"})
        bus_grid.Rows.Add(New String() {"IO Link Master Module (8 port)IP69", "", "-2", "-8", "-2"}) : bus_grid.Rows.Add(New String() {"IO Link Slave Module 16in (8 port)IP69", "", "-16", "", ""}) : bus_grid.Rows.Add(New String() {"IO Link Slave Module 10in6out (8 port)IP69", "", "-10", "-3", "-3"})


        '-------------------load scanners -------------------------------------
        scanner_grid.Rows.Add(New String() {"Scanner - Line (Sick)"}) : scanner_grid.Rows.Add(New String() {"Scanner - Omni (Mini-X)"})
        scanner_grid.Rows.Add(New String() {"Cognex Bar Code Scan DataMan KIT"}) : scanner_grid.Rows.Add(New String() {"Cognex Insite Vision KIT"})

        '------------------- load brake ---------------------------------------
        brake_grid.Rows.Add(New String() {"Clutch/Brake 120VAC Powered", "", "", "", "2"}) : brake_grid.Rows.Add(New String() {"Clutch/Brake 24 VDC Powered", "", "", "", "2"})
        brake_grid.Rows.Add(New String() {"120VAC 0.5kva Power supply"}) : brake_grid.Rows.Add(New String() {"120VAC 2.0kva Power supply"})
        brake_grid.Rows.Add(New String() {"EZLogic Amps needed at 24vdc"}) : brake_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"})
        brake_grid.Rows.Add(New String() {"Total Aux. Amps needed at 24vdc"}) : brake_grid.Rows.Add(New String() {"Non SWD Class2 PWR (Dumb Box)"})
        brake_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"}) : brake_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Drops", "", "", "", "2"})
        brake_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Ext"}) : brake_grid.Rows.Add(New String() {"Add. Motion Amps needed at 24vdc"})
        brake_grid.Rows.Add(New String() {"Amps needed at 24vdc (260 Max)"}) : brake_grid.Rows.Add(New String() {"Non SWD MDR PWR (Dumb Box)", "", "1"})
        brake_grid.Rows.Add(New String() {"Standard Estop / Breaker Inputs", "1", "8"})

        '------------------- LOAD TOTALS --------------------------
        totals_grid.Rows.Add(New String() {"Starter Panel – Materials"}) : totals_grid.Rows.Add(New String() {"Starter Panel - Labor"})
        totals_grid.Rows.Add(New String() {"IO Panel – Materials"}) : totals_grid.Rows.Add(New String() {"IO Panel – Labor"})
        totals_grid.Rows.Add(New String() {"PLC Panel – Materials"}) : totals_grid.Rows.Add(New String() {"PLC Panel – Labor"}) : totals_grid.Rows.Add(New String() {""})
        totals_grid.Rows.Add(New String() {"Field Parts – Materials"}) : totals_grid.Rows.Add(New String() {"Field Parts – Labor"})
        totals_grid.Rows.Add(New String() {"Scanners – Materials"}) : totals_grid.Rows.Add(New String() {"Scanners – Labor"})
        totals_grid.Rows.Add(New String() {"M12 Cables – Materials"}) : totals_grid.Rows.Add(New String() {"M12 Cables – Labor"})
        totals_grid.Rows.Add(New String() {"M12 Estop Cables – Materials"}) : totals_grid.Rows.Add(New String() {"M12 Estop Cables – Labor"}) : totals_grid.Rows.Add(New String() {""})
        totals_grid.Rows.Add(New String() {"SHIPPING & HANDLING I/O"}) : totals_grid.Rows.Add(New String() {"SHIPPING & HANDLING Motor Big"})
        totals_grid.Rows.Add(New String() {"SHIPPING & HANDLING – Scanners"}) : totals_grid.Rows.Add(New String() {""}) : totals_grid.Rows.Add(New String() {"Electrical Installation Labor"})
        totals_grid.Rows.Add(New String() {"Electrical Installation Materials"}) : totals_grid.Rows.Add(New String() {"Electrical Installation Expenses"}) : totals_grid.Rows.Add(New String() {"Electrical Installation Subcontract"})

        For z = 20 To 23
            totals_grid.Rows(z).Cells(0).Style.BackColor = Color.Peru
        Next


        '--------------------INSTALLATION  SPREADSHEET ---------------------------
        For i = 4 To 12
            Install_grid.Columns(i).HeaderCell.Style.BackColor = Color.DarkSalmon
        Next


        Install_grid.Rows.Add(New String() {"INSTALLATION"})
        Install_grid.Rows(0).DefaultCellStyle.BackColor = Color.Gray
        Install_grid.Rows(0).DefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
        Install_grid.Rows(0).ReadOnly = True

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



        M12_cables.Add("ADA-SAC5-02-G", 0)
        M12_cables.Add("ADA-SAC5-03-G", 0)
        M12_cables.Add("ADA-SAC5-05-G", 0)
        M12_cables.Add("ADA-SAC5-10-G", 0)
        M12_cables.Add("ADA-SAC5-15-G", 0)
        M12_cables.Add("ADA-SAC5-20-G", 0)


        M12_ES_cables.Add("ADA-SAC5-02-Y", 0)
        M12_ES_cables.Add("ADA-SAC5-03-Y", 0)
        M12_ES_cables.Add("ADA-SAC5-05-Y", 0)
        M12_ES_cables.Add("ADA-SAC5-10-Y", 0)
        M12_ES_cables.Add("ADA-SAC5-15-Y", 0)
        M12_ES_cables.Add("ADA-SAC5-20-Y", 0)


    End Sub

    Private Sub VoltsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles volts_575.Click
        volts_575.Checked = True
        volts_480.Checked = False
        volts_230.Checked = False

        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub VoltsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles volts_480.Click
        volts_480.Checked = True
        volts_575.Checked = False
        volts_230.Checked = False

        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub volts_230_Click(sender As Object, e As EventArgs) Handles volts_230.Click
        volts_230.Checked = True
        volts_480.Checked = False
        volts_575.Checked = False

        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
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

    Private Sub Button7_MouseEnter(sender As Object, e As EventArgs) Handles Button7.MouseEnter
        Button7.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Button7_MouseLeave(sender As Object, e As EventArgs) Handles Button7.MouseLeave
        Button7.BackColor = Color.Teal
    End Sub

    Private Sub Button3_MouseEnter(sender As Object, e As EventArgs) Handles Button3.MouseEnter
        Button3.BackColor = Color.OrangeRed
    End Sub

    Private Sub Button3_MouseLeave(sender As Object, e As EventArgs) Handles Button3.MouseLeave
        Button3.BackColor = Color.DarkRed
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        '-------------------//////  Create an ADA Set with the data entered in the textboxes ////////-----------------

        If Not ADA_name.Text.Equals("") Then

            If Valid_name(ADA_set_list, ADA_name.Text) = True Then

                If get_inputs() = False Then

                    ADA_set_list.Add(New ADA_Set(ADA_name.Text, Motor_array, inputs_array, push_array, lights_array, field_array, bus_array, scanner_array, brakes_array))  'create set with constructor             
                    Call get_global_set()  'update global variables

                    '--------- enter global settings values -------------
                    ADA_set_list.Item(ADA_set_list.count - 1).Update_globals(Voltage, efficiency, Power_factor, include_ADA_brackets, Use_4pts_IO, Use_8pts_IO, Use_16pts_IO, removal_msbox_t, M23_at_IO_Receptacle,
            M23_Bulkhead_at_ADA, Use_4pts_IO_RB, Use_8pts_IO_RB, Use_16pts_IO_RB_Inputs, Use_16pts_IO_RB_Outputs, Single_Channel_IO_1_Input_per_RB, Single_Channel_IO_1_Output_per_RB, Splitter_Perc_dual_ch, Valve_solenoid_adapter, Smart_Wire_Darwin_Limited,
            Smart_Wire_Darwin_Full, EthernetIP_Full, Percent_overage_inputs, Percent_overage_nm_outputs, Percent_overage_m_outputs)
                    '---------------------------------------------------


                    '-------- load textbox inputs ---------------
                    ' ADA_set_list.Item(ADA_set_list.count - 1).Load_textboxes_inputs()

                    '--------------------( Calculate ADA Parts quantitites )--------------------
                    '-------Starter Panel  (Note: we call this methods twice to make sure all qty are calculated ----------


                    ADA_set_list.Item(ADA_set_list.count - 1).Calculate_Starter_Panel(Motor_horsepower, VFD_horsepower, Starters_values, value_to_index_480, value_to_index_575, Power_supply_480, Power_supply_575)
                    ADA_set_list.Item(ADA_set_list.count - 1).Calculate_IO_Panel
                    ADA_set_list.Item(ADA_set_list.count - 1).Calculate_PLC_Panel
                    ADA_set_list.Item(ADA_set_list.count - 1).Calculate_Field_Panel
                    ADA_set_list.Item(ADA_set_list.count - 1).Calculate_Scanners

                    temp_name = ADA_name.Text  'update temp_ada_name for updating purpouses. Ada name can get lost and we will not be able to access the ada set
                    Call Load_set_names(ADA_set_list)  'reset dropbox with all ada set names
                    Call M12_process()
                    Call Cal_TOTAL_values() 'refresh TOTALS
                    Call gen_PR() 'generate PR
                    Call Cal_PartCost("Starter Panel")

                    MessageBox.Show(ADA_set_list.Item(ADA_set_list.count - 1).myName & " was created successfully!")
                Else
                    MessageBox.Show("Invalid value")   'invalid value message
                End If

            Else
                MessageBox.Show("Duplicate name! Please enter a different one")

            End If
        Else
                MessageBox.Show("Please enter the ADA Set name")  'empty ada set name textbox
        End If

    End Sub

    Sub get_global_set()

        '-----------------------  Add global settings to the set ----------------------------
        If volts_480.Checked = True Then
            Voltage = 480
        ElseIf volts_575.Checked = True Then
            Voltage = 575
        Else
            Voltage = 230
        End If


        efficiency = If((String.IsNullOrEmpty(Efficiency_box.Text) = False And IsNumeric(Efficiency_box.Text) = True), Efficiency_box.Text, 0)
        Power_factor = If((String.IsNullOrEmpty(Power_factor_box.Text) = False And IsNumeric(Power_factor_box.Text) = True), Power_factor_box.Text, 0)
        include_ADA_brackets = If((String.IsNullOrEmpty(ADA_Brackets_box.Text) = False And IsNumeric(ADA_Brackets_box.Text) = True), ADA_Brackets_box.Text, 0)

        Use_4pts_IO = If((String.IsNullOrEmpty(Use_4ptsIO_box.Text) = False And IsNumeric(Use_4ptsIO_box.Text) = True), Use_4ptsIO_box.Text, 0)
        Use_8pts_IO = If((String.IsNullOrEmpty(Use_8ptsIO_box.Text) = False And IsNumeric(Use_8ptsIO_box.Text) = True), Use_8ptsIO_box.Text, 0)
        Use_16pts_IO = If((String.IsNullOrEmpty(Use_16ptsIO_box.Text) = False And IsNumeric(Use_16ptsIO_box.Text) = True), Use_16ptsIO_box.Text, 0)

        removal_msbox_t = If((String.IsNullOrEmpty(removal_msbox.Text) = False And IsNumeric(removal_msbox.Text) = True), removal_msbox.Text, 0)
        M23_at_IO_Receptacle = If((String.IsNullOrEmpty(M23_IO_box.Text) = False And IsNumeric(M23_IO_box.Text) = True), M23_IO_box.Text, 0)
        M23_Bulkhead_at_ADA = If((String.IsNullOrEmpty(M23_bulkhead_box.Text) = False And IsNumeric(M23_bulkhead_box.Text) = True), M23_bulkhead_box.Text, 0)

        Use_4pts_IO_RB = If((String.IsNullOrEmpty(pts_4_rb_box.Text) = False And IsNumeric(pts_4_rb_box.Text) = True), pts_4_rb_box.Text, 0)
        Use_8pts_IO_RB = If((String.IsNullOrEmpty(pts_8_rb_box.Text) = False And IsNumeric(pts_8_rb_box.Text) = True), pts_8_rb_box.Text, 0)
        Use_16pts_IO_RB_Inputs = If((String.IsNullOrEmpty(pts16_recep_inputs_box.Text) = False And IsNumeric(pts16_recep_inputs_box.Text) = True), pts16_recep_inputs_box.Text, 0)
        Use_16pts_IO_RB_Outputs = If((String.IsNullOrEmpty(pts16_recep_outputs_box.Text) = False And IsNumeric(pts16_recep_outputs_box.Text) = True), pts16_recep_outputs_box.Text, 0)

        Single_Channel_IO_1_Input_per_RB = If((String.IsNullOrEmpty(Single_input_box.Text) = False And IsNumeric(Single_input_box.Text) = True), Single_input_box.Text, 0)
        Single_Channel_IO_1_Output_per_RB = If((String.IsNullOrEmpty(single_output_box.Text) = False And IsNumeric(single_output_box.Text) = True), single_output_box.Text, 0)
        Splitter_Perc_dual_ch = If((String.IsNullOrEmpty(splitter_per_box.Text) = False And IsNumeric(splitter_per_box.Text) = True), splitter_per_box.Text, 0)
        Valve_solenoid_adapter = If((String.IsNullOrEmpty(valve_box.Text) = False And IsNumeric(valve_box.Text) = True), valve_box.Text, 0)

        Smart_Wire_Darwin_Limited = If((String.IsNullOrEmpty(swd_limited_box.Text) = False And IsNumeric(swd_limited_box.Text) = True), swd_limited_box.Text, 0)
        Smart_Wire_Darwin_Full = If((String.IsNullOrEmpty(swd_full_box.Text) = False And IsNumeric(swd_full_box.Text) = True), swd_full_box.Text, 0)
        EthernetIP_Full = If((String.IsNullOrEmpty(ethernetip_box.Text) = False And IsNumeric(ethernetip_box.Text) = True), ethernetip_box.Text, 0)
        Percent_overage_inputs = If((String.IsNullOrEmpty(percentage_inputs_box.Text) = False And IsNumeric(percentage_inputs_box.Text) = True), percentage_inputs_box.Text, 0)
        Percent_overage_nm_outputs = If((String.IsNullOrEmpty(percentage_nmoutputs_box.Text) = False And IsNumeric(percentage_nmoutputs_box.Text) = True), percentage_nmoutputs_box.Text, 0)
        Percent_overage_m_outputs = If((String.IsNullOrEmpty(percentage_moutputs_box.Text) = False And IsNumeric(percentage_moutputs_box.Text) = True), percentage_moutputs_box.Text, 0)


    End Sub


    Function get_inputs() As Boolean

        '------------- create arrays and extract all inputs return true if an invalid value is detected

        get_inputs = False
        '  Dim Motor_array(48, 4) As String  'NFPA FLA for Motor Horse Power


        Dim i As Integer = 0

        '------------- Motors datagrid --------------------
        For Each row As DataGridViewRow In Motors_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    Motor_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        Motor_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString

                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- Inputs datagrid --------------------
        For Each row As DataGridViewRow In inputs_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    inputs_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        inputs_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString

                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- push datagrid --------------------
        For Each row As DataGridViewRow In push_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    push_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        push_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- Lights datagrid --------------------
        For Each row As DataGridViewRow In Lights_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    lights_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        lights_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- Field datagrid --------------------
        For Each row As DataGridViewRow In Field_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    field_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        field_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- bus datagrid --------------------
        For Each row As DataGridViewRow In bus_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    bus_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        bus_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- scanner datagrid --------------------
        For Each row As DataGridViewRow In scanner_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    scanner_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        scanner_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        i = 0
        '------------- Brake datagrid --------------------
        For Each row As DataGridViewRow In brake_grid.Rows
            If row.IsNewRow Then Continue For

            For j = 1 To 4
                If String.IsNullOrEmpty(row.Cells(j).Value) = True Then
                    brakes_array(i, j - 1) = 0
                Else
                    If IsNumeric(row.Cells(j).Value) = False Then
                        get_inputs = True
                    Else
                        brakes_array(i, j - 1) = Math.Round(CType(row.Cells(j).Value, Double)).ToString
                    End If
                End If
            Next
            i += 1
        Next

        Dim my_inp As Double : my_inp = myinputs()
        Dim temp_Rem As Double : temp_Rem = 0
        If Not remote_box.SelectedItem Is Nothing Then
            'remote io
            If String.Equals(remote_box.SelectedItem.ToString, "Yes") Then
                temp_Rem = 1
            End If
        End If

        Dim total_i As Double : total_i = my_inp * 16 * temp_Rem

        If (temp_Rem = 0) = False Then

            If get_inputs = False Then  'check for SWD 8 Port 16 I/O Module right qty

                If bus_array(4, 0) <= total_i And bus_array(4, 0) > 0 Then
                    Dim result As DialogResult = MessageBox.Show("SWD 8 Port 16 I/O Module in qty needs to be greater than " & total_i & ". Are you sure you want to continue", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.No) Then
                        get_inputs = True
                    End If
                End If
            End If

            If get_inputs = False Then  'check for SWD 8 Port 16 I/O Module as out

                If bus_array(5, 0) <= total_i And bus_array(5, 0) > 0 Then
                    Dim result As DialogResult = MessageBox.Show("SWD 8 Port 16 I/O Module out qty needs to be greater than " & total_i & ". Are you sure you want to continue", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.No) Then
                        get_inputs = True
                    End If
                End If
            End If

            If get_inputs = False Then  'check for SWD 8 Port 16 Mot Module

                If bus_array(7, 0) <= total_i And bus_array(7, 0) > 0 Then
                    Dim result As DialogResult = MessageBox.Show("SWD 8 Port 16 I/O Module qty needs to be greater than " & total_i & ". Are you sure you want to continue", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.No) Then
                        get_inputs = True
                    End If
                End If
            End If

        End If

    End Function


    Sub Load_set_names(myList As List(Of ADA_Set))

        'this sub will add all the set names to the dropbox

        Dim i As Integer : i = 0
        ComboBox1.Items.Clear()

        For Each ada_set In myList

            ComboBox1.Items.Add(ADA_set_list.Item(i).myName)

            i = i + 1
        Next

    End Sub


    Private Sub ClearAllADAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllADAToolStripMenuItem.Click
        'Remove all the ADA Sets

        Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove all the sets", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then
            ADA_set_list.clear()
            ComboBox1.Items.Clear()
            Call Cal_TOTAL_values() 'refresh TOTALS

            MessageBox.Show("All set were removed")
        End If
    End Sub

    '------------------ UPDATE SET ------------------

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim proceed_update As Boolean : proceed_update = False

        If String.Equals(ADA_name.Text, temp_name) = True Then
            proceed_update = True
        Else
            proceed_update = Valid_name(ADA_set_list, ADA_name.Text)
        End If


        If proceed_update = True Then

            If Not ADA_name.Text.Equals("") Then

                Dim myIndex As Integer : myIndex = Get_set_number(ADA_set_list, temp_name)
                If myIndex > -1 Then

                    If get_inputs() = False Then

                        ADA_set_list.Item(myIndex).Update_inputs(Motor_array, inputs_array, push_array, lights_array, field_array, bus_array, scanner_array, brakes_array)
                        ADA_set_list.Item(myIndex).myName = ADA_name.Text


                        '--------- enter global settings values -------------
                        ADA_set_list.Item(myIndex).Update_globals(Voltage, efficiency, Power_factor, include_ADA_brackets, Use_4pts_IO, Use_8pts_IO, Use_16pts_IO, removal_msbox_t, M23_at_IO_Receptacle,
            M23_Bulkhead_at_ADA, Use_4pts_IO_RB, Use_8pts_IO_RB, Use_16pts_IO_RB_Inputs, Use_16pts_IO_RB_Outputs, Single_Channel_IO_1_Input_per_RB, Single_Channel_IO_1_Output_per_RB, Splitter_Perc_dual_ch, Valve_solenoid_adapter, Smart_Wire_Darwin_Limited,
            Smart_Wire_Darwin_Full, EthernetIP_Full, Percent_overage_inputs, Percent_overage_nm_outputs, Percent_overage_m_outputs)
                        '---------------------------------------------------

                        '-------- load textbox inputs ---------------
                        ADA_set_list.Item(myIndex).Load_textboxes_inputs()

                        '-------- load dropboxes ---------------
                        ADA_set_list.Item(myIndex).load_device_Type_specs()

                        '----------- Calculate ADA parts -------------

                        ADA_set_list.Item(myIndex).Calculate_Starter_Panel(Motor_horsepower, VFD_horsepower, Starters_values, value_to_index_480, value_to_index_575, Power_supply_480, Power_supply_575)
                        ADA_set_list.Item(myIndex).Calculate_IO_Panel
                        ADA_set_list.Item(myIndex).Calculate_PLC_Panel
                        ADA_set_list.Item(myIndex).Calculate_Field_Panel
                        ADA_set_list.Item(myIndex).Calculate_Scanners
                        '  ADA_set_list.Item(myIndex).Calculate_M12
                        '  ADA_set_list.Item(myIndex).Calculate_M12_ES


                        Call Load_set_names(ADA_set_list)
                        Call M12_process() ' Refresh M12
                        Call Cal_TOTAL_values() 'refresh TOTALS
                        Call gen_PR()
                        Call Cal_PartCost("Starter Panel")

                        temp_name = ADA_name.Text
                        MessageBox.Show("Set Updated successfully!")
                    Else
                        MessageBox.Show("Invalid value")
                    End If

                End If
            Else
                MessageBox.Show("No ADA set name entered")
            End If
        Else
            MessageBox.Show("Duplicate ADA set name")
        End If

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        '-------------------------------- LOAD THE ENTIRE SET TO THE TEXTBOXES AND GRIDS WHEN SELECTED------------------------------------------------

        'select an ADA Set from dropboxto display info
        Dim set_name As String : set_name = ""

        If Not ComboBox1.SelectedItem Is Nothing Then
            set_name = ComboBox1.SelectedItem.ToString
        Else
            set_name = ""
        End If

        'look for the specified set
        Dim myIndex As Integer : myIndex = Get_set_number(ADA_set_list, set_name)

        If myIndex > -1 Then

            '----------------------------------------------- FILL DATAGRIDS ---------------------------------------------------

            ' enter all values to motor grid 
            Dim i As Integer : i = 0
            For Each row As DataGridViewRow In Motors_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).motors_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).motors_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to inputs grid 
            i = 0
            For Each row As DataGridViewRow In inputs_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).inputs_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else

                        row.Cells(j).Value = ADA_set_list.Item(myIndex).inputs_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to push grid 
            i = 0
            For Each row As DataGridViewRow In push_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).push_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).push_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to lights grid 
            i = 0
            For Each row As DataGridViewRow In Lights_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).lights_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).lights_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to field grid 
            i = 0
            For Each row As DataGridViewRow In Field_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).field_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).field_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to bus grid 
            i = 0
            For Each row As DataGridViewRow In bus_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).bus_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).bus_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to scanner grid 
            i = 0
            For Each row As DataGridViewRow In scanner_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).scanner_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).scanner_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            ' enter all values to brake grid 
            i = 0
            For Each row As DataGridViewRow In brake_grid.Rows
                If row.IsNewRow Then Continue For
                For j = 1 To 4
                    If ADA_set_list.Item(myIndex).brakes_table(i, j - 1) = "0" Then
                        row.Cells(j).Value = ""
                    Else
                        row.Cells(j).Value = ADA_set_list.Item(myIndex).brakes_table(i, j - 1)
                    End If
                Next
                i += 1
            Next

            temp_name = ADA_set_list.Item(myIndex).myName
            ADA_name.Text = ADA_set_list.Item(myIndex).myName   'display name set

            '-------- fill ADA set textboxes -----------------

            ada_nema12box24_text.Text = ADA_set_list.Item(myIndex).ADA_NEMA12_BOX24_No_Starters : ada_nema12box30_text.Text = ADA_set_list.Item(myIndex).ADA_NEMA12_BOX30_No_Starters
            disconnect_text.Text = ADA_set_list.Item(myIndex).motor_disconnect
            MDRs.Text = ADA_set_list.item(myIndex).MDRs

            '--- fill dropboxes ----------

            'If ADA_set_list.Item(myIndex).panel_bool = 1 Then
            '    Panel_bool_drop.Text = "Yes"
            'Else
            '    Panel_bool_drop.Text = "No"
            'End If

            If ADA_set_list.Item(myIndex).NEMA12 = 1 Then
                Nema12_com.Text = "NEMA12"
            Else
                Nema12_com.Text = "NEMA4X"
            End If


            If ADA_set_list.Item(myIndex).PLC_GateWay_PGW = 1 Then
                PLC_gate_box.Text = "Yes"
            Else
                PLC_gate_box.Text = "No"
            End If

            If ADA_set_list.Item(myIndex).ControlLogix_PLC_Box = 1 Then
                control_box.Text = "Yes"
            Else
                control_box.Text = "No"
            End If

            If ADA_set_list.Item(myIndex).CompactLogix_PLC_Box = 1 Then
                compact_box.Text = "Yes"
            Else
                compact_box.Text = "No"
            End If


            If ADA_set_list.Item(myIndex).PLC_UPS_Battery_Backup = 1 Then
                PLC_UPS_box.Text = "Yes"
            Else
                PLC_UPS_box.Text = "No"
            End If

            If ADA_set_list.Item(myIndex).Amp_Mon_Electro = 1 Then
                amp_mon_box.Text = "Yes"
            Else
                amp_mon_box.Text = "No"
            End If

            If ADA_set_list.Item(myIndex).Remote_IO = 1 Then
                remote_box.Text = "Yes"
            Else
                remote_box.Text = "No"
            End If

        End If

    End Sub

    Function Get_set_number(myList As List(Of ADA_Set), name As String) As Integer

        'get index of a set according to name

        Dim i As Integer : i = 0
        Get_set_number = -1

        For Each ada_set In myList

            If String.Compare(ada_set.myName, name) = 0 Then
                Get_set_number = i
                Exit For
            End If

            i = i + 1
        Next

    End Function


    'makes sure ADA set names are unique
    Function Valid_name(myList As List(Of ADA_Set), name As String) As Boolean

        Valid_name = True

        For Each ada_set In myList

            If String.Compare(ada_set.myName, name) = 0 Then
                Valid_name = False
                Exit For
            End If
        Next

    End Function


    'Clear motor  grids
    Private Sub ClearMotorsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearMotorsToolStripMenuItem.Click
        Call Clear_grids(Motors_grid)
    End Sub

    'Clear push grid
    Private Sub ClearPushbuttonsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearPushbuttonsToolStripMenuItem.Click
        Call Clear_grids(push_grid)
    End Sub

    'Clear inputs
    Private Sub ClearInputsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearInputsToolStripMenuItem.Click
        Call Clear_grids(inputs_grid)
    End Sub

    'Clear lights
    Private Sub ClearLightsAndOutputsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearLightsAndOutputsToolStripMenuItem.Click
        Call Clear_grids(Lights_grid)
    End Sub

    'Clear field
    Private Sub ClearFieldDevicesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearFieldDevicesToolStripMenuItem.Click
        Call Clear_grids(Field_grid)
    End Sub

    'Clear bus
    Private Sub ClearBusExtendersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearBusExtendersToolStripMenuItem.Click
        Call Clear_grids(bus_grid)
    End Sub

    'Clear scanner
    Private Sub ClearScannerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearScannerToolStripMenuItem.Click
        Call Clear_grids(scanner_grid)
    End Sub

    'Clear brakes
    Private Sub ClearBrakesAndPowerSuppliesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearBrakesAndPowerSuppliesToolStripMenuItem.Click
        Call Clear_grids(brake_grid)
    End Sub

    Sub Clear_grids(mydatagrid As DataGridView)

        'This sub clear the passed datagridview
        Dim i As Integer : i = 0

        For Each row As DataGridViewRow In mydatagrid.Rows

            If row.IsNewRow Then Continue For
            row.Cells(1).Value = ""
            row.Cells(2).Value = ""
            row.Cells(3).Value = ""
            row.Cells(4).Value = ""
            i += 1

        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '------------------------  Remove ADA Set from the list and dropbox ------------------------------

        Dim myIndex As Integer : myIndex = Get_set_number(ADA_set_list, ADA_name.Text)

        Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove " & ADA_name.Text & " set", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes And myIndex >= 0) Then
            '  ComboBox1.Items.Remove(ADA_set_list.Item(myIndex).myName)
            ADA_set_list.Removeat(myIndex)
            Call Load_set_names(ADA_set_list)
            Call M12_process() 'refresh M12
            '  Call Cal_PartCost("Starter Panel")
            Call Cal_TOTAL_values() 'refresh TOTALS
            Call gen_PR()
            Call Cal_PartCost("Starter Panel")
        End If

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs)

        '------------- GENERATE PURCHASE REQUEST FORM --------------

        'PR_grid.Rows.Clear()  'Clear the PR Grid
        'Dim count As Integer : count = 0
        'Dim total_q As Double : total_q = 0

        'Try
        '    Dim cmd As New MySqlCommand
        '    cmd.CommandText = "SELECT ADA_number, Part_Name, Part_Description, manufacturer, Type, Primary_Vendor, KIT_DEVICE from ada_active_parts"

        '    cmd.Connection = Login.Connection
        '    Dim reader As MySqlDataReader
        '    reader = cmd.ExecuteReader

        '    If reader.HasRows Then

        '        While reader.Read

        '            If (get_total_Qty(reader(0), ADA_set_list) > 0) Then

        '                PR_grid.Rows.Add(New DataGridViewRow)
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(0).Value = reader(0)  'ADA Number                        
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(1).Value = reader(1)  'part name
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(2).Value = reader(2)  'Part description
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(3).Value = reader(3)  'manufacturer
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(4).Value = reader(4)  'Type
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(5).Value = reader(5)  'vendor             
        '                PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(7).Value = get_total_Qty(reader(0), ADA_set_list) : total_q = total_q + get_total_Qty(reader(0), ADA_set_list)
        '                count = count + 1

        '                If (String.Equals(reader(6), "KIT") = True) Then

        '                    PR_grid.Rows(PR_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.DarkSeaGreen
        '                End If

        '            End If
        '        End While
        '    End If

        '    reader.Close()

        '    For Each row As DataGridViewRow In PR_grid.Rows
        '        If row.IsNewRow Then Continue For
        '        row.Cells(6).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(5).Value)
        '        row.Cells(8).Value = row.Cells(6).Value * row.Cells(7).Value
        '        row.Height = 52
        '    Next


        '    MessageBox.Show("Purchase Request generated succesfully!")

        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try



    End Sub

    Function get_total_Qty(ada_name As String, myList As List(Of ADA_Set)) As Double

        'This function returns the total qty of the part (Used in the PR) (Starter, IO, PLC, Field, Scanners and M12

        get_total_Qty = 0

        For Each ada_set In myList
            If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                get_total_Qty = get_total_Qty + ada_set.Starter_Panel.Item(ada_name) + ada_set.IO_Panel.Item(ada_name) + ada_set.PLC_Panel.Item(ada_name) + ada_set.Field.Item(ada_name) + ada_set.Scanners.Item(ada_name) '+ ada_set.M12.Item(ada_name) + ada_set.M12_ES.Item(ada_name)
            End If

            If M12_cables.ContainsKey(ada_name) Then
                get_total_Qty = get_total_Qty + M12_cables.Item(ada_name)
            End If


            If M12_ES_cables.ContainsKey(ada_name) Then
                get_total_Qty = get_total_Qty + M12_ES_cables.Item(ada_name)
            End If
        Next




        get_total_Qty = Math.Ceiling(get_total_Qty)

    End Function

    Function get_total_bytype(ada_name As String, myList As List(Of ADA_Set), type As String) As Double

        get_total_bytype = 0


        For Each ada_set In myList
            Select Case type
                Case "Starter Panel"
                    If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + ada_set.Starter_Panel.Item(ada_name)
                    End If
                Case "IO Panel"
                    If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + ada_set.IO_Panel.Item(ada_name)
                    End If
                Case "PLC Panel"
                    If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + ada_set.PLC_Panel.Item(ada_name)
                    End If
                Case "Scanners"
                    If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + ada_set.Scanners.Item(ada_name)
                    End If
                Case "Field parts"
                    If ada_set.Starter_Panel.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + ada_set.Field.Item(ada_name)
                    End If
                Case "M12 Cables"
                    If M12_cables.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + M12_cables.Item(ada_name)
                    End If
                Case "M12 ES Cables"
                    If M12_ES_cables.ContainsKey(ada_name) Then
                        get_total_bytype = get_total_bytype + M12_ES_cables.Item(ada_name)
                    End If
            End Select
        Next

        get_total_bytype = Math.Ceiling(get_total_bytype)

    End Function

    Function get_total_byset(ada_name As String, type As String, myindex As Integer) As Double

        get_total_byset = 0

        Select Case type
            Case "Starter Panel"
                If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                    get_total_byset = ADA_set_list.Item(myindex).Starter_Panel.Item(ada_name)
                End If
            Case "IO Panel"
                If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                    get_total_byset = ADA_set_list.Item(myindex).IO_Panel.Item(ada_name)
                End If
            Case "PLC Panel"
                If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                    get_total_byset = ADA_set_list.Item(myindex).PLC_Panel.Item(ada_name)
                End If
            Case "Scanners"
                If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                    get_total_byset = ADA_set_list.Item(myindex).Scanners.Item(ada_name)
                End If
            Case "Field parts"
                If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                    get_total_byset = ADA_set_list.Item(myindex).Field.Item(ada_name)
                End If
                'Case "M12 Cables"
                ' If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                ' get_total_byset = ADA_set_list.Item(myindex).M12.Item(ada_name)
                ' End If
                ' Case "M12 ES Cables"
                ' If ADA_set_list.Item(myindex).Starter_Panel.ContainsKey(ada_name) Then
                ' get_total_byset = ADA_set_list.Item(myindex).M12_ES.Item(ada_name)
                '  End If
        End Select

        get_total_byset = Math.Ceiling(get_total_byset)

    End Function

    Sub Total_rows()

        'Calculate total row number

        Dim n_rows As Double : n_rows = 0
        Dim total_parts As Double : total_parts = 0

        For i = 0 To PR_grid.Rows.Count - 2
            total_parts = total_parts + PR_grid.Rows(i).Cells(7).Value()

            If (String.IsNullOrEmpty(PR_grid.Rows(i).Cells(1).Value())) = False Then
                n_rows = n_rows + 1
            End If
        Next

        Label14.Text = "# Parts: " & total_parts


        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(6).Value) = True And IsNumeric(row.Cells(7).Value)) Then
                row.Cells(8).Value = row.Cells(6).Value * row.Cells(7).Value
            End If
        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

        ''generate quote
        'Dim appPath As String = Application.StartupPath()

        'Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        'Dim xlWorkSheet As Excel.Worksheet
        'xlApp.DisplayAlerts = False

        'If xlApp Is Nothing Then
        '    MessageBox.Show("Excel is not properly installed!!")
        'Else
        '    Try

        '        Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\Template.xlsx")
        '        xlWorkSheet = wb.Sheets("TOTALS")

        '        'Start filling the vaules from TOTAL datagridview

        '        xlWorkSheet.Cells(4, 6) = totals_grid.Rows(0).Cells(1).Value()
        '        xlWorkSheet.Cells(5, 6) = totals_grid.Rows(1).Cells(1).Value()
        '        xlWorkSheet.Cells(6, 6) = totals_grid.Rows(2).Cells(1).Value()
        '        xlWorkSheet.Cells(7, 6) = totals_grid.Rows(3).Cells(1).Value()
        '        xlWorkSheet.Cells(8, 6) = totals_grid.Rows(4).Cells(1).Value()
        '        xlWorkSheet.Cells(9, 6) = totals_grid.Rows(5).Cells(1).Value()
        '        xlWorkSheet.Cells(22, 6) = totals_grid.Rows(7).Cells(1).Value()
        '        xlWorkSheet.Cells(23, 6) = totals_grid.Rows(8).Cells(1).Value()
        '        xlWorkSheet.Cells(24, 6) = totals_grid.Rows(9).Cells(1).Value()
        '        xlWorkSheet.Cells(25, 6) = totals_grid.Rows(10).Cells(1).Value()
        '        xlWorkSheet.Cells(26, 6) = totals_grid.Rows(11).Cells(1).Value()
        '        xlWorkSheet.Cells(27, 6) = totals_grid.Rows(12).Cells(1).Value()
        '        xlWorkSheet.Cells(28, 6) = totals_grid.Rows(13).Cells(1).Value()
        '        xlWorkSheet.Cells(29, 6) = totals_grid.Rows(14).Cells(1).Value()

        '        'installation
        '        xlWorkSheet.Cells(82, 6) = totals_grid.Rows(20).Cells(1).Value()
        '        xlWorkSheet.Cells(83, 6) = totals_grid.Rows(21).Cells(1).Value()
        '        xlWorkSheet.Cells(84, 6) = totals_grid.Rows(22).Cells(1).Value()
        '        xlWorkSheet.Cells(85, 6) = totals_grid.Rows(23).Cells(1).Value()

        '        'est shipping
        '        xlWorkSheet.Cells(89, 3) = totals_grid.Rows(16).Cells(1).Value()
        '        xlWorkSheet.Cells(90, 3) = totals_grid.Rows(17).Cells(1).Value()
        '        xlWorkSheet.Cells(91, 3) = totals_grid.Rows(18).Cells(1).Value()


        '        SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

        '        If SaveFileDialog1.ShowDialog = DialogResult.OK Then

        '            wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)

        '        End If

        '        wb.Close(False)


        '        Marshal.ReleaseComObject(xlApp)

        '        MessageBox.Show("Quote generated successfully!")


        '    Catch ex As Exception
        '        MessageBox.Show("File not found or corrupted")
        '    End Try

        'End If


    End Sub

    Sub update_sets_g(myList As List(Of ADA_Set))

        For Each ada_set In myList
            Call get_global_set()  'update global variables

            '--------- enter global settings values -------------
            ada_set.Update_globals(Voltage, efficiency, Power_factor, include_ADA_brackets, Use_4pts_IO, Use_8pts_IO, Use_16pts_IO, removal_msbox_t, M23_at_IO_Receptacle,
    M23_Bulkhead_at_ADA, Use_4pts_IO_RB, Use_8pts_IO_RB, Use_16pts_IO_RB_Inputs, Use_16pts_IO_RB_Outputs, Single_Channel_IO_1_Input_per_RB, Single_Channel_IO_1_Output_per_RB, Splitter_Perc_dual_ch, Valve_solenoid_adapter, Smart_Wire_Darwin_Limited,
    Smart_Wire_Darwin_Full, EthernetIP_Full, Percent_overage_inputs, Percent_overage_nm_outputs, Percent_overage_m_outputs)
            '---------------------------------------------------

            ada_set.Calculate_Starter_Panel(Motor_horsepower, VFD_horsepower, Starters_values, value_to_index_480, value_to_index_575, Power_supply_480, Power_supply_575)
            ada_set.Calculate_IO_Panel()
            ada_set.Calculate_PLC_Panel()
            ada_set.Calculate_Field_Panel()
            ada_set.Calculate_Scanners()


        Next
        Call M12_process()
        Call Cal_TOTAL_values()
    End Sub

    '///////////////////---------- This update all our sets when we change global settings. Neccesary too keep all sets with the same global settings

    Private Sub Efficiency_box_TextChanged(sender As Object, e As EventArgs) Handles Efficiency_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub Power_factor_box_TextChanged(sender As Object, e As EventArgs) Handles Power_factor_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub ADA_Brackets_box_TextChanged(sender As Object, e As EventArgs) Handles ADA_Brackets_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub Use_4ptsIO_box_TextChanged(sender As Object, e As EventArgs) Handles Use_4ptsIO_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub Use_8ptsIO_box_TextChanged(sender As Object, e As EventArgs) Handles Use_8ptsIO_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub Use_16ptsIO_box_TextChanged(sender As Object, e As EventArgs) Handles Use_16ptsIO_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub removal_msbox_TextChanged(sender As Object, e As EventArgs) Handles removal_msbox.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub M23_IO_box_TextChanged(sender As Object, e As EventArgs) Handles M23_IO_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub M23_bulkhead_box_TextChanged(sender As Object, e As EventArgs) Handles M23_bulkhead_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub pts_4_rb_box_TextChanged(sender As Object, e As EventArgs) Handles pts_4_rb_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub pts_8_rb_box_TextChanged(sender As Object, e As EventArgs) Handles pts_8_rb_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub pts16_recep_inputs_box_TextChanged(sender As Object, e As EventArgs) Handles pts16_recep_inputs_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub pts16_recep_outputs_box_TextChanged(sender As Object, e As EventArgs) Handles pts16_recep_outputs_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub Single_input_box_TextChanged(sender As Object, e As EventArgs) Handles Single_input_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub single_output_box_TextChanged(sender As Object, e As EventArgs) Handles single_output_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub splitter_per_box_TextChanged(sender As Object, e As EventArgs) Handles splitter_per_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub valve_box_TextChanged(sender As Object, e As EventArgs) Handles valve_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub swd_limited_box_TextChanged(sender As Object, e As EventArgs) Handles swd_limited_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub swd_full_box_TextChanged(sender As Object, e As EventArgs) Handles swd_full_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub ethernetip_box_TextChanged(sender As Object, e As EventArgs) Handles ethernetip_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub percentage_inputs_box_TextChanged(sender As Object, e As EventArgs) Handles percentage_inputs_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub percentage_nmoutputs_box_TextChanged(sender As Object, e As EventArgs) Handles percentage_nmoutputs_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    Private Sub percentage_moutputs_box_TextChanged(sender As Object, e As EventArgs) Handles percentage_moutputs_box.TextChanged
        If ComboBox1.Items.Count > 0 Then
            Call update_sets_g(ADA_set_list)
        End If
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub StartQuotingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartQuotingToolStripMenuItem.Click
        ' ----------------  Display qty of parts of all sets by type -------------------
        Call Cal_PartCost("Starter Panel")

    End Sub

    Private Sub PR_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles PR_grid.CellValueChanged
        '--- when we change something in the Purchase request grid, the number of rows and total parts will be uodated
        Call Total_rows()

    End Sub

    Private Sub Install_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Install_grid.CellValueChanged
        Call Cal_install()
    End Sub

    Sub Cal_install()

        'Spreadsheet like calculations installation grid

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

                    Install_grid.Rows(i).Cells(3).Value() = Install_grid.Rows(i).Cells(1).Value() * Install_grid.Rows(i).Cells(2).Value() / 10
                End If
            End If

            '-----------------------
            If (IsNumeric(Install_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value())) Then
                Install_grid.Rows(i).Cells(3).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(3).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value()) = False, 0, Install_grid.Rows(i).Cells(3).Value())
                Install_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(4).Value()) = False, 0, Install_grid.Rows(i).Cells(4).Value())

                Install_grid.Rows(i).Cells(5).Value() = Install_grid.Rows(i).Cells(4).Value() * Install_grid.Rows(i).Cells(3).Value()
            End If

            '----------------------
            If (IsNumeric(Install_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value())) Then

                Install_grid.Rows(i).Cells(6).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(6).Value()) = False, 0, Install_grid.Rows(i).Cells(6).Value())
                Install_grid.Rows(i).Cells(3).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(3).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(3).Value()) = False, 0, Install_grid.Rows(i).Cells(3).Value())

                Install_grid.Rows(i).Cells(7).Value() = Install_grid.Rows(i).Cells(6).Value() * Install_grid.Rows(i).Cells(3).Value()

            End If

            '-----------------------------
            If (IsNumeric(Install_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value())) Then

                Install_grid.Rows(i).Cells(8).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(8).Value()) = False, 0, Install_grid.Rows(i).Cells(8).Value())
                Install_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(Install_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(Install_grid.Rows(i).Cells(1).Value()) = False, 0, Install_grid.Rows(i).Cells(1).Value())

                Install_grid.Rows(i).Cells(9).Value() = Install_grid.Rows(i).Cells(8).Value() * Install_grid.Rows(i).Cells(1).Value()
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

                Install_grid.Rows(i).Cells(12).Value() = Install_grid.Rows(i).Cells(5).Value() + Install_grid.Rows(i).Cells(7).Value() + Install_grid.Rows(i).Cells(9).Value() + Install_grid.Rows(i).Cells(11).Value()
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



        If totals_grid.Rows.Count > 20 Then
            totals_grid.Rows(20).Cells(1).Value() = labor
            totals_grid.Rows(21).Cells(1).Value() = materials
            totals_grid.Rows(22).Cells(1).Value() = expenses
            totals_grid.Rows(23).Cells(1).Value() = subcontract
        End If


    End Sub

    Private Sub MyVariablesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MyVariables.Visible = True
    End Sub



    Private Sub PurchaseRequestFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PurchaseRequestFormToolStripMenuItem.Click
        'export Purchase Request to excel file
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
                xlWorkSheet.Range("C:C").ColumnWidth = 70
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

                MessageBox.Show("Purchase Request Created Succesfully!")
            End If
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'show variables

        If ComboBox1.Items.Count > 0 Then

            Dim z As Integer : z = Get_set_number(ADA_set_list, ADA_name.Text)
            If z > -1 Then

                MyVariables.variables_box.Text = ("ADA Set Name: " & ADA_set_list.Item(z).myName & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
           "--------------------------VARIABLES USED IN STARTER PANEL-------------------" & vbCrLf & vbCrLf &
           "Panel bool: " & ADA_set_list.Item(z).Panel_bool & vbCrLf & vbCrLf & "Total current per group: " & ADA_set_list.Item(z).Total_current_per_group & vbCrLf & vbCrLf & "Total current: " & ADA_set_list.Item(z).total_current & vbCrLf & vbCrLf &
          "Non reversing starters: " & ADA_set_list.Item(z).non_reversing_starters & vbCrLf & vbCrLf & "reversing starters: " & ADA_set_list.Item(z).reversing_starters & vbCrLf & vbCrLf & "VFD Starter: " & ADA_set_list.Item(z).VFD_starters & vbCrLf & vbCrLf &
           "Soft VFD starters: " & ADA_set_list.Item(z).Soft_starters & vbCrLf & vbCrLf & "# starters: " & ADA_set_list.Item(z).Starters & vbCrLf & vbCrLf & "PLC: " & ADA_set_list.Item(z).PLC & vbCrLf & vbCrLf &
           "DC amps on main PS: " & ADA_set_list.Item(z).DC_amps_main_PS & vbCrLf & vbCrLf & "DC amps on auxiliary PS: " & ADA_set_list.Item(z).DC_amps_aux_PS & vbCrLf & vbCrLf & "Primary PS 5A: " & ADA_set_list.Item(z).Primary_PS_5A & vbCrLf & vbCrLf &
           "Primary PS 10A: " & ADA_set_list.Item(z).Primary_PS_10A & vbCrLf & vbCrLf & "Primary PS 20A: " & ADA_set_list.Item(z).Primary_PS_20A & vbCrLf & vbCrLf & "Primary PS 40A: " & ADA_set_list.Item(z).Primary_PS_40A & vbCrLf & vbCrLf &
           "Aux PS 5A: " & ADA_set_list.Item(z).Aux_PS_5A & vbCrLf & vbCrLf & "Aux PS 10A: " & ADA_set_list.Item(z).Aux_PS_10A & vbCrLf & vbCrLf & "Aux PS 20A: " & ADA_set_list.Item(z).Aux_PS_20A & vbCrLf & vbCrLf &
           "Aux PS 40A: " & ADA_set_list.Item(z).Aux_PS_40A & vbCrLf & vbCrLf & "busbar slots: " & ADA_set_list.Item(z).busbar_slots & vbCrLf & vbCrLf & "panel space top: " & ADA_set_list.Item(z).Panel_Space_Top & vbCrLf & vbCrLf &
          "panel space bottom: " & ADA_set_list.Item(z).Panel_Space_Bottom & vbCrLf & vbCrLf & "big panel: " & ADA_set_list.Item(z).big_panels) & vbCrLf & vbCrLf & vbCrLf & "total swire nodes: " & ADA_set_list.Item(z).Total_swire_nodes & vbCrLf & vbCrLf &
           "non reverse amps: " & ADA_set_list.Item(z).non_reversing_amps_p & vbCrLf & vbCrLf & vbCrLf & "reverse amps: " & ADA_set_list.Item(z).reversing_amps_p & vbCrLf & vbCrLf & "vfd amps: " & ADA_set_list.Item(z).vfd_amps_p & vbCrLf & vbCrLf & "soft amps: " & ADA_set_list.Item(z).soft_amps_p & vbCrLf & vbCrLf & vbCrLf &
          "--------------------------VARIABLES USED IN I/O PANEL-------------------" & vbCrLf & vbCrLf &
          "Inputs: " & ADA_set_list.Item(z).Inputs & vbCrLf & vbCrLf & "Monitoring Inputs: " & ADA_set_list.Item(z).Monitoring_Inputs & vbCrLf & vbCrLf & "Non Motion Outputs: " & ADA_set_list.Item(z).Non_Motion_Outputs & vbCrLf & vbCrLf &
          "Motion outputs: " & ADA_set_list.Item(z).Motion_outputs & vbCrLf & vbCrLf & "Has wire?: " & ADA_set_list.Item(z).has_wire & vbCrLf & vbCrLf & "System power require: " & ADA_set_list.Item(z).system_power_req & vbCrLf & vbCrLf &
           "----------------- I/O 4 Points Cards  ------------------" & vbCrLf & vbCrLf &
           "Inputs: " & ADA_set_list.Item(z).IO_4_inputs & vbCrLf & vbCrLf &
           "Non Motion: " & ADA_set_list.Item(z).IO_4_non_motion & vbCrLf & vbCrLf &
           "Motion: " & ADA_set_list.Item(z).IO_4_motion & vbCrLf & vbCrLf &
           "------- Are 4-port IN's used? -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_4 & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_4_convert_4 & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_4_convert_8 & vbCrLf & vbCrLf &
           "ADA-SWD-FF-Coupling: " & ADA_set_list.Item(z).IO_4_ADA_SWD_FF_Coupling & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_4_Pre_combine_8s & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_4_total_JB & vbCrLf & vbCrLf &
            "------- 4-port IN's used: -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_4_pigtail_single & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_4_pigtail_single & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_4_pigtail_single & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_4_pigtail_single & vbCrLf & vbCrLf &
           "------- Are 4-port OUT's used? -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_4_out & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_4_convert_4_out & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_4_convert_8_out & vbCrLf & vbCrLf &
           "convert 16: " & ADA_set_list.Item(z).IO_4_convert_16_out & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_4_Pre_combine_8s_out & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_4_total_JB_out & vbCrLf & vbCrLf &
           "------- 4-port OUT's used: -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_4_pigtail_single_out & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_4_conn_At_JB_single_out & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_4_pigtail_double_out & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_4_conn_At_JB_double_out & vbCrLf & vbCrLf &
           "----------------- I/O 8 Points Cards  ------------------" & vbCrLf & vbCrLf &
            "Inputs: " & ADA_set_list.Item(z).IO_8_inputs & vbCrLf & vbCrLf &
           "Non Motion: " & ADA_set_list.Item(z).IO_8_non_motion & vbCrLf & vbCrLf &
           "Motion: " & ADA_set_list.Item(z).IO_8_motion & vbCrLf & vbCrLf &
           "------- INPUTS -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_8 & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_8_convert_4 & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_8_convert_8 & vbCrLf & vbCrLf &
           "convert 16: " & ADA_set_list.Item(z).IO_8_convert_16 & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_8_Pre_combine_8s & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_8_total_JB & vbCrLf & vbCrLf &
            "------- 8 Port Input Receptacles Used -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_8_pigtail_single & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_8_conn_At_JB_single & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_8_pigtail_double & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_8_conn_At_JB_double & vbCrLf & vbCrLf &
           "------- OUTPUTS -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_8_out & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_8_convert_4_out & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_8_convert_8_out & vbCrLf & vbCrLf &
           "convert 16: " & ADA_set_list.Item(z).IO_8_convert_16_out & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_8_Pre_combine_8s_out & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_8_total_JB_out & vbCrLf & vbCrLf &
           "------- 8 Port Output Receptacles Used -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_8_pigtail_single_out & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_8_conn_At_JB_single_out & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_8_pigtail_double_out & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_8_conn_At_JB_double_out & vbCrLf & vbCrLf &
            "----------------- I/O 16 Points Cards  ------------------" & vbCrLf & vbCrLf &
            "Inputs: " & ADA_set_list.Item(z).IO_16_inputs & vbCrLf & vbCrLf &
           "Non Motion: " & ADA_set_list.Item(z).IO_16_non_motion & vbCrLf & vbCrLf &
           "Motion: " & ADA_set_list.Item(z).IO_16_motion & vbCrLf & vbCrLf &
           "------- INPUTS -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_16 & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_16_convert_4 & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_16_convert_8 & vbCrLf & vbCrLf &
           "convert 16: " & ADA_set_list.Item(z).IO_16_convert_16 & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_16_Pre_combine_8s & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_16_total_JB & vbCrLf & vbCrLf &
            "------- 16 Port Input Receptacles Used -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_16_pigtail_single & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_16_conn_At_JB_single & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_16_pigtail_double & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_16_conn_At_JB_double & vbCrLf & vbCrLf &
           "------- OUTPUTS -----" & vbCrLf & vbCrLf &
           "Totals I/O: " & ADA_set_list.Item(z).Totals_IO_16_out & vbCrLf & vbCrLf &
           "convert 4: " & ADA_set_list.Item(z).IO_16_convert_4_out & vbCrLf & vbCrLf &
           "convert 8: " & ADA_set_list.Item(z).IO_16_convert_8_out & vbCrLf & vbCrLf &
           "convert 16: " & ADA_set_list.Item(z).IO_16_convert_16_out & vbCrLf & vbCrLf &
           "Pre-combine-8's: " & ADA_set_list.Item(z).IO_16_Pre_combine_8s_out & vbCrLf & vbCrLf &
           "Totals J/B: " & ADA_set_list.Item(z).IO_16_total_JB_out & vbCrLf & vbCrLf &
           "------- 16 Port Output Receptacles Used -----" & vbCrLf & vbCrLf &
           "Pigtail Single: " & ADA_set_list.Item(z).IO_16_pigtail_single_out & vbCrLf & vbCrLf &
           "Conn at JB Single: " & ADA_set_list.Item(z).IO_16_conn_At_JB_single_out & vbCrLf & vbCrLf &
           "Pigtail Double: " & ADA_set_list.Item(z).IO_16_pigtail_double_out & vbCrLf & vbCrLf &
           "Conn at JB Double: " & ADA_set_list.Item(z).IO_16_conn_At_JB_double_out & vbCrLf & vbCrLf


                MyVariables.Visible = True
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Parts needed and cost for specified ADA set. default is (Starter panel)
        Parts_needed.Text = ADA_name.Text
        Dim myIndex As Integer : myIndex = Get_set_number(ADA_set_list, ADA_name.Text)

        If myIndex > -1 Then
            Call my_ADA_parts("Starter Panel", myIndex)
        End If
    End Sub


    Sub my_ADA_parts(type As String, myIndex As Integer)

        '////////////////////////   This sub display the parts qty of one particular type in the set specified ///////////////////////////
        Parts_needed.Visible = True
        Parts_needed.Parts_need_grid.Rows.Clear()

        Dim total_cost As Decimal : total_cost = 0
        Dim total_q As Double : total_q = 0
        Dim labor_cost As Double : labor_cost = 0
        Dim minus_labor As Double : minus_labor = 0


        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT ADA_number, Part_Name, Primary_Vendor from ada_active_parts"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read

                    If (get_total_byset(reader(0), type, myIndex) > 0) Then

                        Parts_needed.Parts_need_grid.Rows.Add(New DataGridViewRow)
                        Parts_needed.Parts_need_grid.Rows(Parts_needed.Parts_need_grid.Rows.Count - 1).Cells(0).Value = reader(0)  'ADA Number                        
                        Parts_needed.Parts_need_grid.Rows(Parts_needed.Parts_need_grid.Rows.Count - 1).Cells(1).Value = reader(1)  'part name           
                        Parts_needed.Parts_need_grid.Rows(Parts_needed.Parts_need_grid.Rows.Count - 1).Cells(2).Value = get_total_byset(reader(0), type, myIndex) : total_q = total_q + get_total_byset(reader(0), type, myIndex)
                        Parts_needed.Parts_need_grid.Rows(Parts_needed.Parts_need_grid.Rows.Count - 1).Cells(3).Value = reader(2)  'vendor
                    End If
                End While
            End If

            reader.Close()

            For Each row As DataGridViewRow In Parts_needed.Parts_need_grid.Rows
                row.Cells(4).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(3).Value)
                row.Cells(5).Value = row.Cells(4).Value * row.Cells(2).Value : total_cost = total_cost + row.Cells(5).Value

                If String.Equals(row.Cells(0).Value, "ADA-PSL-IO") = True Or String.Equals(row.Cells(0).Value, "ADA-PSL-MS") = True Or String.Equals(row.Cells(0).Value, "ADA-PSL-PLC") = True Or
                     String.Equals(row.Cells(0).Value, "2PBX Labor") = True Or String.Equals(row.Cells(0).Value, "3PBX Labor") = True Or
                    String.Equals(row.Cells(0).Value, "1PBX E-Stop Labor") = True Or String.Equals(row.Cells(0).Value, "1PBX Labor") = True Or String.Equals(row.Cells(0).Value, "1ERP Labor") = True Then

                    labor_cost = labor_cost + row.Cells(5).Value
                    minus_labor = minus_labor + row.Cells(2).Value

                End If
                row.Height = 52
            Next

            Parts_needed.p_label.Text = "Total Parts Needed:  " & total_q - minus_labor
            Parts_needed.c_label.Text = "Total Cost:  $" & Math.Round(total_cost, 2)
            Parts_needed.m_label.Text = "Material Cost: $" & Math.Round(total_cost - labor_cost, 2)
            Parts_needed.l_label.Text = "Labor Cost: $" & Math.Round(labor_cost, 2)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    '/////////////////////////////////// Generate Part cost summary //////////////////////////////////
    Sub Cal_PartCost(type As String)

        PartCost_summary.Visible = True
        PartCost_summary.Parts_cost_t_grid.Rows.Clear()  'Clear partcost grid

        Dim total_cost As Decimal : total_cost = 0
        Dim total_q As Double : total_q = 0
        Dim labor_cost As Double : labor_cost = 0
        Dim minus_labor As Double : minus_labor = 0

        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT ADA_number, Part_Name, Primary_Vendor from ada_active_parts"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read

                    If (get_total_bytype(reader(0), ADA_set_list, type) > 0) Then

                        PartCost_summary.Parts_cost_t_grid.Rows.Add(New DataGridViewRow)
                        PartCost_summary.Parts_cost_t_grid.Rows(PartCost_summary.Parts_cost_t_grid.Rows.Count - 1).Cells(0).Value = reader(0)  'ADA Number                        
                        PartCost_summary.Parts_cost_t_grid.Rows(PartCost_summary.Parts_cost_t_grid.Rows.Count - 1).Cells(1).Value = reader(1)  'part name           
                        PartCost_summary.Parts_cost_t_grid.Rows(PartCost_summary.Parts_cost_t_grid.Rows.Count - 1).Cells(2).Value = get_total_bytype(reader(0), ADA_set_list, type) : total_q = total_q + get_total_bytype(reader(0), ADA_set_list, type)
                        PartCost_summary.Parts_cost_t_grid.Rows(PartCost_summary.Parts_cost_t_grid.Rows.Count - 1).Cells(3).Value = reader(2)  'vendor
                    End If
                End While
            End If

            reader.Close()

            For Each row As DataGridViewRow In PartCost_summary.Parts_cost_t_grid.Rows
                row.Cells(4).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(3).Value)
                row.Cells(5).Value = row.Cells(4).Value * row.Cells(2).Value : total_cost = total_cost + row.Cells(5).Value
                If String.Equals(row.Cells(0).Value, "ADA-PSL-IO") = True Or String.Equals(row.Cells(0).Value, "ADA-PSL-MS") = True Or String.Equals(row.Cells(0).Value, "ADA-PSL-PLC") = True Or
                     String.Equals(row.Cells(0).Value, "2PBX Labor") = True Or String.Equals(row.Cells(0).Value, "3PBX Labor") = True Or
                    String.Equals(row.Cells(0).Value, "1PBX E-Stop Labor") = True Or String.Equals(row.Cells(0).Value, "1PBX Labor") = True Or String.Equals(row.Cells(0).Value, "1ERP Labor") = True Then

                    labor_cost = labor_cost + row.Cells(5).Value
                    minus_labor = minus_labor + row.Cells(2).Value

                End If
                row.Height = 52
            Next

            PartCost_summary.pt_label.Text = "Parts Needed: " & total_q - minus_labor
            PartCost_summary.ct_label.Text = "Cost: $" & Math.Round(total_cost, 2)
            PartCost_summary.mt_label.Text = "Material Cost: $" & Math.Round(total_cost - labor_cost, 2)
            PartCost_summary.lt_label.Text = "Labor Cost: $" & Math.Round(labor_cost, 2)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub


    Private Sub FindAlternativePartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindAlternativePartToolStripMenuItem.Click
        'Replace the select part with an alternative if found

        If PR_grid.CurrentCell.Value.ToString <> String.Empty Then

            row_in = Find_index(PR_grid.CurrentCell.Value.ToString)
            myQTY = PR_grid.Rows(row_in).Cells(7).Value

            If row_in > -1 Then

                Call Show_alternatives(PR_grid.CurrentCell.Value, PR_grid.Rows(row_in).Cells(1).Value)

            End If
        End If
    End Sub

    Function Find_index(name As String) As Integer

        'Find and return row index of ADA part in datagrid

        Find_index = -1

        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If String.Compare(row.Cells.Item("ADA_number").Value.ToString, name) = 0 Then
                Find_index = row.Index
                Exit For
            End If
        Next

    End Function

    Sub Show_alternatives(ADA_part As String, name_part As String)

        ADA_Alternatives.Visible = True
        ADA_Alternatives.alt_grid.Rows.Clear()

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@ADA_name", ADA_part)
            cmd.Parameters.AddWithValue("@name_part", name_part)
            cmd.CommandText = "select Legacy_ADA_Number, Part_Name, Part_Description, Primary_Vendor from parts_table where LEGACY_ADA_Number = @ADA_name and Part_Name <> @name_part"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    ADA_Alternatives.alt_grid.Rows.Add(New DataGridViewRow)
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(0).Value = reader(0)  'ADA Number                        
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(1).Value = reader(1)  'part name
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(2).Value = reader(2)  'Part description
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(3).Value = reader(3)  'vendor
                End While
            End If

            reader.Close()

            For Each row As DataGridViewRow In ADA_Alternatives.alt_grid.Rows
                If row.IsNewRow Then Continue For
                row.Cells(4).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(3).Value)
                row.Height = 52
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub PR_grid_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles PR_grid.CellMouseDown
        '========= LEFT CLICK SELECT A CELL ===============
        If e.Button = MouseButtons.Right Then
            Try
                PR_grid.CurrentCell = PR_grid.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub InstallationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstallationToolStripMenuItem.Click
        'Place the values in the installation sheet. Calculate all the formulas

        Call Cal_total_Setup()
        'get values

        Install_grid.Rows(1).Cells(1).Value() = t_inputs_array(11) + t_inputs_array(12) + t_inputs_array(13) + t_inputs_array(14) + t_inputs_array(15) + t_inputs_array(16) + t_inputs_array(17)  'PE
        Install_grid.Rows(2).Cells(1).Value() = t_inputs_array(1) + t_inputs_array(2) 'Pull cord stations
        Install_grid.Rows(3).Cells(1).Value() = t_inputs_array(3) * 1.1  'Pull Cord Cable/eyelit (per foot)
        Install_grid.Rows(4).Cells(1).Value() = t_inputs_array(4) * 2.5 'Pull Cord Pulley Kits
        Install_grid.Rows(5).Cells(1).Value() = t_inputs_array(5) + t_inputs_array(6) + t_push_array(0) + t_push_array(0) + t_push_array(1) + t_push_array(2) + t_push_array(3) _
        + t_push_array(4) + t_push_array(5) + t_push_array(6) + t_push_array(7) + t_push_array(8) + t_push_array(9) + t_push_array(16) + t_push_array(17) + t_push_array(18)  'OS Station
        Install_grid.Rows(6).Cells(1).Value() = t_inputs_array(9) + t_inputs_array(10) + t_inputs_array(17)  'Limit/Prox Switch
        Install_grid.Rows(7).Cells(1).Value() = t_inputs_array(8) 'Pressure Switch
        Install_grid.Rows(8).Cells(1).Value() = t_lights_array(5) + t_lights_array(6) + t_lights_array(7) + t_lights_array(8) 'Beacons
        Install_grid.Rows(9).Cells(1).Value() = t_lights_array(9) + t_lights_array(10) 'horns
        Install_grid.Rows(10).Cells(1).Value() = t_lights_array(13) + t_lights_array(14) 'solenoids
        Install_grid.Rows(12).Cells(1).Value() = t_inputs_array(7) 'interlock boxes
        Install_grid.Rows(15).Cells(1).Value() = t_field_array(9) + t_field_array(10) + t_field_array(11)  'encoder

    End Sub

    Sub Cal_total_Setup()
        'calculate total setup values

        t_NEMA12 = 0 : t_NEMA4X = 0 : t_PLC_Gatewat_PGW = 0
        t_control_logic = 0 : t_compact_logic = 0 : t_PLC_UPS = 0 : t_Amp_Mon = 0
        t_Remote_IO = 0 : t_total_current_group = 0 : t_mdr = 0
        t_ADA_NEMA12_BOX24_No_Starters = 0 : t_ADA_NEMA12_BOX30_No_Starters = 0 : t_motor_disconnect = 0

        'initialize variables with zero
        For i = 0 To 47
            t_Motor_array(i) = 0
        Next

        For i = 0 To 18
            t_inputs_array(i) = 0
        Next

        For i = 0 To 18
            t_push_array(i) = 0
        Next

        For i = 0 To 14
            t_lights_array(i) = 0
        Next

        For i = 0 To 11
            t_field_array(i) = 0
        Next

        For i = 0 To 11
            t_bus_array(i) = 0
        Next

        For i = 0 To 3
            t_scanner_array(i) = 0
        Next

        For i = 0 To 14
            t_brakes_array(i) = 0
        Next


        For Each ada_set In ADA_set_list

            t_NEMA12 = t_NEMA12 + CType(ada_set.NEMA12, Double)
            t_NEMA4X = t_NEMA4X + CType(ada_set.NEMA4X, Double)
            t_PLC_Gatewat_PGW = t_PLC_Gatewat_PGW + CType(ada_set.PLC_GateWay_PGW, Double)
            t_control_logic = t_control_logic + CType(ada_set.ControlLogix_PLC_Box, Double)
            t_compact_logic = t_compact_logic + CType(ada_set.CompactLogix_PLC_Box, Double)
            t_PLC_UPS = t_PLC_UPS + CType(ada_set.PLC_UPS_Battery_Backup, Double)
            t_Amp_Mon = t_Amp_Mon + CType(ada_set.Amp_Mon_Electro, Double)
            t_Remote_IO = t_Remote_IO + CType(ada_set.Remote_IO, Double)
            t_total_current_group = t_total_current_group + CType(ada_set.Total_current_per_group, Double)
            t_mdr = t_mdr + CType(ada_set.MDRs, Double)
            t_ADA_NEMA12_BOX24_No_Starters = t_ADA_NEMA12_BOX24_No_Starters + CType(ada_set.ADA_NEMA12_BOX24_No_Starters, Double)
            t_ADA_NEMA12_BOX30_No_Starters = t_ADA_NEMA12_BOX30_No_Starters + CType(ada_set.ADA_NEMA12_BOX30_No_Starters, Double)
            t_motor_disconnect = t_motor_disconnect + CType(ada_set.motor_disconnect, Double)

            'motors
            t_Motor_array(0) = t_Motor_array(0) + CType(ada_set.Motors_table(0, 0), Double) : t_Motor_array(1) = t_Motor_array(1) + CType(ada_set.Motors_table(1, 0), Double) : t_Motor_array(2) = t_Motor_array(2) + CType(ada_set.Motors_table(2, 0), Double)
            t_Motor_array(3) = t_Motor_array(3) + CType(ada_set.Motors_table(3, 0), Double) : t_Motor_array(4) = t_Motor_array(4) + CType(ada_set.Motors_table(4, 0), Double) : t_Motor_array(5) = t_Motor_array(5) + CType(ada_set.Motors_table(5, 0), Double)
            t_Motor_array(6) = t_Motor_array(6) + CType(ada_set.Motors_table(6, 0), Double) : t_Motor_array(7) = t_Motor_array(7) + CType(ada_set.Motors_table(7, 0), Double) : t_Motor_array(8) = t_Motor_array(8) + CType(ada_set.Motors_table(8, 0), Double)
            t_Motor_array(9) = t_Motor_array(9) + CType(ada_set.Motors_table(9, 0), Double) : t_Motor_array(10) = t_Motor_array(10) + CType(ada_set.Motors_table(10, 0), Double) : t_Motor_array(11) = t_Motor_array(11) + CType(ada_set.Motors_table(11, 0), Double)
            t_Motor_array(12) = t_Motor_array(12) + CType(ada_set.Motors_table(12, 0), Double) : t_Motor_array(13) = t_Motor_array(13) + CType(ada_set.Motors_table(13, 0), Double) : t_Motor_array(14) = t_Motor_array(14) + CType(ada_set.Motors_table(14, 0), Double)
            t_Motor_array(15) = t_Motor_array(15) + CType(ada_set.Motors_table(15, 0), Double) : t_Motor_array(16) = t_Motor_array(16) + CType(ada_set.Motors_table(16, 0), Double) : t_Motor_array(17) = t_Motor_array(17) + CType(ada_set.Motors_table(17, 0), Double)
            t_Motor_array(18) = t_Motor_array(18) + CType(ada_set.Motors_table(18, 0), Double) : t_Motor_array(19) = t_Motor_array(19) + CType(ada_set.Motors_table(19, 0), Double) : t_Motor_array(20) = t_Motor_array(20) + CType(ada_set.Motors_table(20, 0), Double)
            t_Motor_array(21) = t_Motor_array(21) + CType(ada_set.Motors_table(21, 0), Double) : t_Motor_array(22) = t_Motor_array(22) + CType(ada_set.Motors_table(22, 0), Double) : t_Motor_array(23) = t_Motor_array(23) + CType(ada_set.Motors_table(23, 0), Double)
            t_Motor_array(24) = t_Motor_array(24) + CType(ada_set.Motors_table(24, 0), Double) : t_Motor_array(25) = t_Motor_array(25) + CType(ada_set.Motors_table(25, 0), Double) : t_Motor_array(26) = t_Motor_array(26) + CType(ada_set.Motors_table(26, 0), Double)
            t_Motor_array(27) = t_Motor_array(27) + CType(ada_set.Motors_table(27, 0), Double) : t_Motor_array(28) = t_Motor_array(28) + CType(ada_set.Motors_table(28, 0), Double) : t_Motor_array(29) = t_Motor_array(29) + CType(ada_set.Motors_table(29, 0), Double)
            t_Motor_array(30) = t_Motor_array(30) + CType(ada_set.Motors_table(30, 0), Double) : t_Motor_array(31) = t_Motor_array(31) + CType(ada_set.Motors_table(31, 0), Double) : t_Motor_array(32) = t_Motor_array(32) + CType(ada_set.Motors_table(32, 0), Double)
            t_Motor_array(33) = t_Motor_array(33) + CType(ada_set.Motors_table(33, 0), Double) : t_Motor_array(34) = t_Motor_array(34) + CType(ada_set.Motors_table(34, 0), Double) : t_Motor_array(35) = t_Motor_array(35) + CType(ada_set.Motors_table(35, 0), Double)
            t_Motor_array(36) = t_Motor_array(36) + CType(ada_set.Motors_table(36, 0), Double) : t_Motor_array(37) = t_Motor_array(37) + CType(ada_set.Motors_table(37, 0), Double) : t_Motor_array(38) = t_Motor_array(38) + CType(ada_set.Motors_table(38, 0), Double)
            t_Motor_array(39) = t_Motor_array(39) + CType(ada_set.Motors_table(39, 0), Double) : t_Motor_array(40) = t_Motor_array(40) + CType(ada_set.Motors_table(40, 0), Double) : t_Motor_array(41) = t_Motor_array(41) + CType(ada_set.Motors_table(41, 0), Double)
            t_Motor_array(42) = t_Motor_array(42) + CType(ada_set.Motors_table(42, 0), Double) : t_Motor_array(43) = t_Motor_array(43) + CType(ada_set.Motors_table(43, 0), Double) : t_Motor_array(44) = t_Motor_array(44) + CType(ada_set.Motors_table(44, 0), Double)
            t_Motor_array(45) = t_Motor_array(45) + CType(ada_set.Motors_table(45, 0), Double) : t_Motor_array(46) = t_Motor_array(46) + CType(ada_set.Motors_table(46, 0), Double) : t_Motor_array(47) = t_Motor_array(47) + CType(ada_set.Motors_table(47, 0), Double)

            'inputs
            t_inputs_array(0) = t_inputs_array(0) + CType(ada_set.inputs_table(0, 0), Double) : t_inputs_array(1) = t_inputs_array(1) + CType(ada_set.inputs_table(1, 0), Double) : t_inputs_array(2) = t_inputs_array(2) + CType(ada_set.inputs_table(2, 0), Double)
            t_inputs_array(3) = t_inputs_array(3) + CType(ada_set.inputs_table(3, 0), Double) : t_inputs_array(4) = t_inputs_array(4) + CType(ada_set.inputs_table(4, 0), Double) : t_inputs_array(5) = t_inputs_array(5) + CType(ada_set.inputs_table(5, 0), Double)
            t_inputs_array(6) = t_inputs_array(6) + CType(ada_set.inputs_table(6, 0), Double) : t_inputs_array(7) = t_inputs_array(7) + CType(ada_set.inputs_table(7, 0), Double) : t_inputs_array(8) = t_inputs_array(8) + CType(ada_set.inputs_table(8, 0), Double)
            t_inputs_array(9) = t_inputs_array(9) + CType(ada_set.inputs_table(9, 0), Double) : t_inputs_array(10) = t_inputs_array(10) + CType(ada_set.inputs_table(10, 0), Double) : t_inputs_array(11) = t_inputs_array(11) + CType(ada_set.inputs_table(11, 0), Double)
            t_inputs_array(12) = t_inputs_array(12) + CType(ada_set.inputs_table(12, 0), Double) : t_inputs_array(13) = t_inputs_array(13) + CType(ada_set.inputs_table(13, 0), Double) : t_inputs_array(14) = t_inputs_array(14) + CType(ada_set.inputs_table(14, 0), Double)
            t_inputs_array(15) = t_inputs_array(15) + CType(ada_set.inputs_table(15, 0), Double) : t_inputs_array(16) = t_inputs_array(16) + CType(ada_set.inputs_table(16, 0), Double) : t_inputs_array(17) = t_inputs_array(17) + CType(ada_set.inputs_table(17, 0), Double)
            t_inputs_array(18) = t_inputs_array(18) + CType(ada_set.inputs_table(18, 0), Double)

            'push
            t_push_array(0) = t_push_array(0) + CType(ada_set.push_table(0, 0), Double) : t_push_array(1) = t_push_array(1) + CType(ada_set.push_table(1, 0), Double) : t_push_array(2) = t_push_array(2) + CType(ada_set.push_table(2, 0), Double)
            t_push_array(3) = t_push_array(3) + CType(ada_set.push_table(3, 0), Double) : t_push_array(4) = t_push_array(4) + CType(ada_set.push_table(4, 0), Double) : t_push_array(5) = t_push_array(5) + CType(ada_set.push_table(5, 0), Double)
            t_push_array(6) = t_push_array(6) + CType(ada_set.push_table(6, 0), Double) : t_push_array(7) = t_push_array(7) + CType(ada_set.push_table(7, 0), Double) : t_push_array(8) = t_push_array(8) + CType(ada_set.push_table(8, 0), Double)
            t_push_array(9) = t_push_array(9) + CType(ada_set.push_table(9, 0), Double) : t_push_array(10) = t_push_array(10) + CType(ada_set.push_table(10, 0), Double) : t_push_array(11) = t_push_array(11) + CType(ada_set.push_table(11, 0), Double)
            t_push_array(12) = t_push_array(12) + CType(ada_set.push_table(12, 0), Double) : t_push_array(13) = t_push_array(13) + CType(ada_set.push_table(13, 0), Double) : t_push_array(14) = t_push_array(14) + CType(ada_set.push_table(14, 0), Double)
            t_push_array(15) = t_push_array(15) + CType(ada_set.push_table(15, 0), Double) : t_push_array(16) = t_push_array(16) + CType(ada_set.push_table(16, 0), Double) : t_push_array(17) = t_push_array(17) + CType(ada_set.push_table(17, 0), Double)
            t_push_array(18) = t_push_array(18) + CType(ada_set.push_table(18, 0), Double)

            'lights

            t_lights_array(0) = t_lights_array(0) + CType(ada_set.lights_table(0, 0), Double) : t_lights_array(1) = t_lights_array(1) + CType(ada_set.lights_table(1, 0), Double) : t_lights_array(2) = t_lights_array(2) + CType(ada_set.lights_table(2, 0), Double)
            t_lights_array(3) = t_lights_array(3) + CType(ada_set.lights_table(3, 0), Double) : t_lights_array(4) = t_lights_array(4) + CType(ada_set.lights_table(4, 0), Double) : t_lights_array(5) = t_lights_array(5) + CType(ada_set.lights_table(5, 0), Double)
            t_lights_array(6) = t_lights_array(6) + CType(ada_set.lights_table(6, 0), Double) : t_lights_array(7) = t_lights_array(7) + CType(ada_set.lights_table(7, 0), Double) : t_lights_array(8) = t_lights_array(8) + CType(ada_set.lights_table(8, 0), Double)
            t_lights_array(9) = t_lights_array(9) + CType(ada_set.lights_table(9, 0), Double) : t_lights_array(10) = t_lights_array(10) + CType(ada_set.lights_table(10, 0), Double) : t_lights_array(11) = t_lights_array(11) + CType(ada_set.lights_table(11, 0), Double)
            t_lights_array(12) = t_lights_array(12) + CType(ada_set.lights_table(12, 0), Double) : t_lights_array(13) = t_lights_array(13) + CType(ada_set.lights_table(13, 0), Double) : t_lights_array(14) = t_lights_array(14) + CType(ada_set.lights_table(14, 0), Double)

            'field
            t_field_array(0) = t_field_array(0) + CType(ada_set.field_table(0, 0), Double) : t_field_array(1) = t_field_array(1) + CType(ada_set.field_table(1, 0), Double) : t_field_array(2) = t_field_array(2) + CType(ada_set.field_table(2, 0), Double)
            t_field_array(3) = t_field_array(3) + CType(ada_set.field_table(3, 0), Double) : t_field_array(4) = t_field_array(4) + CType(ada_set.field_table(4, 0), Double) : t_field_array(5) = t_field_array(5) + CType(ada_set.field_table(5, 0), Double)
            t_field_array(6) = t_field_array(6) + CType(ada_set.field_table(6, 0), Double) : t_field_array(7) = t_field_array(7) + CType(ada_set.field_table(7, 0), Double) : t_field_array(8) = t_field_array(8) + CType(ada_set.field_table(8, 0), Double)
            t_field_array(9) = t_field_array(9) + CType(ada_set.field_table(9, 0), Double) : t_field_array(10) = t_field_array(10) + CType(ada_set.field_table(10, 0), Double) : t_field_array(11) = t_field_array(11) + CType(ada_set.field_table(11, 0), Double)

            'bus
            t_bus_array(0) = t_bus_array(0) + CType(ada_set.bus_table(0, 0), Double) : t_bus_array(1) = t_bus_array(1) + CType(ada_set.bus_table(1, 0), Double) : t_bus_array(2) = t_bus_array(2) + CType(ada_set.bus_table(2, 0), Double)
            t_bus_array(3) = t_bus_array(3) + CType(ada_set.bus_table(3, 0), Double) : t_bus_array(4) = t_bus_array(4) + CType(ada_set.bus_table(4, 0), Double) : t_bus_array(5) = t_bus_array(5) + CType(ada_set.bus_table(5, 0), Double)
            t_bus_array(6) = t_bus_array(6) + CType(ada_set.bus_table(6, 0), Double) : t_bus_array(7) = t_bus_array(7) + CType(ada_set.bus_table(7, 0), Double) : t_bus_array(8) = t_bus_array(8) + CType(ada_set.bus_table(8, 0), Double)
            t_bus_array(9) = t_bus_array(9) + CType(ada_set.bus_table(9, 0), Double) : t_bus_array(10) = t_bus_array(10) + CType(ada_set.bus_table(10, 0), Double) : t_bus_array(11) = t_bus_array(11) + CType(ada_set.bus_table(11, 0), Double)

            'scanner
            t_scanner_array(0) = t_scanner_array(0) + CType(ada_set.bus_table(0, 0), Double) : t_scanner_array(1) = t_scanner_array(1) + CType(ada_set.bus_table(1, 0), Double) : t_scanner_array(2) = t_scanner_array(2) + CType(ada_set.bus_table(2, 0), Double)
            t_scanner_array(3) = t_scanner_array(3) + CType(ada_set.bus_table(3, 0), Double)

            'brakes
            t_brakes_array(0) = t_brakes_array(0) + CType(ada_set.brakes_table(0, 0), Double) : t_brakes_array(1) = t_brakes_array(1) + CType(ada_set.brakes_table(1, 0), Double) : t_brakes_array(2) = t_brakes_array(2) + CType(ada_set.brakes_table(2, 0), Double)
            t_brakes_array(3) = t_brakes_array(3) + CType(ada_set.brakes_table(3, 0), Double) : t_brakes_array(4) = t_brakes_array(4) + CType(ada_set.brakes_table(4, 0), Double) : t_brakes_array(5) = t_brakes_array(5) + CType(ada_set.brakes_table(5, 0), Double)
            t_brakes_array(6) = t_brakes_array(6) + CType(ada_set.brakes_table(6, 0), Double) : t_brakes_array(7) = t_brakes_array(7) + CType(ada_set.brakes_table(7, 0), Double) : t_brakes_array(8) = t_brakes_array(8) + CType(ada_set.brakes_table(8, 0), Double)
            t_brakes_array(9) = t_brakes_array(9) + CType(ada_set.brakes_table(9, 0), Double) : t_brakes_array(10) = t_brakes_array(10) + CType(ada_set.brakes_table(10, 0), Double) : t_brakes_array(11) = t_brakes_array(11) + CType(ada_set.brakes_table(11, 0), Double)
            t_brakes_array(12) = t_brakes_array(12) + CType(ada_set.brakes_table(12, 0), Double) : t_brakes_array(13) = t_brakes_array(13) + CType(ada_set.brakes_table(13, 0), Double) : t_brakes_array(14) = t_brakes_array(14) + CType(ada_set.brakes_table(14, 0), Double)

        Next

    End Sub

    Private Sub ClearInstallationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearInstallationToolStripMenuItem.Click

        'clear install grid
        For Each row As DataGridViewRow In Install_grid.Rows

            If row.IsNewRow Then Continue For
            row.Cells(1).Value = "" : row.Cells(2).Value = "" : row.Cells(3).Value = "" : row.Cells(4).Value = ""
            row.Cells(5).Value = "" : row.Cells(6).Value = "" : row.Cells(7).Value = "" : row.Cells(8).Value = ""
            row.Cells(9).Value = "" : row.Cells(10).Value = "" : row.Cells(11).Value = "" : row.Cells(12).Value = ""

        Next
    End Sub

    Private Sub TotalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TotalToolStripMenuItem.Click
        Setup_total.Visible = True
    End Sub


    Sub Cal_TOTAL_values()
        'This calculate the TOTAL values that goes in the TOTAL tab
        totals_grid.Rows(0).Cells(1).Value() = "$" & Cal_specific_totals("Starter Panel", 1)
        totals_grid.Rows(1).Cells(1).Value() = "$" & Cal_specific_totals("Starter Panel", 2)
        totals_grid.Rows(2).Cells(1).Value() = "$" & Cal_specific_totals("IO Panel", 1)
        totals_grid.Rows(3).Cells(1).Value() = "$" & Cal_specific_totals("IO Panel", 2)
        totals_grid.Rows(4).Cells(1).Value() = "$" & Cal_specific_totals("PLC Panel", 1)
        totals_grid.Rows(5).Cells(1).Value() = "$" & Cal_specific_totals("PLC Panel", 2)
        totals_grid.Rows(7).Cells(1).Value() = "$" & Cal_specific_totals("Field parts", 1)
        totals_grid.Rows(8).Cells(1).Value() = "$" & Cal_specific_totals("Field parts", 2)
        totals_grid.Rows(9).Cells(1).Value() = "$" & Cal_specific_totals("Scanners", 1)
        totals_grid.Rows(10).Cells(1).Value() = "$" & Cal_specific_totals("Scanners", 2)
        totals_grid.Rows(11).Cells(1).Value() = "$" & Cal_specific_totals("M12 Cables", 1)
        totals_grid.Rows(12).Cells(1).Value() = "$" & Cal_specific_totals("M12 Cables", 2)
        totals_grid.Rows(13).Cells(1).Value() = "$" & Cal_specific_totals("M12 ES Cables", 1)
        totals_grid.Rows(14).Cells(1).Value() = "$" & Cal_specific_totals("M12 ES Cables", 2)

        totals_grid.Rows(16).Cells(1).Value() = Cal_est_shipping("I/O")  'SHIPPING & HANDLING I/O
        totals_grid.Rows(17).Cells(1).Value() = Cal_est_shipping("Motor")  'SHIPPING & HANDLING Motor Big
        totals_grid.Rows(18).Cells(1).Value() = Cal_est_shipping("Scanners")  'SHIPPING & HANDLING – Scanners

    End Sub


    Function Cal_specific_totals(type As String, index As Integer) As Double

        'This will calculate the general total cost according to the type. Values found in ADA Panels and options sections in TOTAL tabs
        'index 1 is materials, 2 is labor and 3 totals

        Cal_specific_totals = 0
        'fill a dictionary with the ADA names and vendor
        Dim temp_list As New Dictionary(Of String, String)
        Dim temp_name As New Dictionary(Of String, String)

        Dim total_q As Double : total_q = 0  'total
        Dim labor_q As Double : labor_q = 0 'labor

        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT distinct ADA_Number ,Part_Name, primary_vendor from ada_active_parts where ada_number is not null"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader


            If reader.HasRows Then
                While reader.Read

                    If temp_list.ContainsKey(reader(0)) = False Then
                        temp_list.Add(reader(0), reader(1))  'ada name and part name
                        temp_name.Add(reader(0), reader(2))  'ada name and vendor
                    End If

                End While
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        '-- start total cost calculations once we have our ada parts and respective vendors

        For Each kvp As KeyValuePair(Of String, String) In temp_list.ToArray


            If (get_total_bytype(kvp.Key, ADA_set_list, type) > 0) Then
                total_q = total_q + (get_total_bytype(kvp.Key, ADA_set_list, type)) * Form1.Get_Latest_Cost(Login.Connection, temp_list.Item(kvp.Key), temp_name.Item(kvp.Key))
            End If

        Next

        labor_q = get_total_bytype("ADA-PSL-IO", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Panel Shop Labor I/O box", "Atronix") +
            get_total_bytype("ADA-PSL-MS", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Panel Shop Labor Motor Starter Box", "Atronix") +
            get_total_bytype("ADA-PSL-PLC", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Panel Shop Labor PLC Box", "Atronix") _
            + get_total_bytype("2PBX Labor", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Wiring and assembly 2PBX", "Atronix") +
            get_total_bytype("3PBX Labor", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Wiring and assembly 3PBX", "Atronix") +
            get_total_bytype("1PBX E-Stop Labor", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Wiring and assembly 1PBX-E", "Atronix") +
            get_total_bytype("1PBX Labor", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Wiring and assembly 1PBX", "Atronix") +
            get_total_bytype("1ERP Labor", ADA_set_list, type) * Form1.Get_Latest_Cost(Login.Connection, "Wiring and assembly ROPE PULL", "Atronix")


        If index = 1 Then
            Cal_specific_totals = total_q - labor_q
        ElseIf index = 2 Then
            Cal_specific_totals = labor_q
        Else
            Cal_specific_totals = total_q 'return totals (materials + labor)
        End If


    End Function

    Function Cal_est_shipping(device As String) As Double

        'returns the estimate shipping cost for the totals datagrid

        Cal_est_shipping = 0

        Select Case device
            Case "I/O"
                Cal_est_shipping = 2 * get_total_bytype("ADA-NEMA12-BOX24", ADA_set_list, "IO Panel")
            Case "Motor"
                Cal_est_shipping = get_total_bytype("ADA-NEMA12-BOX24", ADA_set_list, "Starter Panel") + get_total_bytype("ADA-NEMA4X-BOX24", ADA_set_list, "Starter Panel")
            Case "Scanners"
                Cal_est_shipping = get_total_bytype("SICK Bar Code Scan 430 KIT", ADA_set_list, "Scanners") + get_total_bytype("SICK Bar Code Scan 490 OMNI KIT", ADA_set_list, "Scanners") +
                     get_total_bytype("Cognex Bar Code Scan 430 KIT", ADA_set_list, "Scanners")
        End Select


    End Function

    Private Sub SearchADAPartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchADAPartToolStripMenuItem.Click
        mydatagrid = 1
        Search_grid.ShowDialog()
    End Sub

    Private Sub ADA_Setup_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.F AndAlso e.Modifiers = Keys.Control) Then
            mydatagrid = 1
            Search_grid.ShowDialog()
        End If
    End Sub

    Private Sub ADASetsAndInstallationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADASetsAndInstallationToolStripMenuItem.Click
        'Export all data from all sets to an excel doc
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        '  xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")

                xlWorkSheet.Cells(1, 9) = ADA_set_list.count  'number of ada sets

                'Global settings

                xlWorkSheet.Cells(1, 1) = Voltage
                xlWorkSheet.Cells(2, 1) = efficiency
                xlWorkSheet.Cells(3, 1) = Power_factor
                xlWorkSheet.Cells(4, 1) = include_ADA_brackets
                xlWorkSheet.Cells(5, 1) = Use_4pts_IO
                xlWorkSheet.Cells(6, 1) = Use_8pts_IO
                xlWorkSheet.Cells(7, 1) = Use_16pts_IO
                xlWorkSheet.Cells(8, 1) = removal_msbox_t
                xlWorkSheet.Cells(9, 1) = M23_at_IO_Receptacle
                xlWorkSheet.Cells(10, 1) = M23_Bulkhead_at_ADA
                xlWorkSheet.Cells(11, 1) = Use_4pts_IO_RB
                xlWorkSheet.Cells(12, 1) = Use_8pts_IO_RB
                xlWorkSheet.Cells(13, 1) = Use_16pts_IO_RB_Inputs
                xlWorkSheet.Cells(14, 1) = Use_16pts_IO_RB_Outputs
                xlWorkSheet.Cells(15, 1) = Single_Channel_IO_1_Input_per_RB
                xlWorkSheet.Cells(16, 1) = Single_Channel_IO_1_Output_per_RB
                xlWorkSheet.Cells(17, 1) = Splitter_Perc_dual_ch
                xlWorkSheet.Cells(18, 1) = Valve_solenoid_adapter
                xlWorkSheet.Cells(19, 1) = Smart_Wire_Darwin_Limited
                xlWorkSheet.Cells(20, 1) = Smart_Wire_Darwin_Full
                xlWorkSheet.Cells(21, 1) = EthernetIP_Full
                xlWorkSheet.Cells(22, 1) = Percent_overage_inputs
                xlWorkSheet.Cells(23, 1) = Percent_overage_nm_outputs
                xlWorkSheet.Cells(24, 1) = Percent_overage_m_outputs

                Dim z As Integer : z = 0

                'textboxes
                For Each ada_set In ADA_set_list

                    xlWorkSheet.Cells(25 + z, 1) = "--"
                    xlWorkSheet.Cells(26 + z, 1) = ada_set.myName()
                    xlWorkSheet.Cells(27 + z, 1) = ada_set.panel_bool
                    xlWorkSheet.Cells(28 + z, 1) = ada_set.panel_bool
                    xlWorkSheet.Cells(29 + z, 1) = ada_set.panel_bool
                    xlWorkSheet.Cells(30 + z, 1) = ada_set.ADA_NEMA12_BOX24_No_Starters
                    xlWorkSheet.Cells(31 + z, 1) = ada_set.ADA_NEMA12_BOX30_No_Starters
                    xlWorkSheet.Cells(32 + z, 1) = ada_set.motor_disconnect
                    xlWorkSheet.Cells(33 + z, 1) = ada_set.MDRs
                    xlWorkSheet.Cells(34 + z, 1) = ada_set.NEMA12
                    xlWorkSheet.Cells(35 + z, 1) = ada_set.NEMA4X
                    xlWorkSheet.Cells(36 + z, 1) = ada_set.PLC_GateWay_PGW
                    xlWorkSheet.Cells(37 + z, 1) = ada_set.ControlLogix_PLC_Box
                    xlWorkSheet.Cells(38 + z, 1) = ada_set.CompactLogix_PLC_Box
                    xlWorkSheet.Cells(39 + z, 1) = ada_set.PLC_UPS_Battery_Backup
                    xlWorkSheet.Cells(40 + z, 1) = ada_set.Amp_Mon_Electro
                    xlWorkSheet.Cells(41 + z, 1) = ada_set.Remote_IO


                    For i = 0 To 47
                        For j = 0 To 3
                            xlWorkSheet.Cells(42 + i + z, j + 1) = ada_set.Motors_table(i, j)
                        Next
                    Next

                    '---------- input
                    For i = 0 To 18
                        For j = 0 To 3
                            xlWorkSheet.Cells(90 + i + z, j + 1) = ada_set.inputs_table(i, j)
                        Next
                    Next

                    '--------- push
                    For i = 0 To 18
                        For j = 0 To 3
                            xlWorkSheet.Cells(109 + i + z, j + 1) = ada_set.push_table(i, j)
                        Next
                    Next

                    '----------- lights
                    For i = 0 To 14
                        For j = 0 To 3
                            xlWorkSheet.Cells(124 + i + z, j + 1) = ada_set.lights_table(i, j)
                        Next
                    Next

                    '----------- field
                    For i = 0 To 11
                        For j = 0 To 3
                            xlWorkSheet.Cells(139 + i + z, j + 1) = ada_set.field_table(i, j)
                        Next
                    Next

                    '----------- bus
                    For i = 0 To 11
                        For j = 0 To 3
                            xlWorkSheet.Cells(151 + i + z, j + 1) = ada_set.bus_table(i, j)
                        Next
                    Next

                    '----------- scanner
                    For i = 0 To 3
                        For j = 0 To 3
                            xlWorkSheet.Cells(163 + i + z, j + 1) = ada_set.scanner_table(i, j)
                        Next
                    Next

                    '----------- brakes
                    For i = 0 To 14
                        For j = 0 To 3
                            xlWorkSheet.Cells(167 + i + z, j + 1) = ada_set.brakes_table(i, j)
                        Next
                    Next

                    z = z + 157
                Next


                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\AAPL_Setup.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Data Exported Succesfully!")
            End If
        End If



    End Sub

    Private Sub ImportSetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSetupToolStripMenuItem.Click
        ' /////////////////  Select ada file to import. These excel ada files contain all the data you need ////////////////


        Dim result As DialogResult = MessageBox.Show("Importing data will remove all your current sets. Would you like to proceed ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then
            ADA_set_list.clear()
            ComboBox1.Items.Clear()  'remove previous sets


            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            xlApp.DisplayAlerts = False

            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")

            Else

                If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

                    Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
                    '  Dim misValue As Object = System.Reflection.Missing.Value
                    '  wb = xlApp.Workbooks.Add(misValue)
                    '-- the code above block me from getting values from spreadsheet, weird


                    If String.IsNullOrEmpty(wb.Worksheets(1).Cells(1, 9).value) = False And IsNumeric(wb.Worksheets(1).Cells(1, 9).value) = True Then

                        If CType(wb.Worksheets(1).Cells(1, 9).value, Integer) > 0 Then


                            Dim n_sets As Integer : n_sets = CType(wb.Worksheets(1).Cells(1, 9).value, Integer)
                            Dim z As Integer : z = 0

                            If wb.Worksheets(1).Cells(1, 1).value = 480 Then
                                volts_480.Checked = True

                            ElseIf wb.Worksheets(1).Cells(1, 1).value = 575 Then
                                volts_575.Checked = True
                            Else
                                volts_230.Checked = True
                            End If


                            Efficiency_box.Text = wb.Worksheets(1).Cells(2, 1).value
                            Power_factor_box.Text = wb.Worksheets(1).Cells(3, 1).value
                            ADA_Brackets_box.Text = wb.Worksheets(1).Cells(4, 1).value
                            Use_4ptsIO_box.Text = wb.Worksheets(1).Cells(5, 1).value
                            Use_8ptsIO_box.Text = wb.Worksheets(1).Cells(6, 1).value
                            Use_16ptsIO_box.Text = wb.Worksheets(1).Cells(7, 1).value
                            removal_msbox.Text = wb.Worksheets(1).Cells(8, 1).value
                            M23_IO_box.Text = wb.Worksheets(1).Cells(9, 1).value
                            M23_bulkhead_box.Text = wb.Worksheets(1).Cells(10, 1).value
                            pts_4_rb_box.Text = wb.Worksheets(1).Cells(11, 1).value
                            pts_8_rb_box.Text = wb.Worksheets(1).Cells(12, 1).value
                            pts16_recep_inputs_box.Text = wb.Worksheets(1).Cells(13, 1).value
                            pts16_recep_outputs_box.Text = wb.Worksheets(1).Cells(14, 1).value
                            Single_input_box.Text = wb.Worksheets(1).Cells(15, 1).value
                            single_output_box.Text = wb.Worksheets(1).Cells(16, 1).value
                            splitter_per_box.Text = wb.Worksheets(1).Cells(17, 1).value
                            valve_box.Text = wb.Worksheets(1).Cells(18, 1).value
                            swd_limited_box.Text = wb.Worksheets(1).Cells(19, 1).value
                            swd_full_box.Text = wb.Worksheets(1).Cells(20, 1).value
                            ethernetip_box.Text = wb.Worksheets(1).Cells(21, 1).value
                            percentage_inputs_box.Text = wb.Worksheets(1).Cells(22, 1).value
                            percentage_nmoutputs_box.Text = wb.Worksheets(1).Cells(23, 1).value
                            percentage_moutputs_box.Text = wb.Worksheets(1).Cells(24, 1).value


                            For i = 0 To n_sets - 1

                                ProgressBar1.Visible = True

                                Dim Motor_array_l(47, 3) As String
                                Dim inputs_array_l(18, 3) As String
                                Dim push_array_l(18, 3) As String
                                Dim lights_array_l(14, 3) As String
                                Dim field_array_l(11, 3) As String
                                Dim bus_array_l(11, 3) As String
                                Dim scanner_array_l(3, 3) As String
                                Dim brakes_array_l(14, 3) As String

                                For h = 0 To 47
                                    For j = 0 To 3
                                        Motor_array_l(h, j) = wb.Worksheets(1).Cells(42 + h + z, j + 1).value
                                    Next
                                Next

                                '---------- input
                                For h = 0 To 18
                                    For j = 0 To 3
                                        inputs_array_l(h, j) = wb.Worksheets(1).Cells(90 + h + z, j + 1).value
                                    Next
                                Next

                                '--------- push
                                For h = 0 To 18
                                    For j = 0 To 3
                                        push_array_l(h, j) = wb.Worksheets(1).Cells(109 + h + z, j + 1).value
                                    Next
                                Next

                                '----------- lights
                                For h = 0 To 14
                                    For j = 0 To 3
                                        lights_array_l(h, j) = wb.Worksheets(1).Cells(124 + h + z, j + 1).value
                                    Next
                                Next

                                '----------- field
                                For h = 0 To 11
                                    For j = 0 To 3
                                        field_array_l(h, j) = wb.Worksheets(1).Cells(139 + h + z, j + 1).value
                                    Next
                                Next

                                '----------- bus
                                For h = 0 To 11
                                    For j = 0 To 3
                                        bus_array_l(h, j) = wb.Worksheets(1).Cells(151 + h + z, j + 1).value
                                    Next
                                Next

                                '----------- scanner
                                For h = 0 To 3
                                    For j = 0 To 3
                                        scanner_array_l(h, j) = wb.Worksheets(1).Cells(163 + h + z, j + 1).value
                                    Next
                                Next

                                '----------- brakes
                                For h = 0 To 14
                                    For j = 0 To 3
                                        brakes_array_l(h, j) = wb.Worksheets(1).Cells(167 + h + z, j + 1).value
                                    Next
                                Next


                                ADA_set_list.Add(New ADA_Set(wb.Worksheets(1).Cells(26 + z, 1).value, Motor_array_l, inputs_array_l, push_array_l, lights_array_l, field_array_l, bus_array_l, scanner_array_l, brakes_array_l))  'create set with constructor             

                                'start variables

                                ADA_set_list.Item(i).Panel_bool = wb.Worksheets(1).Cells(27 + z, 1).value
                                '   ADA_set_list.Item(i).R2_setup = wb.Worksheets(1).Cells(28 + z, 1).value
                                '  ADA_set_list.Item(i).Amp_30_Drops = wb.Worksheets(1).Cells(29 + z, 1).value
                                ADA_set_list.Item(i).ADA_NEMA12_BOX24_No_Starters = wb.Worksheets(1).Cells(30 + z, 1).value
                                ADA_set_list.Item(i).ADA_NEMA12_BOX30_No_Starters = wb.Worksheets(1).Cells(31 + z, 1).value
                                ADA_set_list.Item(i).motor_disconnect = wb.Worksheets(1).Cells(32 + z, 1).value
                                ADA_set_list.Item(i).MDRs = wb.Worksheets(1).Cells(33 + z, 1).value
                                ADA_set_list.Item(i).NEMA12 = wb.Worksheets(1).Cells(34 + z, 1).value
                                ADA_set_list.Item(i).NEMA4X = wb.Worksheets(1).Cells(35 + z, 1).value
                                ADA_set_list.Item(i).PLC_GateWay_PGW = wb.Worksheets(1).Cells(36 + z, 1).value
                                ADA_set_list.Item(i).ControlLogix_PLC_Box = wb.Worksheets(1).Cells(37 + z, 1).value
                                ADA_set_list.Item(i).CompactLogix_PLC_Box = wb.Worksheets(1).Cells(38 + z, 1).value
                                ADA_set_list.Item(i).PLC_UPS_Battery_Backup = wb.Worksheets(1).Cells(39 + z, 1).value
                                ADA_set_list.Item(i).Amp_Mon_Electro = wb.Worksheets(1).Cells(40 + z, 1).value
                                ADA_set_list.Item(i).Remote_IO = wb.Worksheets(1).Cells(41 + z, 1).value

                                Call get_global_set()  'update global variables

                                '--------- enter global settings values -------------
                                ADA_set_list.Item(i).Update_globals(Voltage, efficiency, Power_factor, include_ADA_brackets, Use_4pts_IO, Use_8pts_IO, Use_16pts_IO, removal_msbox_t, M23_at_IO_Receptacle,
                        M23_Bulkhead_at_ADA, Use_4pts_IO_RB, Use_8pts_IO_RB, Use_16pts_IO_RB_Inputs, Use_16pts_IO_RB_Outputs, Single_Channel_IO_1_Input_per_RB, Single_Channel_IO_1_Output_per_RB, Splitter_Perc_dual_ch, Valve_solenoid_adapter, Smart_Wire_Darwin_Limited,
                        Smart_Wire_Darwin_Full, EthernetIP_Full, Percent_overage_inputs, Percent_overage_nm_outputs, Percent_overage_m_outputs)
                                '---------------------------------------------------

                                ADA_set_list.Item(i).Calculate_Starter_Panel(Motor_horsepower, VFD_horsepower, Starters_values, value_to_index_480, value_to_index_575, Power_supply_480, Power_supply_575)
                                ADA_set_list.Item(i).Calculate_IO_Panel
                                ADA_set_list.Item(i).Calculate_PLC_Panel
                                ADA_set_list.Item(i).Calculate_Field_Panel
                                ADA_set_list.Item(i).Calculate_Scanners
                                ' ADA_set_list.Item(i).Calculate_M12
                                '  ADA_set_list.Item(i).Calculate_M12_ES


                                z = z + 157

                                If (ProgressBar1.Value < 400) Then
                                    ProgressBar1.Value = ProgressBar1.Value + 10
                                End If
                            Next

                            'clear textboxes


                            Call Load_set_names(ADA_set_list)  'load dropbox with ada set names
                            Call Cal_TOTAL_values() 'refresh TOTALS
                            temp_name = ""
                            ADA_name.Text = ""
                            ProgressBar1.Visible = False
                            ProgressBar1.Value = ProgressBar1.Minimum
                            MessageBox.Show("ADA Sets imported successfully!")

                        End If
                    End If

                    wb.Close(False)

                End If

            End If
        End If
    End Sub

    Function Max_number(num1 As Double, num2 As Double, num3 As Double, num4 As Double, num5 As Double) As Double

        'Return the max number of the parameters passed

        Max_number = num1

        Dim array_t(4) As Double
        array_t(0) = num1
        array_t(1) = num2
        array_t(2) = num3
        array_t(3) = num4
        array_t(4) = num5

        Max_number = array_t.Max

    End Function

    Private Sub ToolStripMenuItem30_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem30.Click
        'Remove a row in Purchase order
        If PR_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In PR_grid.SelectedRows
                PR_grid.Rows.Remove(r)
            Next
            Call Total_rows()
        Else
            MessageBox.Show("Select the row you want to delete")
        End If
    End Sub



    Function myinputs() As Double

        ''---------- Inputs calculation -----------------
        myinputs = 0

        For i = 0 To 47
            myinputs = myinputs + (CType(Motor_array(i, 1), Double) * CType(Motor_array(i, 0), Double))
        Next

        For i = 0 To 18
            myinputs = myinputs + (CType(inputs_array(i, 1), Double) * CType(inputs_array(i, 0), Double))
        Next

        For i = 0 To 18
            myinputs = myinputs + (CType(push_array(i, 1), Double) * CType(push_array(i, 0), Double))
        Next

        For i = 0 To 14
            myinputs = myinputs + (CType(lights_array(i, 1), Double) * CType(lights_array(i, 0), Double))
        Next

        For i = 0 To 11
            myinputs = myinputs + (CType(field_array(i, 1), Double) * CType(field_array(i, 0), Double))
        Next

        For i = 0 To 7  'check
            myinputs = myinputs + (CType(bus_array(i, 1), Double) * CType(bus_array(i, 0), Double))
        Next

        For i = 0 To 3
            myinputs = myinputs + (CType(scanner_array(i, 1), Double) * CType(scanner_array(i, 0), Double))
        Next

        For i = 0 To 14
            myinputs = myinputs + (CType(brakes_array(i, 1), Double) * CType(brakes_array(i, 0), Double))
        Next

        Dim per As Double : per = 0

        If IsNumeric(percentage_inputs_box.Text) = True Then
            per = Math.Ceiling((CType(percentage_inputs_box.Text, Double) / 100) * (myinputs - 8))
        End If

        myinputs = myinputs + per


    End Function

    Private Sub ToolStripMenuItem19_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem19.Click
        Quote_form.Visible = True
    End Sub

    '---------------------------------- Cables --------------------------------
    Sub M12_process()

        'calculate the number of M12 cables we need
        Dim remote_io As Double 'this control the loop to end in 15 or
        Dim limit_feet As Double : limit_feet = 15
        avoid_change_m12 = False

        Dim n_rows As Double : n_rows = 0
        Dim ran_feet As Integer : ran_feet = 3
        M12_grid.Rows.Clear()
        M12_ES_grid.Rows.Clear()

        For Each ada_set In ADA_set_list
            remote_io = ada_set.Remote_IO
            n_rows = n_rows + ada_set.cr_input + ada_set.cr_n_motion + ada_set.cr_motion
        Next

        If remote_io > 0 Then
            limit_feet = 15
        Else
            limit_feet = 30
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

        For Each ada_set In ADA_set_list
            count_rows = count_rows + CType(ada_set.inputs_table(1, 0), Double) + CType(ada_set.inputs_table(2, 0), Double) + CType(ada_set.inputs_table(5, 0), Double) + CType(ada_set.inputs_table(6, 0), Double)
        Next
        count_rows = count_rows + 1

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
                fixed_l2(1) = fixed_l2(1) + 1
                temp_l2(1) = 1
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

        M12_cables("ADA-SAC5-02-G") = fixed_l(0)
        M12_cables("ADA-SAC5-03-G") = fixed_l(1)
        M12_cables("ADA-SAC5-05-G") = fixed_l(2)
        M12_cables("ADA-SAC5-10-G") = fixed_l(3)
        M12_cables("ADA-SAC5-15-G") = fixed_l(4)
        M12_cables("ADA-SAC5-20-G") = fixed_l(5)

        M12_ES_cables("ADA-SAC5-02-Y") = fixed_l2(0)
        M12_ES_cables("ADA-SAC5-03-Y") = fixed_l2(1)
        M12_ES_cables("ADA-SAC5-05-Y") = fixed_l2(2)
        M12_ES_cables("ADA-SAC5-10-Y") = fixed_l2(3)
        M12_ES_cables("ADA-SAC5-15-Y") = fixed_l2(4)
        M12_ES_cables("ADA-SAC5-20-Y") = fixed_l2(5)

        avoid_change_m12 = True

    End Sub



    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        ''generate quote
        Dim appPath As String = Application.StartupPath()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkSheet As Excel.Worksheet
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            Try

                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\Template.xlsx")
                xlWorkSheet = wb.Sheets("TOTALS")

                '        'Start filling the vaules from TOTAL datagridview

                xlWorkSheet.Cells(4, 6) = totals_grid.Rows(0).Cells(1).Value()
                xlWorkSheet.Cells(5, 6) = totals_grid.Rows(1).Cells(1).Value()
                xlWorkSheet.Cells(6, 6) = totals_grid.Rows(2).Cells(1).Value()
                xlWorkSheet.Cells(7, 6) = totals_grid.Rows(3).Cells(1).Value()
                xlWorkSheet.Cells(8, 6) = totals_grid.Rows(4).Cells(1).Value()
                xlWorkSheet.Cells(9, 6) = totals_grid.Rows(5).Cells(1).Value()
                xlWorkSheet.Cells(22, 6) = totals_grid.Rows(7).Cells(1).Value()
                xlWorkSheet.Cells(23, 6) = totals_grid.Rows(8).Cells(1).Value()
                xlWorkSheet.Cells(24, 6) = totals_grid.Rows(9).Cells(1).Value()
                xlWorkSheet.Cells(25, 6) = totals_grid.Rows(10).Cells(1).Value()
                xlWorkSheet.Cells(26, 6) = totals_grid.Rows(11).Cells(1).Value()
                xlWorkSheet.Cells(27, 6) = totals_grid.Rows(12).Cells(1).Value()
                xlWorkSheet.Cells(28, 6) = totals_grid.Rows(13).Cells(1).Value()
                xlWorkSheet.Cells(29, 6) = totals_grid.Rows(14).Cells(1).Value()

                '        'installation
                xlWorkSheet.Cells(82, 6) = totals_grid.Rows(20).Cells(1).Value()
                xlWorkSheet.Cells(83, 6) = totals_grid.Rows(21).Cells(1).Value()
                xlWorkSheet.Cells(84, 6) = totals_grid.Rows(22).Cells(1).Value()
                xlWorkSheet.Cells(85, 6) = totals_grid.Rows(23).Cells(1).Value()

                '        'est shipping
                xlWorkSheet.Cells(89, 3) = totals_grid.Rows(16).Cells(1).Value()
                xlWorkSheet.Cells(90, 3) = totals_grid.Rows(17).Cells(1).Value()
                xlWorkSheet.Cells(91, 3) = totals_grid.Rows(18).Cells(1).Value()


                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then

                    wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)

                End If

                wb.Close(False)


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Quote generated successfully!")


            Catch ex As Exception
                MessageBox.Show("File not found or corrupted")
            End Try

        End If
    End Sub

    Private Sub SpreasheetViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpreasheetViewToolStripMenuItem.Click
        Setup_sheet.Visible = True
    End Sub

    Private Sub M12_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles M12_grid.CellValueChanged

        If avoid_change_m12 = True Then

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

                If (IsNumeric(M12_grid.Rows(i).Cells(0).Value)) = True Then

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
                End If
            Next

            For j = 0 To 5
                fixed_l(j) = Math.Ceiling(fixed_l(j) * 1.1)
            Next

            If M12_cables.ContainsKey("ADA-SAC5-02-G") Then
                M12_cables("ADA-SAC5-02-G") = fixed_l(0)
            End If

            If M12_cables.ContainsKey("ADA-SAC5-03-G") Then
                M12_cables("ADA-SAC5-03-G") = fixed_l(1)
            End If

            If M12_cables.ContainsKey("ADA-SAC5-05-G") Then
                M12_cables("ADA-SAC5-05-G") = fixed_l(2)
            End If

            If M12_cables.ContainsKey("ADA-SAC5-10-G") Then
                M12_cables("ADA-SAC5-10-G") = fixed_l(3)
            End If

            If M12_cables.ContainsKey("ADA-SAC5-15-G") Then
                M12_cables("ADA-SAC5-15-G") = fixed_l(4)
            End If

            If M12_cables.ContainsKey("ADA-SAC5-20-G") Then
                M12_cables("ADA-SAC5-20-G") = fixed_l(5)
            End If

            Call Cal_TOTAL_values() 'refresh TOTALS
            Call gen_PR()
            Call Cal_PartCost("Starter Panel")
        End If
    End Sub

    Private Sub M12_ES_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles M12_ES_grid.CellValueChanged

        If avoid_change_m12 = True Then

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
                temp_l2(5) = Math.Floor(M12_ES_grid.Rows(i).Cells(0).Value / 65.6)

                temp2 = M12_ES_grid.Rows(i).Cells(0).Value Mod 65.6

                If temp2 <= 6.56 Then
                    fixed_l2(0) = fixed_l2(0) + 1
                    temp_l2(0) = 1
                End If

                If temp2 <= 9.84 And temp2 > 6.56 Then
                    fixed_l2(1) = fixed_l2(1) + 1
                    temp_l2(1) = 1
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

            If M12_ES_cables.ContainsKey("ADA-SAC5-02-Y") Then
                M12_ES_cables("ADA-SAC5-02-Y") = fixed_l2(0)
            End If

            If M12_ES_cables.ContainsKey("ADA-SAC5-03-Y") Then
                M12_ES_cables("ADA-SAC5-03-Y") = fixed_l2(1)
            End If

            If M12_ES_cables.ContainsKey("ADA-SAC5-05-Y") Then
                M12_ES_cables("ADA-SAC5-05-Y") = fixed_l2(2)
            End If

            If M12_ES_cables.ContainsKey("ADA-SAC5-10-Y") Then
                M12_ES_cables("ADA-SAC5-10-Y") = fixed_l2(3)
            End If

            If M12_ES_cables.ContainsKey("ADA-SAC5-15-Y") Then
                M12_ES_cables("ADA-SAC5-15-Y") = fixed_l2(4)
            End If

            If M12_ES_cables.ContainsKey("ADA-SAC5-20-Y") Then
                M12_ES_cables("ADA-SAC5-20-Y") = fixed_l2(5)
            End If

            Call Cal_TOTAL_values() 'refresh TOTALS
            Call gen_PR()
            Call Cal_PartCost("Starter Panel")

        End If
    End Sub

    Sub gen_PR()
        '------------- GENERATE PURCHASE REQUEST FORM --------------

        PR_grid.Rows.Clear()  'Clear the PR Grid
        Dim count As Integer : count = 0
        Dim total_q As Double : total_q = 0

        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT ADA_number, Part_Name, Part_Description, manufacturer, Type, Primary_Vendor, KIT_DEVICE from ada_active_parts"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read

                    If (get_total_Qty(reader(0), ADA_set_list) > 0) Then

                        PR_grid.Rows.Add(New DataGridViewRow)
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(0).Value = reader(0)  'ADA Number                        
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(1).Value = reader(1)  'part name
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(2).Value = reader(2)  'Part description
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(3).Value = reader(3)  'manufacturer
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(4).Value = reader(4)  'Type
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(5).Value = reader(5)  'vendor             
                        PR_grid.Rows(PR_grid.Rows.Count - 2).Cells(7).Value = get_total_Qty(reader(0), ADA_set_list) : total_q = total_q + get_total_Qty(reader(0), ADA_set_list)
                        count = count + 1

                        If (String.Equals(reader(6), "KIT") = True) Then

                            PR_grid.Rows(PR_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                        End If

                    End If
                End While
            End If

            reader.Close()

            For Each row As DataGridViewRow In PR_grid.Rows
                If row.IsNewRow Then Continue For
                row.Cells(6).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(5).Value)
                row.Cells(8).Value = row.Cells(6).Value * row.Cells(7).Value
                row.Height = 52
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateMasterPackingListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateMasterPackingListToolStripMenuItem.Click
        If ADA_set_list.count > 0 Then

            '---- my assemblies ---
            Dim EPC101 As Double : EPC101 = 0
            Dim EPC102 As Double : EPC102 = 0
            Dim EPB101 As Double : EPB101 = 0
            Dim EPB102 As Double : EPB102 = 0
            Dim A2CS101 As Double : A2CS101 = 0
            Dim A3CS101 As Double : A3CS101 = 0
            Dim A3CS102 As Double : A3CS102 = 0
            Dim A3CS103 As Double : A3CS103 = 0
            Dim A3CS104 As Double : A3CS104 = 0
            Dim A3CS105 As Double : A3CS105 = 0
            Dim A1CS102 As Double : A1CS102 = 0
            Dim A1CS103 As Double : A1CS103 = 0
            Dim A1CS104 As Double : A1CS104 = 0
            Dim A1CS101 As Double : A1CS101 = 0



            Dim count_add_rows As Integer : count_add_rows = 8
            Dim appPath As String = Application.StartupPath()

            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            Dim xlWorkSheet As Excel.Worksheet
            xlApp.DisplayAlerts = False

            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
            Else
                Try
                    ProgressBar1.Visible = True
                    Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\1XXXXX ADA Packing List r2.0.xlsm")
                    xlWorkSheet = wb.Sheets("Master Packing List")

                    For Each ada_set In ADA_set_list

                        EPC101 = EPC101 + CType(ada_set.inputs_table(1, 0), Double)
                        EPC102 = EPC102 + CType(ada_set.inputs_table(2, 0), Double)
                        EPB101 = EPB101 + CType(ada_set.push_table(5, 0), Double)
                        EPB102 = EPB102 + CType(ada_set.push_table(6, 0), Double)
                        A2CS101 = A2CS101 + CType(ada_set.push_table(0, 0), Double)
                        A3CS101 = A3CS101 + CType(ada_set.push_table(1, 0), Double)
                        A3CS102 = A3CS102 + CType(ada_set.push_table(2, 0), Double)
                        A3CS103 = A3CS103 + CType(ada_set.push_table(3, 0), Double)
                        A3CS104 = A3CS104 + CType(ada_set.push_table(4, 0), Double)
                        A3CS105 = A3CS105 + CType(ada_set.push_table(5, 0), Double)
                        A1CS102 = A1CS102 + CType(ada_set.push_table(6, 0), Double)
                        A1CS103 = A1CS103 + CType(ada_set.push_table(7, 0), Double)
                        A1CS104 = A1CS104 + CType(ada_set.push_table(8, 0), Double)
                        A1CS101 = A1CS101 + CType(ada_set.push_table(9, 0), Double)

                    Next




                    If EPC101 > 0 Then
                        ' xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = EPC101 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-EPC101" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "Estop Pull Cord Single" ' description
                        count_add_rows = count_add_rows + 1
                    End If


                    If EPC102 > 0 Then
                        '  xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = EPC102 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-EPC102" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "Estop Pull Cord Double" ' description
                        count_add_rows = count_add_rows + 1
                    End If


                    If EPB101 > 0 Then
                        '  xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = EPB101 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-EPB101" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "E-Stop PB Station" ' description
                        count_add_rows = count_add_rows + 1
                    End If


                    If EPB102 > 0 Then
                        '  xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = EPB102 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-EPB102" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "E-Stop PB Station Guarded" ' description
                        count_add_rows = count_add_rows + 1
                    End If


                    If A2CS101 > 0 Then
                        '  xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A2CS101 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-2CS101" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "START / STOP (RD/GN)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A3CS101 > 0 Then
                        '    xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A3CS101 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-3CS101" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "3PBX-RED-GRN-BLU-2NO-1NC" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A3CS102 > 0 Then
                        ' xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A3CS102 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-3CS102" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "STOP / START / RESET (RD/GN/BU)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A3CS103 > 0 Then
                        '  xlWorkSheet.Rows(count_add_rows).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A3CS103 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-3CS103" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "STOP / START / AUTO-MAN (RD/GN/2S)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A3CS104 > 0 Then
                        '   xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A3CS104 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-3CS104" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "RESET / AUTO-MAN / HOME (BU/2S/BK)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A3CS105 > 0 Then
                        '  xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A3CS105 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-3CS105" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "HOME / AUTO-MAN / LEFT-RIGHT (BK/2S/3S)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A1CS102 > 0 Then
                        '  xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A1CS102 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-1CS102" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "RESET (BU)" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A1CS103 > 0 Then
                        '  xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A1CS103 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-1CS103" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "ENABLE (BU W/ LT)" ' description
                        count_add_rows = count_add_rows + 1
                    End If



                    If A1CS104 > 0 Then
                        '  xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A1CS104 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-1CS104" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "1PBX-BLU-NO-ES" ' description
                        count_add_rows = count_add_rows + 1
                    End If

                    If A1CS101 > 0 Then
                        'xlWorkSheet.Rows(9).Insert()
                        xlWorkSheet.Cells(count_add_rows, 1) = A1CS101 ' qty
                        xlWorkSheet.Cells(count_add_rows, 3) = "Atronix Panel Shop" ' supplier
                        xlWorkSheet.Cells(count_add_rows, 4) = "ADA-ASM-1CS101" 'part name
                        xlWorkSheet.Cells(count_add_rows, 5) = "1PBX-BLK-NO" ' description
                        count_add_rows = count_add_rows + 1
                    End If


                    xlWorkSheet.Range("7:7").Copy(xlWorkSheet.Range(count_add_rows & ":" & count_add_rows))
                    xlWorkSheet.Cells(count_add_rows, 3) = "Field"
                    count_add_rows = count_add_rows + 1
                    '  xlWorkSheet.Range("8:8").Copy(xlWorkSheet.Range(count_add_rows & ":" & count_add_rows))

                    For i = 0 To PR_grid.Rows.Count - 1

                        If (ProgressBar1.Value < 400) Then
                            ProgressBar1.Value = ProgressBar1.Value + 10
                        End If

                        If PR_grid.Rows(i).Cells(7).Value > 0 And String.Equals(PR_grid.Rows(i).Cells(4).Value, "Field") = True Then
                            '  xlWorkSheet.Rows(count_add_rows).Insert()
                            xlWorkSheet.Cells(count_add_rows, 1) = PR_grid.Rows(i).Cells(7).Value ' qty
                            xlWorkSheet.Cells(count_add_rows, 3) = PR_grid.Rows(i).Cells(5).Value ' supplier
                            xlWorkSheet.Cells(count_add_rows, 4) = PR_grid.Rows(i).Cells(1).Value 'part name
                            xlWorkSheet.Cells(count_add_rows, 5) = PR_grid.Rows(i).Cells(2).Value ' description

                            count_add_rows = count_add_rows + 1
                        End If


                    Next

                    xlWorkSheet.Range("D:D").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWorkSheet.Range("E:E").HorizontalAlignment = Excel.Constants.xlCenter

                    SaveFileDialog1.Filter = "Excel Files|*.xlsm"

                    If SaveFileDialog1.ShowDialog = DialogResult.OK Then

                        wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)

                    End If

                    wb.Close(False)


                    Marshal.ReleaseComObject(xlApp)
                    ProgressBar1.Visible = False
                    ProgressBar1.Value = ProgressBar1.Minimum
                    MessageBox.Show("Master Packing List generated successfully!")


                Catch ex As Exception
                    MessageBox.Show("Master Packing List Template not found or corrupted")
                End Try

            End If



        End If
    End Sub
End Class